"""
SPS Notification Processing Tool — Flask web application.
Run via start.bat; opens at http://localhost:5000
"""
import sys
from pathlib import Path as _Path
sys.path.insert(0, str(_Path(__file__).parent))

import os
import json
import logging
import traceback
from pathlib import Path

from flask import Flask, request, jsonify, render_template
from dotenv import load_dotenv

# Load .env from the tool's own directory
BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / '.env')

import parser as sps_parser
import llm as sps_llm
import word_writer
import dept_lookup

# ── App setup ────────────────────────────────────────────────────────────────
app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format='%(levelname)s %(message)s')
log = logging.getLogger(__name__)

# ── Config ───────────────────────────────────────────────────────────────────
CONFIG_FILE = BASE_DIR / 'config.json'

DEFAULT_CONFIG = {
    'output_dir':   os.environ.get('OUTPUT_DIR', ''),
    'api_key':      os.environ.get('ANTHROPIC_API_KEY', ''),
    'llm_backend':  'ollama',  # 'ollama' (local, free) or 'anthropic'
    'ollama_model': 'qwen2.5:7b',
}


def load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, encoding='utf-8') as f:
            saved = json.load(f)
        cfg = {**DEFAULT_CONFIG, **saved}
    else:
        cfg = dict(DEFAULT_CONFIG)
    # ENV vars always override saved config for the API key
    if os.environ.get('ANTHROPIC_API_KEY'):
        cfg['api_key'] = os.environ['ANTHROPIC_API_KEY']
    return cfg


def save_config(cfg: dict):
    safe = {k: v for k, v in cfg.items() if k != 'api_key'}  # don't save key to JSON
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(safe, f, ensure_ascii=False, indent=2)


# ── Terminology ──────────────────────────────────────────────────────────────
TERMINOLOGY_FILE = BASE_DIR / 'terminology.json'
_terminology: dict | None = None


def load_terminology() -> dict:
    global _terminology
    if _terminology is None:
        if TERMINOLOGY_FILE.exists():
            with open(TERMINOLOGY_FILE, encoding='utf-8') as f:
                _terminology = json.load(f)
        else:
            _terminology = {}
    return _terminology



# ── Core processing pipeline ──────────────────────────────────────────────────
def process_single_file(docx_path: str, cfg: dict, terminology: dict | None = None) -> dict:
    """
    Full processing pipeline for one WTO SPS notification file.
    Returns a result dict for display in the UI.
    """
    result = {
        'filename':          Path(docx_path).name,
        'doc_number':        '',
        'notifying_country': '',
        'title_kr':          '',
        'type':              '',
        'success':           False,
        'error':             None,
        'word_file':         '',
        'importance':        '',
        'category':          '',
    }

    try:
        # ── 1. Parse ──────────────────────────────────────────────────────
        log.info('[%s] 파싱 중...', result['filename'])
        parsed = sps_parser.parse_notification(docx_path)
        result['doc_number'] = parsed.get('doc_number', '')
        result['type'] = '긴급' if parsed['is_emergency'] else '일반'

        if not result['doc_number']:
            result['error'] = '문서번호를 찾을 수 없습니다. 파일을 확인해주세요.'
            return result

        # ── 2. Regions ────────────────────────────────────────────────────
        regions_raw = parsed.get('regions', '')
        regions_kr = dept_lookup.translate_regions(regions_raw)
        parsed['regions_kr'] = regions_kr
        is_all_partners = regions_kr == '모든 교역국'

        # ── 3. LLM ───────────────────────────────────────────────────────
        log.info('[%s] LLM 처리 중 (번역 + 분류)...', result['filename'])
        if terminology is None:
            terminology = load_terminology()
        llm_result = sps_llm.process_notification(
            parsed=parsed,
            export_items='',
            terminology=terminology,
            api_key=cfg.get('api_key', ''),
            llm_backend=cfg.get('llm_backend', 'ollama'),
            ollama_model=cfg.get('ollama_model', 'qwen2.5:7b'),
        )

        result['title_kr']          = llm_result.get('제목', '')
        result['category']          = llm_result.get('구분', '')
        result['notifying_country'] = parsed.get('notifying_member', '')

        _notifying_kr = dept_lookup.translate_regions(parsed.get('notifying_member', ''))
        if _notifying_kr:
            llm_result['통보국_kr'] = _notifying_kr

        importance = llm_result.get('중요도', '')
        if not is_all_partners and '한국' not in regions_kr:
            importance = '-'
        result['importance'] = importance

        # ── 4. Create bilingual Word file ─────────────────────────────────
        log.info('[%s] 번역본 Word 파일 생성 중...', result['filename'])
        output_word = word_writer.create_bilingual_docx(
            source_path=docx_path,
            translations={
                **llm_result,
                '통보국_kr': llm_result.get('통보국_kr', ''),
                '해당국가':  regions_kr or llm_result.get('해당국가', ''),
            },
            is_addendum=parsed['is_addendum'],
        )
        result['word_file'] = Path(output_word).name

        result['success'] = True
        log.info('[%s] 완료 ✓', result['filename'])

    except Exception as e:
        log.error('[%s] 오류: %s', result['filename'], traceback.format_exc())
        result['error'] = str(e)

    return result


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    cfg = load_config()
    missing = []
    if cfg.get('llm_backend', 'ollama') == 'anthropic' and not cfg.get('api_key'):
        missing.append('ANTHROPIC_API_KEY')
    return render_template('index.html', config=cfg, missing=missing)


@app.route('/process', methods=['POST'])
def process():
    cfg = load_config()
    terminology = load_terminology()

    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '파일이 선택되지 않았습니다.'}), 400

    results = []
    for uploaded_file in files:
        if not uploaded_file.filename.endswith('.docx'):
            results.append({
                'filename': uploaded_file.filename,
                'success': False,
                'error': '.docx 파일만 처리할 수 있습니다.',
            })
            continue

        output_dir = Path(cfg.get('output_dir', '') or BASE_DIR)
        if not output_dir.exists():
            output_dir = BASE_DIR

        tmp_path = output_dir / uploaded_file.filename
        uploaded_file.save(str(tmp_path))

        result = process_single_file(str(tmp_path), cfg, terminology)
        results.append(result)

    return jsonify({'results': results})


@app.route('/settings', methods=['GET', 'POST'])
def settings():
    cfg = load_config()
    message = ''

    if request.method == 'POST':
        new_cfg = {
            'output_dir':   request.form.get('output_dir', '').strip(),
            'llm_backend':  request.form.get('llm_backend', 'ollama'),
            'ollama_model': request.form.get('ollama_model', 'qwen2.5:7b').strip(),
        }
        new_api_key = request.form.get('api_key', '').strip()

        if new_api_key:
            env_path = BASE_DIR / '.env'
            env_path.write_text(f'ANTHROPIC_API_KEY={new_api_key}\n', encoding='utf-8')
            load_dotenv(env_path, override=True)
            cfg['api_key'] = new_api_key

        cfg.update(new_cfg)
        save_config(cfg)
        message = '설정이 저장되었습니다.'

    return render_template('settings.html', config=cfg, message=message)


@app.route('/health')
def health():
    cfg = load_config()
    return jsonify({
        'api_key_set':  bool(cfg.get('api_key')),
        'llm_backend':  cfg.get('llm_backend', 'ollama'),
        'ollama_model': cfg.get('ollama_model', 'qwen2.5:7b'),
    })


@app.route('/ollama-status')
def ollama_status():
    cfg = load_config()
    model = cfg.get('ollama_model', 'qwen2.5:7b')
    status = sps_llm.check_ollama_status(model)
    status['model'] = model  # include model name so frontend can display it
    return jsonify(status)


if __name__ == '__main__':
    import webbrowser
    import threading
    threading.Timer(1.5, lambda: webbrowser.open('http://localhost:5000')).start()
    app.run(host='127.0.0.1', port=5000, debug=False)
