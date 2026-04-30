"""
SPS Notification Processing Tool — Flask web application.
Run via start.bat; opens at http://localhost:5000
"""
import os
import json
import logging
import traceback
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, render_template, redirect, url_for
from dotenv import load_dotenv

# Load .env from the tool's own directory
BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / '.env')

import parser as sps_parser
import llm as sps_llm
import date_engine
import export_lookup as exp_lookup
import excel_writer
import word_writer

# ── App setup ────────────────────────────────────────────────────────────────
app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format='%(levelname)s %(message)s')
log = logging.getLogger(__name__)

# ── Config ───────────────────────────────────────────────────────────────────
CONFIG_FILE = BASE_DIR / 'config.json'

DEFAULT_CONFIG = {
    'excel_path':       os.environ.get('EXCEL_PATH', ''),
    'export_data_path': os.environ.get('EXPORT_DATA_PATH', ''),
    'api_key':          os.environ.get('ANTHROPIC_API_KEY', ''),
    'target_month':     '',      # e.g. '26.4월'; empty = auto-detect
    'llm_backend':      'ollama',  # 'ollama' (local, free) or 'anthropic'
    'ollama_model':     'qwen2.5:7b',
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


def load_terminology() -> dict:
    if TERMINOLOGY_FILE.exists():
        with open(TERMINOLOGY_FILE, encoding='utf-8') as f:
            return json.load(f)
    return {}


# ── Export lookup (loaded once at startup) ────────────────────────────────────
_export_lookup = exp_lookup.get_lookup()


def ensure_export_loaded(cfg: dict):
    path = cfg.get('export_data_path', '')
    if path and Path(path).exists() and not _export_lookup.is_loaded():
        try:
            _export_lookup.load(path)
            log.info('수출 실적 데이터 로드 완료: %s', path)
        except Exception as e:
            log.warning('수출 데이터 로드 실패: %s', e)


# ── Core processing pipeline ──────────────────────────────────────────────────
def process_single_file(docx_path: str, cfg: dict, terminology: dict | None = None) -> dict:
    """
    Full processing pipeline for one WTO SPS notification file.
    Returns a result dict for display in the UI.
    """
    result = {
        'filename':      Path(docx_path).name,
        'doc_number':    '',
        'notifying_country': '',
        'title_kr':      '',
        'type':          '',
        'success':       False,
        'error':         None,
        'excel_updated': False,
        'word_file':     '',
        'flags':         [],
        'importance':    '',
        'category':      '',
        'row_idx':       None,
        'skipped':       False,
    }

    try:
        # ── 1. Parse Word document ─────────────────────────────────────────
        log.info('[%s] 파싱 중...', result['filename'])
        parsed = sps_parser.parse_notification(docx_path)
        result['doc_number'] = parsed.get('doc_number', '')
        result['type'] = '긴급' if parsed['is_emergency'] else (
            '추가' if parsed['is_addendum'] else '일반'
        )

        if not result['doc_number']:
            result['error'] = '문서번호를 찾을 수 없습니다. 파일을 확인해주세요.'
            return result

        # ── 2. Get 배포일 from Excel (needed for date calculations) ─────────
        excel_path = cfg.get('excel_path', '')
        base_date = None
        if excel_path and Path(excel_path).exists():
            base_date = excel_writer.get_base_date(
                excel_path, result['doc_number'], cfg.get('target_month')
            )

        # ── 3. Export item lookup ──────────────────────────────────────────
        is_all_partners = '모든 교역국' in parsed.get('regions', '') or \
                          'all trading partners' in parsed.get('regions', '').lower()
        export_items, export_uncertain = _export_lookup.lookup(
            notifying_country=parsed.get('notifying_member', ''),
            product_text=parsed.get('products', '') + ' ' + parsed.get('description', ''),
            is_all_partners=is_all_partners,
            category='',
        ) if _export_lookup.is_loaded() else ('-', False)

        # ── 4. LLM processing ──────────────────────────────────────────────
        log.info('[%s] LLM 처리 중 (번역 + 분류)...', result['filename'])
        if terminology is None:
            terminology = load_terminology()
        llm_result = sps_llm.process_notification(
            parsed=parsed,
            export_items=export_items,
            terminology=terminology,
            api_key=cfg.get('api_key', ''),
            llm_backend=cfg.get('llm_backend', 'ollama'),
            ollama_model=cfg.get('ollama_model', 'qwen2.5:7b'),
        )

        result['title_kr']  = llm_result.get('제목', '')
        result['importance'] = llm_result.get('중요도', '')
        result['category']   = llm_result.get('구분', '')
        result['notifying_country'] = parsed.get('notifying_member', '')

        # ── 5. Date calculations ───────────────────────────────────────────
        date_fields = {}
        if base_date:
            if parsed.get('comment_deadline_raw'):
                resolved, _ = date_engine.resolve_date(
                    parsed['comment_deadline_raw'],
                    base_date,
                    is_emergency=parsed['is_emergency'],
                )
                date_fields['의견마감일'] = resolved
            if parsed.get('entry_force_raw'):
                resolved, _ = date_engine.resolve_date(
                    parsed['entry_force_raw'],
                    base_date,
                )
                date_fields['발효일'] = resolved
        # Emergency: comment deadline is always '-'
        if parsed['is_emergency']:
            date_fields['의견마감일'] = '-'

        # ── 6. Assemble all Excel fields ───────────────────────────────────
        is_non_english = llm_result.get('source_language', 'en') != 'en'

        all_fields = {
            '중요도':       llm_result.get('중요도', ''),
            '제목':         llm_result.get('제목', ''),
            '내용':         llm_result.get('내용', ''),
            '해당품목':     llm_result.get('해당품목', ''),
            '목적':         llm_result.get('목적', ''),
            '해당국가':     llm_result.get('해당국가', ''),
            '국내수출품목': export_items if export_items else '-',
            '관련부서':     llm_result.get('관련부서', ''),
            '주간보고':     llm_result.get('주간보고', ''),
            '구분':         llm_result.get('구분', ''),
            '품목':         llm_result.get('품목', ''),
            **date_fields,
        }

        # Collect uncertainty flags
        flags = list(llm_result.get('flags', []))
        if export_uncertain:
            flags.append('국내수출품목')
        if len(llm_result.get('관련부서', '').split('\n')) > 2:
            flags.append('관련부서')

        result['flags'] = flags

        # ── 7. Write to Excel ──────────────────────────────────────────────
        if excel_path and Path(excel_path).exists():
            ok, err, row_idx = excel_writer.load_and_process(
                excel_path=excel_path,
                doc_number=result['doc_number'],
                fields=all_fields,
                uncertain_fields=flags,
                is_non_english=is_non_english,
                target_month=cfg.get('target_month'),
            )
            if ok:
                result['excel_updated'] = True
                result['row_idx'] = row_idx
            else:
                # Not finding the row is a warning, not an error — still proceed
                if '찾을 수 없습니다' in err:
                    result['skipped'] = True
                    result['error'] = err
                else:
                    result['error'] = f'Excel 오류: {err}'
        else:
            result['error'] = 'Excel 파일 경로가 설정되지 않았습니다. 설정(Settings)을 확인해주세요.'

        # ── 8. Create bilingual Word file ──────────────────────────────────
        log.info('[%s] 번역본 Word 파일 생성 중...', result['filename'])
        output_word = word_writer.create_bilingual_docx(
            source_path=docx_path,
            translations={**llm_result, '통보국_kr': llm_result.get('통보국_kr', '')},
            is_non_english=is_non_english,
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
    ensure_export_loaded(cfg)
    missing = []
    if cfg.get('llm_backend', 'ollama') == 'anthropic' and not cfg.get('api_key'):
        missing.append('ANTHROPIC_API_KEY')
    if not cfg.get('excel_path') or not Path(cfg['excel_path']).exists():
        missing.append('Excel 파일 경로')
    return render_template('index.html', config=cfg, missing=missing)


@app.route('/process', methods=['POST'])
def process():
    cfg = load_config()
    ensure_export_loaded(cfg)
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

        # Save uploaded file to a temp location alongside the source
        # We save it to the same folder as the Excel file for now
        excel_dir = Path(cfg.get('excel_path', BASE_DIR)).parent
        if not excel_dir.exists():
            excel_dir = BASE_DIR

        tmp_path = excel_dir / uploaded_file.filename
        uploaded_file.save(str(tmp_path))

        result = process_single_file(str(tmp_path), cfg, terminology)
        results.append(result)

        # If the file was uploaded (not already there), we leave the original
        # and the _번역.docx alongside it.

    return jsonify({'results': results})


@app.route('/settings', methods=['GET', 'POST'])
def settings():
    cfg = load_config()
    message = ''

    if request.method == 'POST':
        new_cfg = {
            'excel_path':       request.form.get('excel_path', '').strip(),
            'export_data_path': request.form.get('export_data_path', '').strip(),
            'target_month':     request.form.get('target_month', '').strip(),
            'llm_backend':      request.form.get('llm_backend', 'ollama'),
            'ollama_model':     request.form.get('ollama_model', 'qwen2.5:7b').strip(),
        }
        new_api_key = request.form.get('api_key', '').strip()

        # Save API key to .env so it persists
        if new_api_key:
            env_path = BASE_DIR / '.env'
            env_content = f'ANTHROPIC_API_KEY={new_api_key}\n'
            env_content += f'EXCEL_PATH={new_cfg["excel_path"]}\n'
            env_content += f'EXPORT_DATA_PATH={new_cfg["export_data_path"]}\n'
            env_path.write_text(env_content, encoding='utf-8')
            load_dotenv(env_path, override=True)
            cfg['api_key'] = new_api_key

        cfg.update(new_cfg)
        save_config(cfg)

        # Reload export data if path changed
        if new_cfg['export_data_path']:
            _export_lookup._df = None
            _export_lookup._path = None
            ensure_export_loaded(cfg)

        message = '설정이 저장되었습니다.'

    return render_template('settings.html', config=cfg, message=message)


@app.route('/health')
def health():
    cfg = load_config()
    return jsonify({
        'api_key_set':    bool(cfg.get('api_key')),
        'excel_exists':   bool(cfg.get('excel_path')) and Path(cfg['excel_path']).exists(),
        'export_loaded':  _export_lookup.is_loaded(),
        'llm_backend':    cfg.get('llm_backend', 'ollama'),
        'ollama_model':   cfg.get('ollama_model', 'qwen2.5:7b'),
    })


@app.route('/ollama-status')
def ollama_status():
    cfg = load_config()
    model = cfg.get('ollama_model', 'qwen2.5:7b')
    status = sps_llm.check_ollama_status(model)
    return jsonify(status)


if __name__ == '__main__':
    import webbrowser
    import threading
    cfg = load_config()
    ensure_export_loaded(cfg)
    threading.Timer(1.5, lambda: webbrowser.open('http://localhost:5000')).start()
    app.run(host='127.0.0.1', port=5000, debug=False)
