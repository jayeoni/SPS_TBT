"""
LLM integration for WTO SPS notification processing.
Supports Ollama (local, no API key) and Anthropic Claude (cloud).
"""
import json
import os
import re
import urllib.request
import urllib.error

MODEL_ANTHROPIC = 'claude-sonnet-4-6'
MODEL_OLLAMA_DEFAULT = 'qwen2.5:7b'
OLLAMA_BASE_URL = 'http://localhost:11434'

SYSTEM_PROMPT = """You are an expert assistant for the Korean Ministry of Agriculture, Food and Rural Affairs (농림축산식품부, MAFRA) processing WTO SPS (Sanitary and Phytosanitary) notifications.

Your tasks:
1. Translate notification content from English/Spanish/Portuguese into formal Korean government language (공문체)
2. Classify notifications per MAFRA internal manual rules
3. Recommend handling based on domestic relevance

Rules for Korean government style:
- Use 개조식 for 내용 and 주간보고 fields: sentences must end in ~음/함/됨/임 style (e.g., "…공표함.", "…규정됨.", "…해당함."). NEVER use ~습니다/~합니다/~입니다 endings.
- Include scientific names in 국문명(학명) format (e.g., 아보카도(Persea americana))
- Use standard institutional Korean terms, not casual translations
- 목적 field: use only the approved standard phrases, semicolons between multiples
- 주간보고: write as a single action phrase, like "브라질산 아보카도 식물체의 수입검역요건 개정"

Always output valid JSON only. No explanation text before or after the JSON."""


def _build_user_prompt(parsed: dict, export_items: str, terminology: dict) -> str:
    term_lines = '\n'.join(f'  {k} → {v}' for k, v in list(terminology.items())[:80])

    objectives_str = '; '.join(parsed.get('objectives_korean', [])) or '(확인 필요)'

    is_addendum = parsed.get('is_addendum', False)
    is_emergency = parsed.get('is_emergency', False)
    notif_type_str = '긴급' if is_emergency else ('추가(Addendum)' if is_addendum else '일반')

    addendum_info = ''
    if is_addendum and parsed.get('addendum'):
        add = parsed['addendum']
        addendum_info = f"""
ADDENDUM INFO:
  Concerns: {', '.join(add.get('addendum_concerns', []))}
  Country: {add.get('addendum_country', '')}
  Received: {add.get('addendum_received_date', '')}"""

    export_section = (
        f'Korean exports found: {export_items}'
        if export_items and export_items != '-'
        else 'No Korean exports found for this country/product combination (write "-" for 국내수출품목).'
    )

    return f"""Process this WTO SPS notification:

DOCUMENT: {parsed.get('doc_number', '')}
TYPE: {notif_type_str}
NOTIFYING COUNTRY: {parsed.get('notifying_member', '')}
AGENCY RESPONSIBLE: {parsed.get('agency', '')}
SOURCE LANGUAGE: {parsed.get('source_language', 'en')}{addendum_info}

--- EXTRACTED FIELDS ---
Title: {parsed.get('title', '')}
Products covered: {parsed.get('products', '')}
Regions/countries affected: {parsed.get('regions', '')}
Objectives (checked): {objectives_str}
Description: {parsed.get('description', '')}
Other relevant documents: {parsed.get('other_docs', '')}
Comment deadline (raw): {parsed.get('comment_deadline_raw', '')}
Entry into force (raw): {parsed.get('entry_force_raw', '')}

--- DOMESTIC EXPORT DATA ---
{export_section}

--- TERMINOLOGY DICTIONARY (use these translations) ---
{term_lines}

--- CLASSIFICATION RULES ---

[중요도]
검토: (1) target is Korea/모든 교역국 AND Korea has exports of this product, OR (2) MRL stricter than Korean domestic standard, OR (3) sensitive issues: electronic phytosanitary cert, GMO, BSE, beef plant registration, customs tightening
참고: (1) from 24 export quarantine agreement countries with all-partners scope, OR (2) MRL same/weaker/absent vs domestic, OR (3) minor but relevant for quarantine practitioners
- (dash): (1) other ministry jurisdiction (MFDS/해수부/환경부), OR (2) third-country restriction not involving Korea, OR (3) no domestic exports, no quarantine negotiations
Note: If export_items is not "-", lean toward 검토 or 참고 depending on scope.
The 24 export agreement countries include: USA, Japan, EU, China, Australia, Canada, New Zealand, Philippines, Vietnam, Taiwan, Thailand, Singapore, Indonesia, Malaysia, Hong Kong, UAE, Russia, Kazakhstan, Mexico, Chile, Peru, Colombia, India, Saudi Arabia.

[구분]
식물: plant quarantine, plant pest regulations, seeds/planting material, wood packaging, oilseed crops (excl. processed), mushrooms/ginseng, organic produce, insects/sericulture, plant fertilizers, plant GMO/LMO
동물: animal quarantine, veterinary drug MRL, livestock feed/feed additives, pet animals, wildlife/hunting trophies, antibiotic regulations, HPAI/ASF/FMD/BSE suspensions, animal GMO/LMO
식품: pesticide MRL (agricultural products), processed food standards, food additives, Codex standards, new food materials, aquatic/fisheries products, tobacco

--- NOTIFICATION CONTENT CATEGORIES (통보내용) ---
Select the single best matching category for '통보내용' output field:
식물검역 | 비료 | 동물검역 | 사료첨가제 | 침입외래종 | 농약 | 동물용의약품 | GMO/LMO |
농산물 | 축산물 | 사료 | 특용작물 | 친환경·유기농산물 | 식용곤충·양잠 | 팽이버섯 |
신소재식품 | 할랄식품 | 식품첨가물 | 미생물/가공식품/제조시설 | 수산물 | 물/살생물제품 | 담배

For '통보_세부', select the best sub-type within the chosen 통보내용 (leave empty if none applies):
식물검역: 식물, 종자, 목재, 식물성비료/농기계, 목재포장재, 병해충
동물검역: 동물, 축산물, 동물성비료, 야생동물, 수산물
사료첨가제: 가축, 반려동물
침입외래종: 동물, 식물체
농약: 농산물, 축산물, 사료, 천연식물보호제, 규정
동물용의약품: MRL, 항생제, 규정
GMO/LMO: 사료, 식물체, 종자, 식품
농산물: 품질, 중금속, 곰팡이독소
축산물: 위생·안전, 품질
수산물: 위생품질

--- 주간보고 EXAMPLES (match these styles) ---
벨기에산 번식용 옥수수(Zea mays) 종자의 수입검역요건 발효
아르헨티나산 벳지(Vicia villosa) 종자의 수입검역요건(안) 제정
미국산 번식용 아보카도(Persea americana) 구근의 수입검역요건 개정
캐나다산 양과 염소의 수입을 위한 위생요건 제정
HPAI 발생에 따른 아르헨티나산 가금 및 가금제품의 수입 일시중단(90일)
고병원성 조류인플루엔자(HPAI) 확산 방지를 위한 폴란드산 살아있는 가금 및 가금류 지육의 수입 또는 경유 일시중단 관련 재개요건 추가
HPAI 발생에 따른 프랑스 루아르아틀랑티크(Loire-Atlantique)산 가금육, 알류 및 그 제품의 일시 수입금지 해제
식품의 규격 및 기준의 제정 - 자색차(Purple tea)
식품의 규격 및 기준의 개정 - 참치 및 가다랑어 통조림
캐나다 규제병해충 목록 개정 - 일부 병해충 삭제
신선 식용 블루베리(Vaccinium spp.) 수입 가능국가 추가-칠레, 멕시코, 모로코, 페루, 미국
개·고양이·수생생물 외 사료첨가제 재허가

--- OUTPUT FORMAT ---
Return ONLY this JSON object (no other text):
{{
  "제목": "Full verbatim Korean translation of the title; include scientific name as 국문명(학명) if present",
  "내용": "Full Korean translation of description in 개조식 (sentence endings: ~음/함/됨/임, never ~습니다/~합니다); translate entire text faithfully, do not summarize; use \\n between sentences",
  "해당품목": "Korean product name; keep scientific name in parentheses e.g., 아보카도(Persea americana)",
  "기타문서": "Korean translation of 'Other relevant documents'; omit URLs; translate language notes (e.g., 'available in Spanish' → '스페인어로 이용가능'); empty string if none",
  "목적": "ONLY these exact phrases, semicolons between multiples: 식품안전/동물위생/식물보호/동식물 해충·질병으로부터 사람 보호/해충으로 인한 피해로부터의 영토 보호",
  "해당국가": "Korean country name or '모든 교역국'",
  "통보국_kr": "Korean name of the notifying member country",
  "담당기관_kr": "Korean name of the agency; keep acronym in parentheses e.g. 동식물위생관리규제청(AGROCALIDAD)",
  "주간보고": "Single 개조식 Korean action line — follow the 주간보고 EXAMPLES patterns above",
  "구분": "동물 or 식물 or 식품",
  "구분_reason": "1-sentence reasoning",
  "중요도": "검토 or 참고 or -",
  "중요도_reason": "1-sentence reasoning citing specific rule",
  "관련부서": "Department 1\\nDepartment 2\\n(one per line)",
  "통보내용": "one value from the 통보내용 list above",
  "통보_세부": "one sub-type from the list above, or empty string",
  "품목": "Short Korean product label (e.g., 옥수수(Zea mays) 종자 or 가금 및 가금제품)",
  "flags": ["list of field names that are uncertain or need review"],
  "source_language": "en or es or pt"
}}"""


def _parse_llm_response(raw: str) -> dict:
    """Extract and parse JSON from LLM response, handling markdown fences."""
    # Strip markdown code fences that some models add
    raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.MULTILINE)
    raw = re.sub(r'\s*```\s*$', '', raw, flags=re.MULTILINE)
    raw = raw.strip()

    if raw.startswith('{'):
        json_str = raw
    else:
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if not m:
            raise ValueError(f'LLM 응답에서 JSON을 찾을 수 없습니다: {raw[:300]}')
        json_str = m.group()

    return json.loads(json_str)


def _process_with_anthropic(parsed: dict, export_items: str, terminology: dict, api_key: str) -> dict:
    import anthropic
    key = api_key or os.environ.get('ANTHROPIC_API_KEY', '')
    if not key:
        raise ValueError('ANTHROPIC_API_KEY가 설정되지 않았습니다.')
    client = anthropic.Anthropic(api_key=key)
    user_prompt = _build_user_prompt(parsed, export_items, terminology)
    message = client.messages.create(
        model=MODEL_ANTHROPIC,
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=[{'role': 'user', 'content': user_prompt}],
    )
    raw = message.content[0].text.strip()
    return _parse_llm_response(raw)


def _process_with_ollama(parsed: dict, export_items: str, terminology: dict, model: str) -> dict:
    user_prompt = _build_user_prompt(parsed, export_items, terminology)
    payload = json.dumps({
        'model': model,
        'messages': [
            {'role': 'system', 'content': SYSTEM_PROMPT},
            {'role': 'user', 'content': user_prompt},
        ],
        'stream': False,
        'options': {'temperature': 0.1, 'num_predict': 4096},
    }).encode('utf-8')

    for attempt in range(2):  # retry once on timeout (model cold-start)
        req = urllib.request.Request(
            f'{OLLAMA_BASE_URL}/api/chat',
            data=payload,
            method='POST',
            headers={'Content-Type': 'application/json'},
        )
        try:
            with urllib.request.urlopen(req, timeout=600) as resp:
                data = json.loads(resp.read())
                raw = data['message']['content'].strip()
            return _parse_llm_response(raw)
        except TimeoutError:
            if attempt == 0:
                continue  # model may still be loading; retry once
            raise ValueError(
                'Ollama 응답 시간 초과 (2회 시도).\n'
                'Ollama가 실행 중인지, 모델이 정상적으로 로드되었는지 확인하세요.'
            )
        except urllib.error.URLError as e:
            msg = str(e).lower()
            if 'connection refused' in msg or 'connect' in msg:
                raise ValueError(
                    'Ollama에 연결할 수 없습니다.\n'
                    '1. https://ollama.com 에서 Ollama를 설치하세요.\n'
                    '2. 터미널에서 실행: ollama serve\n'
                    f'3. 모델 다운로드: ollama pull {model}'
                )
            raise ValueError(f'Ollama 오류: {e}')
        except Exception as e:
            resp_text = ''
            if hasattr(e, 'read'):
                try:
                    resp_text = e.read().decode('utf-8', errors='replace')
                except Exception:
                    pass
            if 'model' in resp_text.lower() and 'not found' in resp_text.lower():
                raise ValueError(
                    f'Ollama 모델 "{model}"을 찾을 수 없습니다.\n'
                    f'설치 명령: ollama pull {model}'
                )
            raise ValueError(f'Ollama 처리 오류: {e}')


def check_ollama_status(model: str = MODEL_OLLAMA_DEFAULT) -> dict:
    """Check if Ollama is running and the model is available. Returns status dict."""
    try:
        req = urllib.request.Request(f'{OLLAMA_BASE_URL}/api/tags', method='GET')
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read())
        models = [m['name'].split(':')[0] for m in data.get('models', [])]
        model_base = model.split(':')[0]
        return {
            'running': True,
            'model_available': model_base in models,
            'available_models': models,
        }
    except Exception:
        return {'running': False, 'model_available': False, 'available_models': []}


def process_notification(
    parsed: dict,
    export_items: str,
    terminology: dict,
    api_key: str = None,
    llm_backend: str = 'ollama',
    ollama_model: str = MODEL_OLLAMA_DEFAULT,
) -> dict:
    """
    Translate, classify, and summarize a parsed WTO SPS notification.

    llm_backend: 'ollama' (local, no key) or 'anthropic' (cloud, needs key)
    """
    if llm_backend == 'anthropic':
        return _process_with_anthropic(parsed, export_items, terminology, api_key)
    else:
        return _process_with_ollama(parsed, export_items, terminology, ollama_model)
