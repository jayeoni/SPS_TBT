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

# ── Pre-compiled regexes (module-level to avoid recompilation per file) ────────
_RE_LANG_PAGE = re.compile(
    r'^\s*(?:language|n[uú]mero de p[aá]ginas|number of pages)',
    re.IGNORECASE,
)
_RE_LOCAL_GOV = re.compile(
    r'(?:if applicable.*?involved|local government)[^\n]*:\s*([^\n]+)',
    re.IGNORECASE,
)
_RE_MD_FENCE  = re.compile(r'^```(?:json)?\s*', re.MULTILINE)
_RE_MD_CLOSE  = re.compile(r'\s*```\s*$', re.MULTILINE)
_RE_JSON_OBJ  = re.compile(r'\{.*\}', re.DOTALL)

# Behavioral rules only — content-specific rules belong in the user prompt
# where they are near the actual data.
SYSTEM_PROMPT = """You are an expert Korean government document analyst processing WTO SPS notifications for MAFRA (농림축산식품부).

Output ONLY valid JSON — no markdown fences, no explanation, no text outside the JSON object."""


def _select_terms(terminology: dict, doc_text: str, max_terms: int = 60) -> list:
    """
    Return up to max_terms (key, value) pairs, prioritising terms that appear
    in doc_text so that the most relevant translations are always included.
    """
    doc_lower = doc_text.lower()
    relevant, seen = [], set()
    for k, v in terminology.items():
        if k.lower() in doc_lower:
            relevant.append((k, v))
            seen.add(k)
    # Fill remaining slots from the full list (preserves insertion-order priority)
    filler = [(k, v) for k, v in terminology.items() if k not in seen]
    return (relevant + filler)[:max_terms]


def _build_user_prompt(parsed: dict, export_items: str, terminology: dict) -> str:
    # Combine document text for term-relevance scoring
    doc_text = ' '.join(filter(None, [
        parsed.get('title', ''),
        parsed.get('products', ''),
        parsed.get('description', ''),
        parsed.get('other_docs', ''),
        parsed.get('objective_text', ''),
    ]))
    selected_terms = _select_terms(terminology, doc_text)
    term_lines = '\n'.join(f'  {k} → {v}' for k, v in selected_terms)

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

    # Extract only the country name from notifying_member (first non-empty line).
    # The raw field often includes boilerplate like "If applicable, name of local government involved:"
    # which can confuse small LLMs into swapping 통보국_kr and 해당국가.
    notifying_raw = parsed.get('notifying_member', '')
    notifying_country = next(
        (ln.strip() for ln in notifying_raw.split('\n') if ln.strip()),
        '',
    )
    local_gov_m = _RE_LOCAL_GOV.search(notifying_raw)
    local_gov = local_gov_m.group(1).strip() if local_gov_m else ''
    local_gov_line = f'\nLOCAL GOVERNMENT (field 1, if any): {local_gov}' if local_gov else ''

    # Strip "Language(s):" and "Number of pages:" metadata lines from the title cell.
    title_raw = parsed.get('title', '')
    title_clean = '\n'.join(
        ln for ln in title_raw.split('\n')
        if not _RE_LANG_PAGE.match(ln)
    ).strip()

    # WTO documents use non-breaking spaces (U+00A0) in "Law\xa0No.\xa01020" — normalise.
    other_docs_clean = parsed.get('other_docs', '').replace('\xa0', ' ')

    return f"""Process this WTO SPS notification:

DOCUMENT: {parsed.get('doc_number', '')}
TYPE: {notif_type_str}
NOTIFYING COUNTRY: {notifying_country}{local_gov_line}
AGENCY RESPONSIBLE: {parsed.get('agency', '')}
SOURCE LANGUAGE: {parsed.get('source_language', 'en')}{addendum_info}

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

--- 원산지 표현 RULE ---
Translate origin phrases as '[Country_Korean]산':
  'originating in and coming from Chile' → '칠레산'
  'coming from Argentina' → '아르헨티나산'
  'procedente de Nicaragua' → '니카라과산'
  'en provenance de France' → '프랑스산'

--- 제목 TRANSLATION RULES ---
CRITICAL: Translate ONLY the actual document title text. Do NOT include language or page-count metadata in 제목.
Resolution number format: "Resolution No. X-Y-Z" or "Resolución No. X-Y-Z" → ALWAYS write as "결의안 제X-Y-Z호" (NEVER keep "No." in Korean output).
Product name: translate fully into Korean; keep scientific name in parentheses if present.
Structure for "Resolution establishing requirements for [product] originating in [country]":
  → [country_kr]산 [product_kr] 수입에 대한 [requirements type_kr]을 규정하는 결의안 제[number]호

--- 제목 TRANSLATION EXAMPLES ---
Resolution No. 1##-2026-IPSA establishing phytosanitary requirements for the importation of fresh pears (Pyrus communis) originating in Peru → 페루산 신선한 배(Pyrus communis) 수입에 대한 식물검역요건을 규정하는 결의안 제175-2026-IPSA호
Resolution No. 1##-2026-IPSA establishing phytosanitary requirements for the importation of unmanufactured tobacco (Nicotiana tabacum) originating in the United States → 미국산 담배(Nicotiana tabacum) 수입에 대한 식물검역요건을 규정하는 결의안 제159-2026-IPSA호
Resolución No. 045-2025 que establece los requisitos sanitarios para la importación de carne de res (Bos taurus) procedente de Argentina → 아르헨티나산 쇠고기(Bos taurus) 수입에 대한 위생요건을 규정하는 결의안 제045-2025호
Resolución No. 012-2024-MAG estableciendo requisitos fitosanitarios para importación de manzanas frescas (Malus domestica) originarias de Chile → 칠레산 신선 사과(Malus domestica) 수입에 대한 식물검역요건을 규정하는 결의안 제012-2024-MAG호

--- 기타문서 TRANSLATION RULES ---
"Law No. X" or "Ley No. X" → ALWAYS write as "법령 제X호" (NEVER keep "No." or "Ley" in Korean).
Format: 법령 제[number]호 "[Korean law title]", ([language] 이용 가능)
Output ONLY the document reference line(s) — NO commentary, NO explanatory sentences, NO preamble.
One line per document; multiple documents separated by \\n.

--- 기타문서 EXAMPLE ---
Law No. 1020, "Ley de Protección Fitosanitaria de Nicaragua" (available in Spanish) → 법령 제1020호 "니카라과 식물검역 및 보호법", (스페인어로 이용 가능)

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

--- TERMINOLOGY DICTIONARY (use these translations) ---
{term_lines}

--- EXTRACTED FIELDS ---
Title: {title_clean}
Products covered: {parsed.get('products', '')}
Regions/countries affected: {parsed.get('regions', '')}
Objectives (checked): {objectives_str}
Objective/rationale text: {parsed.get('objective_text', '')}
Description (→ translate into "내용"): {parsed.get('description', '')}
--- END OF DESCRIPTION — do NOT include anything below this line in "내용" ---
Other relevant documents (→ translate into "기타문서"): {other_docs_clean}
Comment deadline (raw): {parsed.get('comment_deadline_raw', '')}
Entry into force (raw): {parsed.get('entry_force_raw', '')}

--- DOMESTIC EXPORT DATA ---
{export_section}

--- OUTPUT FORMAT ---
IMPORTANT: "내용" must be a complete Korean 개조식 translation of the "Description" field above.
Translate the actual Description text word-for-word — NOT the title, NOT any example from this prompt.

Return ONLY this JSON object (no other text):
{{
  "지방정부_kr": "Korean name of local/regional government from field 1 (e.g. '캘리포니아 주'); empty string if none",
  "제목": "Full verbatim Korean translation of the title; include scientific name as 국문명(학명) if present",
  "내용": "Complete 개조식 Korean translation of EVERY clause, requirement, species name, date, and document number in Description. Endings: 됨/함/임/어야 함. Never ~습니다/~합니다. Use \\n between items.",
  "해당품목": "Korean translation of 'Products covered'. Keep scientific names in parentheses. Korean modifier-first order: qualifying phrases (disease risk, origin, conditions) come BEFORE the noun — e.g., 'pigs and genetic material, products and by-products of swine origin at risk of transmitting the Aujeszky\\'s disease virus' → '오제스키병 바이러스 전파 위험이 있는 돼지 및 돼지 유래 유전 물질, 제품 및 부산물'. Do NOT copy from 주간보고.",
  "기타문서": "Follow 기타문서 TRANSLATION RULES above. Reference lines only — no URLs, no commentary. Empty string if genuinely empty.",
  "목적": "ONLY objectives explicitly checked in 'Objectives (checked)'. Exact phrases, semicolons between: 식품안전/동물위생/식물보호/동식물 해충·질병으로부터 사람 보호/해충으로 인한 피해로부터의 영토 보호. Empty string if none checked.",
  "목적_근거": "개조식 Korean translation of ONLY the free-text rationale from 'Objective/rationale text' input. Do NOT include any content from 'Description'. Empty string if 'Objective/rationale text' is empty or contains only checkboxes.",
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
  "품목": "Short Korean product label (e.g. 옥수수(Zea mays) 종자 or 가금 및 가금제품)",
  "flags": ["list of field names that are uncertain or need review"],
  "source_language": "en or es or pt"
}}"""


def _parse_llm_response(raw: str) -> dict:
    """Extract and parse JSON from LLM response, handling markdown fences."""
    raw = _RE_MD_FENCE.sub('', raw)
    raw = _RE_MD_CLOSE.sub('', raw)
    raw = raw.strip()

    if raw.startswith('{'):
        json_str = raw
    else:
        m = _RE_JSON_OBJ.search(raw)
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
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{'role': 'user', 'content': user_prompt}],
    )
    raw = message.content[0].text.strip()
    return _parse_llm_response(raw)


def _ollama_call(payload_bytes: bytes, timeout: int = 600, model: str = '') -> str:
    """Send a payload to Ollama /api/chat and return the response content string."""
    for attempt in range(2):
        req = urllib.request.Request(
            f'{OLLAMA_BASE_URL}/api/chat',
            data=payload_bytes,
            method='POST',
            headers={'Content-Type': 'application/json'},
        )
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read())
                return data['message']['content'].strip()
        except TimeoutError:
            if attempt == 0:
                continue
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
                    f'3. 모델 다운로드: ollama pull {model or "<model>"}'
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
    return ''


def _translate_title_and_products_ollama(
    title_text: str, products_text: str, terminology: dict, model: str
) -> dict:
    """
    Dedicated focused Ollama call for 제목 + 해당품목 translation only.

    Small LLMs perform much better on a two-field translation prompt than when
    buried inside the 18-field classification prompt.  No 주간보고 examples
    are included here, so the model cannot accidentally copy product names.
    Returns a dict with keys '제목' and '해당품목'; empty strings on failure.
    """
    fallback = {'제목': '', '해당품목': ''}
    if not title_text and not products_text:
        return fallback

    doc_text = f'{title_text} {products_text}'
    selected = _select_terms(terminology, doc_text, max_terms=30)
    term_lines = '\n'.join(f'  {k} → {v}' for k, v in selected)

    prompt = f"""Translate two fields from a WTO SPS notification into Korean.
Return ONLY valid JSON with keys "제목" and "해당품목". No other text.

=== 제목 (Title) RULES ===
- Translate ONLY the title text — ignore any Language/Number-of-pages metadata lines.
- Resolution format: "Resolution No. X-Y-Z" or "Resolución No. X-Y-Z"
    → "결의안 제X-Y-Z호" (NEVER keep "No." in the Korean output)
- Standard structure:
    "[country_kr]산 [product_kr] 수입에 대한 [requirement_kr]을 규정하는 결의안 제[number]호"
- Keep scientific names in parentheses exactly as written.
- Translate origin phrases as '[country_kr]산':
    "originating in Peru" → "페루산" | "procedente de Argentina" → "아르헨티나산"
- If title is NOT a Resolution, translate it naturally into Korean.

제목 EXAMPLES (format only — never copy product/country names into unrelated documents):
  "Resolution No. 175-2026-IPSA establishing phytosanitary requirements for fresh [X] (Xx. xx) from [Y]"
    → "[Y_kr]산 신선한 [X_kr](Xx. xx) 수입에 대한 식물검역요건을 규정하는 결의안 제175-2026-IPSA호"
  "Resolución No. 045-2025 que establece los requisitos sanitarios para la importación de [X] (Xx. xx) procedente de [Y]"
    → "[Y_kr]산 [X_kr](Xx. xx) 수입에 대한 위생요건을 규정하는 결의안 제045-2025호"

=== 해당품목 (Products covered) RULES ===
- Keep scientific names in parentheses exactly as written — never translate them.
- Korean modifier-first: qualifying clauses come BEFORE the noun.
    "pigs at risk of transmitting Aujeszky's disease virus"
      → "오제스키병 바이러스 전파 위험이 있는 돼지"
    "products and by-products of swine origin"
      → "돼지 유래 제품 및 부산물"

=== TERMINOLOGY ===
{term_lines}

=== INPUT ===
Title: {title_text}
Products covered: {products_text}

=== OUTPUT (JSON only) ===
{{"제목": "...", "해당품목": "..."}}"""

    payload = json.dumps({
        'model': model,
        'messages': [
            {'role': 'system', 'content': 'You are a Korean agricultural document translator. Output ONLY valid JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        'stream': False,
        'options': {
            'temperature': 0.05,
            'num_predict': 512,
            'num_ctx': 8192,
        },
    }).encode('utf-8')

    try:
        raw = _ollama_call(payload, timeout=180, model=model)
        raw = _RE_MD_FENCE.sub('', raw)
        raw = _RE_MD_CLOSE.sub('', raw).strip()
        if not raw.startswith('{'):
            m = _RE_JSON_OBJ.search(raw)
            raw = m.group() if m else '{}'
        parsed = json.loads(raw)
        result = {}
        for key in ('제목', '해당품목'):
            val = str(parsed.get(key, '')).strip()
            # Reject if it looks like a verbatim copy of the English input
            if val and val.lower() != (title_text if key == '제목' else products_text).lower():
                result[key] = val
        return result
    except Exception:
        return fallback


def _process_with_ollama(parsed: dict, export_items: str, terminology: dict, model: str) -> dict:
    user_prompt = _build_user_prompt(parsed, export_items, terminology)
    payload = json.dumps({
        'model': model,
        'messages': [
            {'role': 'system', 'content': SYSTEM_PROMPT},
            {'role': 'user', 'content': user_prompt},
        ],
        'stream': False,
        'options': {
            'temperature': 0.1,
            'num_predict': 4096,
            'num_ctx': 16384,   # prevent silent truncation on long documents
        },
    }).encode('utf-8')

    raw = _ollama_call(payload, timeout=600, model=model)
    result = _parse_llm_response(raw)

    # Override 제목 + 해당품목 with a dedicated focused two-field translation call.
    # The small model translates these text fields much better when given a focused
    # prompt without all the classification rules and 주간보고 examples.
    title_text = '\n'.join(
        ln for ln in parsed.get('title', '').split('\n')
        if not _RE_LANG_PAGE.match(ln)
    ).strip()
    products_text = parsed.get('products', '')
    if title_text or products_text:
        overrides = _translate_title_and_products_ollama(
            title_text, products_text, terminology, model
        )
        for key in ('제목', '해당품목'):
            if overrides.get(key):
                result[key] = overrides[key]

    return result


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
