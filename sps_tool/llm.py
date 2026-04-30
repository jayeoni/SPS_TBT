"""
Claude API integration for WTO SPS notification processing.
Handles translation, normalization, classification, and summary generation.
"""
import json
import os
import anthropic

MODEL = 'claude-sonnet-4-6'

SYSTEM_PROMPT = """You are an expert assistant for the Korean Ministry of Agriculture, Food and Rural Affairs (농림축산식품부, MAFRA) processing WTO SPS (Sanitary and Phytosanitary) notifications.

Your tasks:
1. Translate notification content from English/Spanish/Portuguese into formal Korean government language (공문체)
2. Classify notifications per MAFRA internal manual rules
3. Recommend handling based on domestic relevance

Rules for Korean government style:
- Use 개조식 (noun-phrase, concise bullet style) for 내용 and 주간보고 fields
- Include scientific names in 국문명(학명) format (e.g., 아보카도(Persea americana))
- Use standard institutional Korean terms, not casual translations
- 목적 field: use only the approved standard phrases, semicolons between multiples
- 주간보고: write as a single action phrase, like "브라질산 아보카도 식물체의 수입검역요건 개정"

Always output valid JSON only. No explanation text before or after the JSON."""


def _build_user_prompt(parsed: dict, export_items: str, terminology: dict) -> str:
    term_lines = '\n'.join(f'  {k} → {v}' for k, v in list(terminology.items())[:50])

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
SOURCE LANGUAGE: {parsed.get('source_language', 'en')}{addendum_info}

--- EXTRACTED FIELDS ---
Title: {parsed.get('title', '')}
Products covered: {parsed.get('products', '')}
Regions/countries affected: {parsed.get('regions', '')}
Objectives (checked): {objectives_str}
Description: {parsed.get('description', '')}
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

[관련부서]
Plant quarantine (seeds, wood): 수출지원과(검본)
Plant quarantine + wood/pest: 수출지원과(검본), 위험관리과(검본)
Plant pests/diseases: 수출지원과(검본), 위험관리과(검본)
Oilseed crops (차, 참깨, 견과류, 커피 etc): 원예산업과
Mushrooms/ginseng/medicinal: 농식품수출진흥과, 원예산업과
Organic/eco-friendly: 친환경농업과, 인증관리과(농관원)
Insects/sericulture: 그린바이오산업팀
Fertilizer (plant): 농산업수출진흥과
GMO/LMO (plant): 수출지원과(검본), 연구개발과(농진청), 생물안전성과(농과원)
Seeds: 종자산업지원과, 연구개발과(농진청)
Animal quarantine (livestock): 동물검역과(검본), 위험평가과(검본)
Livestock products/meat: 위험평가과(검본), 축산물수출위생팀(검본)
Veterinary drug MRL: 동물약품평가과(검본), 축산물수출위생팀(검본)
Antibiotics: 조류인플루엔자방역과, 동물약품평가과(검본), 축산물수출위생팀(검본)
Pet animals: 반려산업동물의료과, 축산환경자원과, 축산물수출위생팀(검본)
HPAI: 동물검역과(검본), 위험평가과(검본), 축산물수출위생팀(검본), 조류인플루엔자방역과
ASF/FMD: 동물검역과(검본), 위험평가과(검본), 축산물수출위생팀(검본)
Feed/feed additives (livestock): 축산환경자원과, 축산물수출위생팀(검본)
Wildlife/invasive species: 기후부
Pesticide MRL (농산물): 농식품수출진흥과, 잔류화학평가과(농과원)
  (if export product exists: 농식품수출진흥과 already primary)
Agricultural quality: 품질조사과(농관원), 식약처
Heavy metals/mycotoxins: 안전성분석과(농과원), 식약처
Processed food/food additives/Codex: 식약처
Enoki mushroom (Listeria): 농식품수출진흥과, 소비안전과(농관원)
New food materials: 연구개발과(농진청), 식약처
GMO/LMO (food): 연구개발과(농진청), 생물안전성과(농과원), 식약처
Aquatic/fisheries: 해수부, 식약처
Feed standards: 축산환경자원과
Tobacco: 식약처, 보건복지부

--- OUTPUT FORMAT ---
Return ONLY this JSON object (no other text):
{{
  "제목": "Korean title (include scientific name as 국문명(학명) if present)",
  "내용": "Korean content summary in 개조식, 2-3 sentences",
  "해당품목": "Korean product name with 학명 if applicable",
  "목적": "Korean purpose phrase(s), semicolons between multiples. Use ONLY: 식품안전/동물위생/식물보호/동식물 병해충 또는 질병으로부터 사람 보호/해충으로 인한 피해로부터 영토 보호",
  "해당국가": "Korean country name or '모든 교역국'",
  "통보국_kr": "Korean name of the notifying member country",
  "주간보고": "Single-line 개조식 Korean summary of what this notification does",
  "구분": "동물 or 식물 or 식품",
  "구분_reason": "1-sentence reasoning",
  "중요도": "검토 or 참고 or -",
  "중요도_reason": "1-sentence reasoning citing specific rule",
  "관련부서": "Department 1\\nDepartment 2\\n(one per line)",
  "품목": "Short Korean product label (e.g., 옥수수(Zea mays) 종자 or 가금 및 가금제품)",
  "flags": ["list of field names that are uncertain or need review"],
  "source_language": "en or es or pt"
}}"""


def process_notification(
    parsed: dict,
    export_items: str,
    terminology: dict,
    api_key: str = None,
) -> dict:
    """
    Send the parsed notification to Claude for translation, classification,
    and summary generation.

    Returns a dict with Korean fields + recommendations + flags.
    """
    key = api_key or os.environ.get('ANTHROPIC_API_KEY', '')
    if not key:
        raise ValueError('ANTHROPIC_API_KEY가 설정되지 않았습니다.')

    client = anthropic.Anthropic(api_key=key)
    user_prompt = _build_user_prompt(parsed, export_items, terminology)

    message = client.messages.create(
        model=MODEL,
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=[{'role': 'user', 'content': user_prompt}],
    )

    raw = message.content[0].text.strip()

    # Extract JSON from the response
    json_match = None
    if raw.startswith('{'):
        json_match = raw
    else:
        import re
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if m:
            json_match = m.group()

    if not json_match:
        raise ValueError(f'LLM 응답에서 JSON을 찾을 수 없습니다: {raw[:200]}')

    result = json.loads(json_match)
    return result
