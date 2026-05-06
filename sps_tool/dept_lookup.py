"""
Department lookup and country-name translation.

DEPT_TABLE mirrors the 관련부서 table on the '★매뉴얼' sheet (first sheet) of
the monthly tracking Excel.  Key: (구분, 통보내용, 세부사항|None) → [dept list].
Lookup falls back from the specific (구분, 통보내용, 세부사항) triple to the
generic (구분, 통보내용, None) when the sub-type is not found.

COUNTRY_KR provides direct English → Korean country-name translation so the
'해당국가' column is filled without relying on the LLM.
"""
import re

# ── Department lookup table ───────────────────────────────────────────────────
# Exactly matches the ★매뉴얼 sheet rows 3-45 (columns F-K).
# None as 세부사항 = default / no sub-type required.

DEPT_TABLE = {
    # ── 식물 ──────────────────────────────────────────────────────────────
    ('식물', '식물검역', None):              ['수출지원과'],
    ('식물', '식물검역', '식물'):            ['수출지원과'],
    ('식물', '식물검역', '종자'):            ['수출지원과'],
    ('식물', '식물검역', '목재'):            ['수출지원과'],
    ('식물', '식물검역', '식물성비료/농기계'): ['농산업수출진흥과', '수출지원과'],
    ('식물', '식물검역', '목재포장재'):       ['식물방제과'],
    ('식물', '식물검역', '병해충'):           ['수출지원과', '위험관리과'],
    ('식물', '비료', None):                  ['농산업수출진흥과'],
    # ── 동물 ──────────────────────────────────────────────────────────────
    ('동물', '동물검역', None):              ['동물검역과', '위험평가과'],
    ('동물', '동물검역', '동물'):            ['동물검역과', '위험평가과'],
    ('동물', '동물검역', '축산물'):          ['위험평가과', '축산물수출위생팀'],
    ('동물', '동물검역', '동물성비료'):      ['농산업수출진흥과', '위험평가과', '축산물수출위생팀'],
    ('동물', '동물검역', '야생동물'):        ['기후부'],
    ('동물', '동물검역', '수산물'):          ['해수부'],
    ('동물', '사료첨가제', None):            ['축산환경자원과', '축산물수출위생팀'],
    ('동물', '사료첨가제', '가축'):          ['축산환경자원과', '축산물수출위생팀'],
    ('동물', '사료첨가제', '반려동물'):      ['반려산업동물의료과', '축산환경자원과', '축산물수출위생팀'],
    ('동물', '침입외래종', None):            ['기후부'],
    ('동물', '침입외래종', '동물'):          ['기후부'],
    ('동물', '침입외래종', '식물체'):        ['산림청'],
    # ── 식품 ──────────────────────────────────────────────────────────────
    ('식품', '농약', None):                  ['농식품수출진흥과', '잔류화학평가과(농과원)'],
    ('식품', '농약', '농산물'):              ['농식품수출진흥과', '잔류화학평가과(농과원)'],
    ('식품', '농약', '축산물'):              ['동물약품평가과', '축산물수출위생팀'],
    ('식품', '농약', '사료'):               ['축산환경자원과', '동물약품평가과'],
    ('식품', '농약', '천연식물보호제'):      ['안전성분석과(농관원)'],
    ('식품', '농약', '규정'):               ['농산업수출진흥과'],
    ('식품', '동물용의약품', None):          ['동물약품평가과'],
    ('식품', '동물용의약품', 'MRL'):         ['동물약품평가과'],
    ('식품', '동물용의약품', '항생제'):      ['조류인플루엔자방역과', '동물약품평가과', '축산물수출위생팀'],
    ('식품', '동물용의약품', '규정'):        ['농산업수출진흥과'],
    ('식품', 'GMO/LMO', None):             ['수출지원과', '연구개발과(농진청)', '생물안전성과(농과원)'],
    ('식품', 'GMO/LMO', '사료'):           ['축산환경자원과', '연구개발과(농진청)', '생물안전성과(농과원)'],
    ('식품', 'GMO/LMO', '식물체'):         ['수출지원과', '연구개발과(농진청)', '생물안전성과(농과원)'],
    ('식품', 'GMO/LMO', '종자'):           ['종자산업지원과', '연구개발과(농진청)', '생물안전성과(농과원)'],
    ('식품', 'GMO/LMO', '식품'):           ['연구개발과(농진청)', '생물안전성과(농과원)', '식약처'],
    ('식품', '농산물', None):               ['품질조사과(농관원)', '식약처'],
    ('식품', '농산물', '품질'):             ['품질조사과(농관원)', '식약처'],
    ('식품', '농산물', '중금속'):           ['안전성분석과(농과원)', '식약처'],
    ('식품', '농산물', '곰팡이독소'):       ['안전성분석과(농과원)', '농업미생물과(농과원)', '식약처'],
    ('식품', '축산물', None):               ['축산물수출위생팀', '식약처'],
    ('식품', '축산물', '위생·안전'):        ['축산물수출위생팀', '식약처'],
    ('식품', '축산물', '품질'):             ['식약처'],
    ('식품', '사료', None):                ['축산환경자원과'],
    ('식품', '사료', '기준 및 규격'):       ['축산환경자원과'],
    ('식품', '특용작물', None):             ['농식품수출진흥과', '원예산업과'],
    ('식품', '친환경·유기농산물', None):    ['친환경농업과', '인증관리과(농관원)'],
    ('식품', '식용곤충·양잠', None):        ['그린바이오산업팀'],
    ('식품', '팽이버섯', None):             ['농식품수출진흥과', '소비안전과(농관원)'],
    ('식품', '신소재식품', None):           ['연구개발과(농진청)', '식약처'],
    ('식품', '할랄식품', None):             ['농식품수출진흥과'],
    ('식품', '식품첨가물', None):           ['식약처'],
    ('식품', '미생물/가공식품/제조시설', None): ['식약처'],
    ('식품', '수산물', None):              ['해수부', '식약처'],
    ('식품', '수산물', '위생품질'):         ['해수부', '식약처'],
    ('식품', '물/살생물제품', None):        ['기후부'],
    ('식품', '담배', None):               ['식약처', '보건복지부'],
}


def lookup_dept(구분: str, 통보내용: str, 세부사항: str = '') -> str:
    """
    Look up 관련부서 from the 매뉴얼 table.
    Returns newline-joined department string, or '' if no match found.
    """
    if not 구분 or not 통보내용:
        return ''
    세부 = 세부사항.strip() if 세부사항 else None
    # Try exact triple first
    if 세부:
        depts = DEPT_TABLE.get((구분, 통보내용, 세부))
        if depts:
            return '\n'.join(depts)
    # Fall back to default (None sub-type)
    depts = DEPT_TABLE.get((구분, 통보내용, None))
    if depts:
        return '\n'.join(depts)
    return ''


# ── Country name lookup ───────────────────────────────────────────────────────

COUNTRY_KR = {
    # Korean peninsula
    'republic of korea':        '대한민국',
    'south korea':              '대한민국',
    'korea':                    '대한민국',
    'democratic people':        '북한',
    'north korea':              '북한',
    # Americas
    'united states of america': '미국',
    'united states':            '미국',
    'usa':                      '미국',
    'canada':                   '캐나다',
    'mexico':                   '멕시코',
    'brazil':                   '브라질',
    'argentina':                '아르헨티나',
    'chile':                    '칠레',
    'peru':                     '페루',
    'colombia':                 '콜롬비아',
    'ecuador':                  '에콰도르',
    'bolivia':                  '볼리비아',
    'uruguay':                  '우루과이',
    'paraguay':                 '파라과이',
    'venezuela':                '베네수엘라',
    'costa rica':               '코스타리카',
    'guatemala':                '과테말라',
    'panama':                   '파나마',
    'honduras':                 '온두라스',
    'nicaragua':                '니카라과',
    'el salvador':              '엘살바도르',
    'cuba':                     '쿠바',
    'dominican republic':       '도미니카공화국',
    'jamaica':                  '자메이카',
    'trinidad':                 '트리니다드',
    # Europe / EU
    'european union':           '유럽연합',
    'eu':                       '유럽연합',
    'united kingdom':           '영국',
    'germany':                  '독일',
    'france':                   '프랑스',
    'italy':                    '이탈리아',
    'spain':                    '스페인',
    'netherlands':              '네덜란드',
    'belgium':                  '벨기에',
    'poland':                   '폴란드',
    'austria':                  '오스트리아',
    'switzerland':              '스위스',
    'sweden':                   '스웨덴',
    'norway':                   '노르웨이',
    'denmark':                  '덴마크',
    'finland':                  '핀란드',
    'portugal':                 '포르투갈',
    'greece':                   '그리스',
    'czech republic':           '체코',
    'czechia':                  '체코',
    'hungary':                  '헝가리',
    'romania':                  '루마니아',
    'bulgaria':                 '불가리아',
    'croatia':                  '크로아티아',
    'serbia':                   '세르비아',
    'ukraine':                  '우크라이나',
    'russia':                   '러시아',
    'russian federation':       '러시아',
    'turkey':                   '튀르키예',
    'türkiye':                  '튀르키예',
    'kazakhstan':               '카자흐스탄',
    'uzbekistan':               '우즈베키스탄',
    'georgia':                  '조지아',
    # Asia-Pacific
    'china':                    '중국',
    "people's republic of china": '중국',
    'japan':                    '일본',
    'taiwan':                   '대만',
    'hong kong':                '홍콩',
    'india':                    '인도',
    'indonesia':                '인도네시아',
    'philippines':              '필리핀',
    'vietnam':                  '베트남',
    'viet nam':                 '베트남',
    'thailand':                 '태국',
    'malaysia':                 '말레이시아',
    'singapore':                '싱가포르',
    'myanmar':                  '미얀마',
    'cambodia':                 '캄보디아',
    'laos':                     '라오스',
    'bangladesh':               '방글라데시',
    'pakistan':                 '파키스탄',
    'sri lanka':                '스리랑카',
    'nepal':                    '네팔',
    'australia':                '호주',
    'new zealand':              '뉴질랜드',
    'papua new guinea':         '파푸아뉴기니',
    # Middle East
    'saudi arabia':             '사우디아라비아',
    'united arab emirates':     '아랍에미리트',
    'uae':                      '아랍에미리트',
    'israel':                   '이스라엘',
    'iran':                     '이란',
    'iraq':                     '이라크',
    'jordan':                   '요르단',
    'egypt':                    '이집트',
    'oman':                     '오만',
    'kuwait':                   '쿠웨이트',
    'qatar':                    '카타르',
    'bahrain':                  '바레인',
    # Africa
    'south africa':             '남아프리카공화국',
    'nigeria':                  '나이지리아',
    'kenya':                    '케냐',
    'ethiopia':                 '에티오피아',
    'ghana':                    '가나',
    'tanzania':                 '탄자니아',
    'morocco':                  '모로코',
    'algeria':                  '알제리',
    'Tunisia':                  '튀니지',
    'senegal':                  '세네갈',
    'cameroon':                 '카메룬',
    'ivory coast':              '코트디부아르',
    "côte d'ivoire":            '코트디부아르',
    'zambia':                   '잠비아',
    'zimbabwe':                 '짐바브웨',
    'mozambique':               '모잠비크',
    'madagascar':               '마다가스카르',
}


def translate_regions(regions_text: str) -> str:
    """
    Translate the raw 'regions' string from the WTO form to Korean.

    Returns:
      '모든 교역국' for "All trading partners"
      Korean country name(s) for specific countries
      Original text if no match found (as fallback)
    """
    if not regions_text:
        return ''
    text = regions_text.strip()

    # Already translated
    if text == '모든 교역국':
        return text

    # All-partners phrases
    if re.search(r'\ball trading partners\b', text, re.IGNORECASE):
        return '모든 교역국'

    # Try to split multiple countries (e.g., "Republic of Korea, Japan")
    parts = [p.strip() for p in re.split(r'[;,/]', text) if p.strip()]
    kr_parts = []
    for part in parts:
        lower = part.lower()
        matched = None
        # Longest-match first
        for eng, kor in sorted(COUNTRY_KR.items(), key=lambda x: -len(x[0])):
            if eng in lower:
                matched = kor
                break
        kr_parts.append(matched or part)

    result = ', '.join(kr_parts)
    return result
