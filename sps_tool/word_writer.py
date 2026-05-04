"""
Bilingual Word document generator.
Translates all 13 rows of the WTO SPS notification form following the
exact format used in reference translated files (26.3월 SPS 통보문_금영★).
"""
import re
import shutil
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

KOREAN_FONT = '맑은 고딕'
LIME_RGB    = (204, 255, 153)

LANG_KR = {
    'english': '영어', 'spanish': '스페인어', 'portuguese': '포르투갈어',
    'french': '프랑스어', 'arabic': '아랍어', 'chinese': '중국어',
    'russian': '러시아어', 'german': '독일어', 'japanese': '일본어',
}

MONTH_KR = {
    'january': '1월', 'february': '2월', 'march': '3월', 'april': '4월',
    'may': '5월', 'june': '6월', 'july': '7월', 'august': '8월',
    'september': '9월', 'october': '10월', 'november': '11월', 'december': '12월',
}

# Row detection patterns: checked against first 150 chars of content cell (lowercase)
ROW_PATTERNS = {
    'notifying_member': ['notifying member', '통보회원국'],
    'agency':           ['agency responsible', '담당기관'],
    'products':         ['products covered', '적용대상품목'],
    'regions':          ['regions or countries', 'regions/countries', '지역/국가', '국가/지역'],
    'title':            ['title of the notified', 'title of the proposed', '제목'],
    'description':      ['description of content', '내용의 설명'],
    'objective':        ['objective and rationale'],
    'standards':        ['is there a relevant international standard'],
    'urgent_reason':    ['nature of the urgent problem'],
    'other_docs':       ['other relevant documents'],
    'adoption_date':    ['proposed date of adoption'],
    'entry_force':      ['proposed date of entry into force'],
    'comments':         ['final date for comments'],
    'texts_available':  ['text(s) available from', 'texts available from'],
}

OBJECTIVE_OPTIONS = [
    ('food safety',            '식품안전'),
    ('animal health',          '동물위생'),
    ('plant protection',       '식물보호'),
    ('protect humans from',    '동식물 해충/질병으로부터 사람 보호'),
    ('protect territory from', '해충으로 인한 피해로부터의 영토 보호'),
]

# IPPC ISPM titles in Korean (official translations)
ISPM_KR = {
    1:  'SPS 적용 원칙',
    2:  '병해충위험분석 개요',
    3:  '생물학적 방제제의 수출, 운송, 수입 및 방출 지침',
    4:  '병해충발생비지역 설정 요건',
    5:  '식물위생 용어 해설',
    6:  '모니터링 지침',
    7:  '식물위생증명체계',
    8:  '지역의 병해충 발생현황 규명',
    9:  '병해충 박멸 프로그램 지침',
    10: '병해충비발생 생산지 및 생산장소 설정 요건',
    11: '검역병해충에 대한 병해충위험분석',
    12: '식물위생증명서',
    13: '부적합 및 긴급조치의 통보 지침',
    14: '병해충 위험관리를 위한 종합방제체계 지침',
    15: '국제무역에 있어서 목재포장재의 규정',
    16: '규제비검역병해충: 개념 및 적용',
    17: '병해충 보고',
    18: '식물위생조치로서 방사선 조사의 사용 지침',
    19: '규제 병해충 목록 지침',
    20: '식물위생 수입 규정 체계 지침',
    21: '규제비검역병해충의 병해충위험분석',
    22: '병해충 저발생 지역 설정 요건',
    23: '검사 지침',
    24: '동등성 규정 및 인정 지침',
    25: '환적 화물',
    26: '과실파리(Tephritidae)의 병해충발생비지역 설정',
    27: '규제 병해충 진단 규정',
    28: '규제 병해충 식물위생처리',
    29: '병해충발생비지역 및 저발생 지역의 인정',
    30: '과실파리(Tephritidae)의 저발생 지역 설정',
    31: '화물 표본추출 방법론',
    32: '병해충 위험에 따른 품목 분류',
    33: '병해충발생비 감자(Solanum tuberosum) 종자',
    34: '식물 수입 후 격리검역소의 설계 및 운영',
    35: '과실파리 위험관리를 위한 종합방제체계',
    36: '재배용 식물의 종합방제조치',
    38: '종자의 국제이동',
    39: '신선 과실 및 채소의 국제이동',
    40: '재배용 식물과 관련된 재배 배지의 국제이동',
    41: '사용한 차량, 기계 및 장비의 국제이동',
    42: '식물위생조치로서 온도처리 사용 요건',
    43: '목재 식물위생조치로서 열처리 사용 요건',
}


# ── Helpers ──────────────────────────────────────────────────────────────────

def _unique_cells(row):
    seen = set()
    result = []
    for cell in row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid)
            result.append(cell)
    return result


def _translate_date(text: str) -> str:
    """Convert 'D Month YYYY' patterns to 'YYYY년 M월 D일'."""
    def _repl(m):
        day = str(int(m.group(1)))
        month_kr = MONTH_KR.get(m.group(2).lower(), m.group(2))
        return f'{m.group(3)}년 {month_kr} {day}일'
    return re.sub(
        r'(\d{1,2})\s+'
        r'(January|February|March|April|May|June|July|August|'
        r'September|October|November|December)\s+(\d{4})',
        _repl, text, flags=re.IGNORECASE,
    )


def _translate_date_phrase(text: str) -> str:
    """Translate standard date phrases and apply date conversion."""
    text = re.sub(r'(?i)to be determined after the end of the consultation period\.?',
                  '의견수렴기간 종료 후 결정', text)
    text = re.sub(r'(?i)to be determined', '추후 결정', text)
    return _translate_date(text)


def _checkbox(text: str, option_prefix: str) -> str:
    """Return '[X]' or '[  ]' based on whether option is checked in text."""
    m = re.search(r'\[([Xx☒ ]*)\]\s*' + re.escape(option_prefix[:18]),
                  text, re.IGNORECASE)
    if m:
        inner = m.group(1).strip().lower()
        return '[X]' if inner in ('x', '☒') else '[  ]'
    return '[  ]'


def _expand_ispm_numbers(raw_text: str) -> list:
    """Return list of 'ISPM N Korean-title' strings parsed from raw_text."""
    numbers = re.findall(r'\b(\d+)\b', raw_text)
    if not numbers:
        return []
    result = []
    for n in numbers:
        kr = ISPM_KR.get(int(n), '')
        result.append(f'ISPM {n} {kr}'.strip())
    return result


def _get_cell_font_size(cell):
    for para in cell.paragraphs:
        for run in para.runs:
            if run.font.size:
                return run.font.size
    return None


def _get_cell_para_style(cell):
    """Return the style name of the first non-empty paragraph in the cell."""
    for para in cell.paragraphs:
        if para.runs:
            return para.style.name
    return None


def _apply_korean_font(run):
    run.font.name = KOREAN_FONT
    rPr = run._r.get_or_add_rPr()
    for existing in rPr.findall(qn('w:rFonts')):
        rPr.remove(existing)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'),    KOREAN_FONT)
    rFonts.set(qn('w:hAnsi'),   KOREAN_FONT)
    rFonts.set(qn('w:eastAsia'), KOREAN_FONT)
    rPr.append(rFonts)


def _set_cell_bg(cell, rgb: tuple):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    hex_color = '{:02X}{:02X}{:02X}'.format(*rgb)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    for existing in tcPr.findall(qn('w:shd')):
        tcPr.remove(existing)
    tcPr.append(shd)


def _add_paragraph(cell, text: str, font_size=None, style_name=None):
    try:
        para = cell.add_paragraph(style=style_name)
    except Exception:
        para = cell.add_paragraph()
    run = para.add_run(text)
    _apply_korean_font(run)
    if font_size:
        run.font.size = font_size


# ── Per-row translation builders ──────────────────────────────────────────────

def _row_notifying_member(cell_text, t):
    country_kr = t.get('통보국_kr', '')
    if not country_kr:
        return []
    return [
        f'통보 회원국: {country_kr}',
        '적용 가능할 경우, 관련 지방 정부의 명칭 기재:',
    ]


def _row_agency(cell_text, t):
    agency_kr = t.get('담당기관_kr', '')
    if not agency_kr:
        return []
    return [f'담당 기관: {agency_kr}']


def _row_products(cell_text, t):
    products_kr = t.get('해당품목', '')
    if not products_kr:
        return []
    return [f'해당 품목(WTO의 국가별 세부 HS코드번호; 가능할 시 ICS 번호 표기): {products_kr}']


def _row_regions(cell_text, t):
    all_cb  = _checkbox(cell_text, 'All trading partners')
    spec_cb = _checkbox(cell_text, 'Specific regions')

    countries_kr = t.get('해당국가', '')
    spec_text = '특정 지역 및 국가: '
    if spec_cb == '[X]' and countries_kr and countries_kr != '모든 교역국':
        spec_text += countries_kr

    return [
        '영향을 받을 수 있는 지역이나 국가(관련되거나 가능한 한 최대한): ',
        f'{all_cb}\t모든 교역국  ',
        f'{spec_cb}    {spec_text}',
    ]


def _row_title(cell_text, t):
    title_kr = t.get('제목', '')
    if not title_kr:
        return []

    lang_m = re.search(r'Language\(s\):\s*(\w+)', cell_text, re.IGNORECASE)
    lang_en = lang_m.group(1).lower() if lang_m else ''
    lang_kr = LANG_KR.get(lang_en, lang_en)

    pages_m = re.search(r'Number of pages:\s*(\S+)', cell_text, re.IGNORECASE)
    pages = pages_m.group(1) if pages_m else ''

    line = f'통보 문서의 제목: {title_kr}'
    if lang_kr:
        line += f' 언어: {lang_kr}'
    if pages:
        line += f'  페이지수: {pages}'

    urls = re.findall(r'https?://\S+', cell_text)
    return [line] + urls


def _row_description(cell_text, t):
    desc_kr = t.get('내용', '')
    if not desc_kr:
        return []
    return [f'내용 설명: {desc_kr}']


def _row_objective(cell_text, t):
    parts = []
    for eng_prefix, kr_label in OBJECTIVE_OPTIONS:
        cb = _checkbox(cell_text, eng_prefix)
        parts.append(f'{cb} {kr_label}')
    return ['목적 및 근거: ' + ', '.join(parts)]


def _row_standards(cell_text, t):
    codex_cb = _checkbox(cell_text, 'Codex Alimentarius')
    oie_cb   = _checkbox(cell_text, 'World O')  # matches both Organisation/Organization
    ippc_cb  = _checkbox(cell_text, 'International Plant Protection')
    none_cb  = _checkbox(cell_text, 'None')

    # Preserve any content written after ISPM/Codex/OIE checkbox lines
    def _after_option(pattern):
        m = re.search(pattern + r'[^:\n]*:\s*([^\n\[]+)', cell_text, re.IGNORECASE)
        return m.group(1).strip() if m else ''

    codex_extra = _after_option(r'Codex Alimentarius')
    oie_extra   = _after_option(r'World Organ\w+ for Animal Health')
    ippc_extra  = _after_option(r'International Plant Protection')

    lines = ['관련 국제기준이 있는가? 있다면, 해당 기준을 표시']
    lines.append(f'{codex_cb} 국제식품규격위원회(Codex Alimentarius Commission) [예 ; Codex 규정 또는 관련문서의 제목 또는 문서번호] : {codex_extra}')
    lines.append(f'{oie_cb}  세계동물보건기구(OIE) (예 : 육상동물 또는 수생동물 위생규약, Chapter 번호) :  {oie_extra}')
    ippc_label = f'{ippc_cb} 국제식물보호협약(International Plant Protection Convention) [예: 식물위생조치를 위한 국제 기준(ISPM) 번호] :'
    ispm_lines = _expand_ispm_numbers(ippc_extra)
    if ispm_lines:
        lines.append(ippc_label)
        lines.extend(ispm_lines)
    else:
        lines.append(f'{ippc_label} {ippc_extra}')
    lines.append(f'{none_cb}  없음')
    lines.append('제안된 규정이 관련 국제기준과 일치하는가?')
    lines.append('[ ] 예   [ ]   아니오')
    lines.append('그렇지 않다면 가능한 경우는 항상 국제기준과 어떻게 다르고 왜 그러한지 설명:')
    return lines


def _row_other_docs(cell_text, t):
    return ['활용 가능한 다른 관련문서 및 언어 :']


def _row_adoption_date(cell_text, t):
    def _extract(label_pattern):
        m = re.search(label_pattern + r'[^:]*:\s*(.+?)(?=\n|Proposed date of pub|$)',
                      cell_text, re.IGNORECASE | re.DOTALL)
        return _translate_date_phrase(m.group(1).strip()) if m else '추후 결정'

    adopt = _extract(r'Proposed date of adoption')
    pub   = _extract(r'Proposed date of publication')
    return [
        f'채택예정일 [날짜(일/월/년)]: {adopt}',
        f'공표예정일 [날짜(일/월/년)]: {pub}',
    ]


def _row_entry_force(cell_text, t):
    six_cb   = _checkbox(cell_text, 'Six months')
    trade_cb = _checkbox(cell_text, 'Trade facilitating')

    m = re.search(r'and/or\s*\(dd/mm/yy\):\s*(.+?)(?:\n|$)', cell_text, re.IGNORECASE)
    date_kr = _translate_date_phrase(m.group(1).strip()) if m else ''

    return [
        f'효력 발생일: {six_cb} 공표일로부터6개월 후, 그리고/또는 [날짜(일/월/년)]: {date_kr}',
        f'{trade_cb}  무역 원활화 조치',
    ]


def _row_comments(cell_text, t):
    sixty_cb = _checkbox(cell_text, 'Sixty days')
    nna_cb   = _checkbox(cell_text, 'National Notification Authority')
    neq_cb   = _checkbox(cell_text, 'National Enquiry Point')

    m = re.search(r'and/or\s*\(dd/mm/yy\):\s*(.+?)(?:\n|Agency|$)',
                  cell_text, re.IGNORECASE | re.DOTALL)
    date_kr = _translate_date_phrase(m.group(1).strip()) if m else ''

    lines = [
        f'의견제출 마감일: {sixty_cb} 통보문 배포일로부터 60일 후 그리고/또는 [날짜(일/월/년)]: {date_kr}',
        f'의견 처리 담당기관 또는 관계당국: {nna_cb} 국가 통보처, {neq_cb} 국가 문의처 또는 (존재할 경우) 타 기관의 주소, 팩스 번호, 이메일 주소:',
    ]
    return lines


def _row_texts_available(cell_text, t):
    nna_cb = _checkbox(cell_text, 'National Notification Authority')
    neq_cb = _checkbox(cell_text, 'National Enquiry Point')
    return [
        f'전문 입수가 가능한 곳: {nna_cb} 국가 통보처, {neq_cb} 국가 문의처 또는 (존재할 경우) 타 기관의 주소, 팩스 번호, 이메일 주소: ',
    ]


ROW_BUILDERS = {
    'notifying_member': _row_notifying_member,
    'agency':           _row_agency,
    'products':         _row_products,
    'regions':          _row_regions,
    'title':            _row_title,
    'description':      _row_description,
    'objective':        _row_objective,
    'standards':        _row_standards,
    'other_docs':       _row_other_docs,
    'adoption_date':    _row_adoption_date,
    'entry_force':      _row_entry_force,
    'comments':         _row_comments,
    'texts_available':  _row_texts_available,
}


# ── Main ──────────────────────────────────────────────────────────────────────

def _detect_row_type(text: str):
    prefix = text[:150].lower()
    for row_type, patterns in ROW_PATTERNS.items():
        if any(p in prefix for p in patterns):
            return row_type
    return None


def create_bilingual_docx(
    source_path: str,
    translations: dict,
    is_non_english: bool = False,
) -> str:
    """
    Create a bilingual Word file by appending Korean translations to each
    content cell. Returns the output file path (*_번역.docx).
    """
    source = Path(source_path)
    output_path = source.parent / (source.stem + '_번역.docx')
    shutil.copy2(source_path, output_path)

    doc = Document(str(output_path))

    for table in doc.tables:
        for row in table.rows:
            cells = _unique_cells(row)
            if len(cells) < 2:
                continue

            content_cell = cells[-1]
            row_type = _detect_row_type(content_cell.text)

            # Fallback: check non-last cells (layout B: label | content)
            if not row_type and len(cells) > 1:
                for c in cells[:-1]:
                    row_type = _detect_row_type(c.text)
                    if row_type:
                        break

            if not row_type or row_type not in ROW_BUILDERS:
                continue

            korean_lines = ROW_BUILDERS[row_type](content_cell.text, translations)
            if not korean_lines:
                continue

            font_size  = _get_cell_font_size(content_cell)
            para_style = _get_cell_para_style(content_cell)
            for line in korean_lines:
                _add_paragraph(content_cell, line, font_size, para_style)

            if is_non_english and row_type in ('title', 'description'):
                _set_cell_bg(content_cell, LIME_RGB)

    doc.save(str(output_path))
    return str(output_path)
