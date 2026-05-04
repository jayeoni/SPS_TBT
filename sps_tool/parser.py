"""
WTO SPS notification Word document parser.
Extracts structured fields from the standardized WTO SPS form.
"""
import re
import os
from pathlib import Path
import docx
from docx.oxml.ns import qn


# ── Label patterns for each field (English and Korean variants) ──────────────
LABEL_PATTERNS = {
    'notifying_member': ['notifying member', '통보회원국', 'notifying country'],
    'agency':           ['agency responsible', '담당기관', 'responsible agency'],
    'products':         ['products covered', '적용대상품목', 'products'],
    'regions':          ['regions/countries', 'regions', 'countries', '지역/국가', '국가/지역'],
    'title':            ['title', '제목'],
    'description':      ['description of content', '내용', 'description'],
    'objective':        ['objective', '목적', 'reason'],
    'standards':        ['international standards', '국제기준', 'international standard'],
    'adoption_date':    ['proposed date of adoption', 'date of adoption', '채택예정일', 'proposed date of publication', '발행예정일'],
    'comment_deadline': ['final date for comments', '의견마감일', 'comment period', 'date for comments'],
    'entry_force':      ['proposed date of entry into force', '발효예정일', 'entry into force'],
    'distribution':     ['distribution date', '배포일', 'circulated', 'circulation date'],
}

OBJECTIVE_MAP = {
    'food safety':        '식품안전',
    'animal health':      '동물위생',
    'plant protection':   '식물보호',
    'protect humans':     '동식물 병해충 또는 질병으로부터 사람 보호',
    'protect territory':  '해충으로 인한 피해로부터 영토 보호',
    'protect humans from animal': '동식물 병해충 또는 질병으로부터 사람 보호',
}

DOC_NUMBER_RE = re.compile(
    r'G/SPS/[A-Z]+/[A-Z]{2,3}/[\d]+(?:/Add\.[\d]+)?',
    re.IGNORECASE
)


def _unique_cells(row):
    """Return deduplicated cells from a table row (merged cells repeat in python-docx)."""
    seen = set()
    result = []
    for cell in row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid)
            result.append(cell)
    return result


def _cell_text(cell):
    return cell.text.strip()


def _all_text(doc):
    """Get all text from paragraphs + tables for pattern searching."""
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in _unique_cells(row):
                parts.append(cell.text)
    return '\n'.join(parts)


def _extract_doc_number(text, filename=''):
    """Find the WTO document symbol in text, fall back to filename parsing."""
    m = DOC_NUMBER_RE.search(text)
    if m:
        return m.group().upper()

    # Filename fallback: GSPSNBRA2474 → G/SPS/N/BRA/2474
    base = Path(filename).stem.upper()
    # Remove _번역 suffix if present
    base = re.sub(r'_번역$', '', base)
    m2 = re.match(r'GSPSE?N([A-Z]{2,3})(\d+)(A(\d+))?$', base)
    if m2:
        country = m2.group(1)
        number  = m2.group(2)
        add_num = m2.group(4)
        result  = f'G/SPS/N/{country}/{number}'
        if add_num:
            result += f'/Add.{add_num}'
        return result
    return ''


def _detect_type(full_text, doc_number, filename=''):
    """Return dict with is_emergency and is_addendum flags."""
    head = full_text[:600].lower()
    is_emergency = (
        'emergency' in head or
        'g/sps/n/ems' in doc_number.lower()
    )
    is_addendum = (
        'addendum' in head or
        '/add.' in doc_number.lower() or
        re.search(r'A\d+$', Path(filename).stem.upper()) is not None
    )
    return {'is_emergency': is_emergency, 'is_addendum': is_addendum}


def _match_label(cell_text, patterns):
    t = cell_text.lower()
    return any(p in t for p in patterns)


def _extract_field_from_tables(doc, label_patterns):
    """
    Find the content for a field by its label patterns.

    Layout A (WTO standard): ['1.', 'Label: content...']
      The label is embedded at the start of the last cell.
      Extract everything after the first colon.

    Layout B (older format): ['Label cell', 'Content cell']
      The label is in a dedicated earlier cell; return the last cell.
    """
    for table in doc.tables:
        for row in table.rows:
            cells = _unique_cells(row)
            if len(cells) < 2:
                continue

            content_cell = cells[-1]
            content_text = _cell_text(content_cell)
            if not content_text:
                continue

            # Layout A: label at the start of the content cell (first line)
            first_line = content_text.split('\n')[0][:150]
            if _match_label(first_line, label_patterns):
                colon_pos = content_text.find(':')
                if colon_pos != -1:
                    return content_text[colon_pos + 1:].strip()
                return content_text

            # Layout B: label is in a dedicated earlier cell
            for cell in cells[:-1]:
                if _match_label(_cell_text(cell), label_patterns):
                    if content_text and len(content_text) > 1 and content_text != _cell_text(cell):
                        return content_text
    return ''


def _extract_objectives(doc):
    """
    Find checked objectives ([X] or ☒ markers) and return Korean phrases.
    """
    checked = []
    for table in doc.tables:
        for row in table.rows:
            for cell in _unique_cells(row):
                text = cell.text
                # Only process cells that have a checked mark
                if not ('[x]' in text.lower() or '☒' in text or '[X]' in text):
                    continue
                text_lower = text.lower()
                for eng_key, kor_val in OBJECTIVE_MAP.items():
                    if eng_key in text_lower and kor_val not in checked:
                        checked.append(kor_val)
    return checked


def _extract_regions(doc):
    """
    Extract the regions/countries field. Returns '모든 교역국' if all trading
    partners is checked, otherwise returns specific country names.
    """
    regions_raw = _extract_field_from_tables(doc, LABEL_PATTERNS['regions'])

    # Look for 'all trading partners' checkbox anywhere
    all_text = _all_text(doc)
    if re.search(r'\[x\].*?all trading partners', all_text, re.IGNORECASE | re.DOTALL):
        return '모든 교역국'
    if re.search(r'\[x\].*?모든 교역국', all_text, re.DOTALL):
        return '모든 교역국'

    # Look for specific regions checked
    specific_match = re.search(
        r'\[x\][^\n]*?specific regions?(?:\s+or\s+countries?)?\s*:\s*([^\n\[]+)',
        all_text, re.IGNORECASE
    )
    if specific_match:
        return specific_match.group(1).strip()

    return regions_raw


def _detect_language(text):
    """
    Detect dominant source language from character distribution.
    Returns 'en', 'es', or 'pt'.
    """
    if not text:
        return 'en'
    # Spanish/Portuguese indicator characters
    sp_pt_chars = set('áéíóúàèìòùâêîôûãõñüçÁÉÍÓÚÀÈÌÒÙÂÊÎÔÛÃÕÑÜÇ')
    count = sum(1 for c in text if c in sp_pt_chars)
    if count == 0:
        return 'en'
    # Very rough heuristic: ã/õ = probably Portuguese, ñ = probably Spanish
    pt_chars = set('ãõÃÕ')
    es_chars = set('ñÑ')
    pt_count = sum(1 for c in text if c in pt_chars)
    es_count = sum(1 for c in text if c in es_chars)
    if pt_count > es_count:
        return 'pt'
    return 'es'


def _extract_addendum_fields(doc, full_text):
    """
    For addendum documents, extract the specific addendum metadata.
    Returns a dict with addendum_type and relevant content.
    """
    result = {
        'addendum_concerns': [],
        'addendum_country': '',
        'addendum_received_date': '',
        'addendum_content': '',
    }

    # Extract circulation country from opening sentence
    m = re.search(
        r'being circulated at the request of the delegation of ([A-Z][A-Z ]+)',
        full_text, re.IGNORECASE
    )
    if m:
        result['addendum_country'] = m.group(1).strip()

    # Extract received date
    m2 = re.search(
        r'received on ([0-9\w ,]+)',
        full_text, re.IGNORECASE
    )
    if m2:
        result['addendum_received_date'] = m2.group(1).strip()

    # Find checked addendum type boxes
    addendum_types = {
        'notification of adoption':  '채택·발행·발효 통보',
        'modification of final date': '의견마감일 변경',
        'modification of content':   '내용/범위 변경',
        'withdrawal':                '규정 철회',
        'change in proposed dates':  '제안 일자 변경',
    }
    for eng, kor in addendum_types.items():
        pattern = r'\[x\][^\n]*' + re.escape(eng)
        if re.search(pattern, full_text, re.IGNORECASE):
            result['addendum_concerns'].append(kor)

    return result


def parse_notification(docx_path: str) -> dict:
    """
    Parse a WTO SPS notification Word file and return structured fields.

    Returns a dict with all extracted raw fields. The LLM will later
    translate and normalize these into Korean institutional language.
    """
    doc = docx.Document(docx_path)
    filename = os.path.basename(docx_path)
    full_text = _all_text(doc)

    result = {
        'filename':             filename,
        'doc_number':           '',
        'is_emergency':         False,
        'is_addendum':          False,
        'notifying_member':     '',
        'agency':               '',
        'products':             '',
        'regions':              '',
        'title':                '',
        'description':          '',
        'objectives_raw':       [],
        'objectives_korean':    [],
        'comment_deadline_raw': '',
        'entry_force_raw':      '',
        'adoption_date_raw':    '',
        'source_language':      'en',
        'addendum':             {},
    }

    # ── Document number ────────────────────────────────────────────────────
    result['doc_number'] = _extract_doc_number(full_text, filename)

    # ── Notification type ──────────────────────────────────────────────────
    type_flags = _detect_type(full_text, result['doc_number'], filename)
    result['is_emergency'] = type_flags['is_emergency']
    result['is_addendum']  = type_flags['is_addendum']

    # ── Field extraction ───────────────────────────────────────────────────
    result['notifying_member'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['notifying_member'])
    result['agency'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['agency'])
    result['products'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['products'])
    result['regions'] = _extract_regions(doc)
    result['title'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['title'])
    result['description'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['description'])

    # ── Dates ──────────────────────────────────────────────────────────────
    result['comment_deadline_raw'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['comment_deadline'])
    result['entry_force_raw'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['entry_force'])
    result['adoption_date_raw'] = _extract_field_from_tables(
        doc, LABEL_PATTERNS['adoption_date'])

    # ── Objectives (checkboxes) ────────────────────────────────────────────
    result['objectives_korean'] = _extract_objectives(doc)

    # ── Language detection ────────────────────────────────────────────────
    detect_text = result['description'] or result['title'] or result['products']
    result['source_language'] = _detect_language(detect_text)

    # ── Addendum-specific fields ──────────────────────────────────────────
    if result['is_addendum']:
        result['addendum'] = _extract_addendum_fields(doc, full_text)

    return result
