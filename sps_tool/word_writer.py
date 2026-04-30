"""
Bilingual Word document generator.
Inserts Korean translation paragraphs into WTO SPS notification Word files.
"""
import os
import shutil
import copy
from pathlib import Path
import docx
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_COLOR_INDEX

# Fields we insert Korean text for, keyed by partial label match
TRANSLATE_LABELS = {
    'title':       ['title', '제목'],
    'products':    ['products covered', '적용대상품목', 'products'],
    'regions':     ['regions/countries', 'regions', 'countries', '지역/국가'],
    'description': ['description of content', '내용', 'content'],
    'objective':   ['objective', '목적', 'reason'],
    'notifying_member': ['notifying member', '통보회원국'],
}

# Map from TRANSLATE_LABELS keys → field keys in the LLM output dict
FIELD_MAP = {
    'title':       '제목',
    'products':    '해당품목',
    'regions':     '해당국가',
    'description': '내용',
    'objective':   '목적',
    'notifying_member': None,  # we just copy the country name from parsed data
}

LIME_RGB   = (204, 255, 153)   # #CCFF99 — non-English source flag
KOREAN_FONT = '맑은 고딕'


def _unique_cells(row):
    seen = set()
    result = []
    for cell in row.cells:
        cid = id(cell._tc)
        if cid not in seen:
            seen.add(cid)
            result.append(cell)
    return result


def _cell_matches_label(cell_text: str, patterns: list) -> bool:
    t = cell_text.lower()
    return any(p in t for p in patterns)


def _is_content_cell(cell, row_cells):
    """True if this cell is the last (content) cell in a row, not a label cell."""
    return cell == row_cells[-1] and len(row_cells) > 1


def _set_cell_bg(cell, rgb: tuple):
    """Set cell background shading to an RGB colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    hex_color = '{:02X}{:02X}{:02X}'.format(*rgb)
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    # Remove existing shd if present
    for existing in tcPr.findall(qn('w:shd')):
        tcPr.remove(existing)
    tcPr.append(shd)


def _add_korean_paragraph(cell, korean_text: str, is_lime: bool = False):
    """
    Add a Korean translation paragraph at the end of a table cell.
    Uses 맑은 고딕 font, same size as existing paragraphs.
    """
    if not korean_text:
        return

    # Detect existing font size from first paragraph
    existing_size = None
    if cell.paragraphs:
        for para in cell.paragraphs:
            for run in para.runs:
                if run.font.size:
                    existing_size = run.font.size
                    break
            if existing_size:
                break

    # Add a new paragraph
    para = cell.add_paragraph()
    run = para.add_run(korean_text)
    run.font.name = KOREAN_FONT
    # Set East Asian font
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), KOREAN_FONT)
    rPr.append(rFonts)
    if existing_size:
        run.font.size = existing_size

    if is_lime:
        # Highlight the Korean paragraph's run
        _set_cell_bg(cell, LIME_RGB)


def create_bilingual_docx(
    source_path: str,
    translations: dict,
    is_non_english: bool = False,
) -> str:
    """
    Create a bilingual Word file by inserting Korean translations into the source docx.

    translations dict should contain keys like 제목, 내용, 해당품목, 목적, 해당국가.
    Returns the output file path (*_번역.docx).
    """
    # Determine output path
    source = Path(source_path)
    output_path = source.parent / (source.stem + '_번역.docx')

    # Copy source file
    shutil.copy2(source_path, output_path)

    doc = Document(str(output_path))

    # Build a lookup: label_key → which translation dict key to use
    label_to_translation = {
        'title':            translations.get('제목', ''),
        'products':         translations.get('해당품목', ''),
        'regions':          translations.get('해당국가', ''),
        'description':      translations.get('내용', ''),
        'objective':        translations.get('목적', ''),
        'notifying_member': translations.get('통보국_kr', ''),
    }

    for table in doc.tables:
        for row in table.rows:
            cells = _unique_cells(row)
            if len(cells) < 2:
                continue

            # Check label cells (all except last) against known patterns
            matched_key = None
            for label_key, patterns in TRANSLATE_LABELS.items():
                # Check all non-last cells for label match
                for cell in cells[:-1]:
                    if _cell_matches_label(cell.text, patterns):
                        matched_key = label_key
                        break
                if matched_key:
                    break

            if not matched_key:
                continue

            korean_text = label_to_translation.get(matched_key, '')
            if not korean_text:
                continue

            # Add Korean to the content (last) cell
            content_cell = cells[-1]
            apply_lime = is_non_english and matched_key in ('title', 'description')
            _add_korean_paragraph(content_cell, korean_text, apply_lime)

    doc.save(str(output_path))
    return str(output_path)
