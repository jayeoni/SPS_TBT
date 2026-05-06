"""
Excel row matching and cell writing for the SPS notification tracking workbook.
Finds the pre-populated row by 문서번호 and fills in all computed/LLM fields.
"""
import re
import shutil
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

# ── Column index mapping (1-based) ────────────────────────────────────────
# Based on observed structure: 담당자|순번|중요도|통보유형|통보국|배포일|문서번호|
#   제목|내용|해당품목|목적|해당국가|의견마감일|발효일|국내수출품목|
#   관련부서|주간보고|구분|품목
COL = {
    '담당자':      1,
    '순번':        2,
    '중요도':      3,
    '통보유형':    4,
    '통보국':      5,
    '배포일':      6,
    '문서번호':    7,
    '제목':        8,
    '내용':        9,
    '해당품목':   10,
    '목적':       11,
    '해당국가':   12,
    '의견마감일': 13,
    '발효일':     14,
    '국내수출품목': 15,
    '관련부서':   16,
    '주간보고':   17,
    '구분':       18,
    '품목':       19,
    '검토메모':   20,  # reviewer notes
}

YELLOW_FILL = PatternFill('solid', fgColor='FFFF00')
LIME_FILL   = PatternFill('solid', fgColor='CCFF99')
NO_FILL     = PatternFill('none')

# Fields the tool writes (skips pre-filled identification fields)
WRITABLE_FIELDS = [
    '중요도', '제목', '내용', '해당품목', '목적', '해당국가',
    '의견마감일', '발효일', '국내수출품목', '관련부서', '주간보고', '구분', '품목',
]


def _detect_col_map(ws) -> dict:
    """
    Read the header row to detect actual column positions.
    Returns COL-compatible dict; falls back to hardcoded COL if headers not found.
    """
    detected = {}
    for cell in ws[1]:
        if cell.value is None:
            continue
        name = str(cell.value).strip()
        if name in COL:
            detected[name] = cell.column
    # Use detected mapping if at least half the expected columns were found
    if len(detected) >= len(COL) // 2:
        return {**COL, **detected}
    return dict(COL)


def _get_month_sheet(wb, target_month: str = None):
    """
    Return the correct month sheet from the workbook.
    If target_month is given (e.g., '26.4월'), use that.
    Otherwise, auto-detect by current month.
    """
    if target_month:
        for name in wb.sheetnames:
            if target_month in name:
                return wb[name]

    # Auto-detect: find the sheet matching the current year/month
    now = datetime.now()
    year_suffix = str(now.year)[2:]  # '26' from 2026
    month_str = f'{year_suffix}.{now.month}월'
    for name in wb.sheetnames:
        if month_str in name or name.startswith(month_str):
            return wb[name]

    # Fallback: use the first data sheet (skip the manual sheet)
    for name in wb.sheetnames:
        if '매뉴얼' not in name and '월' in name:
            return wb[name]

    return None


def _normalize_doc_number(doc_num: str) -> str:
    """Normalize document number for comparison (strip spaces, upper case)."""
    return re.sub(r'\s+', '', doc_num).upper()


def find_row(wb, doc_number: str, target_month: str = None):
    """
    Find the Excel row matching the given document number.

    Returns (worksheet, row_index, base_date, col_map) or (None, None, None, COL).
    base_date is the 배포일 from the matched row (used for date calculations).
    col_map is the detected column-name→index mapping for this sheet.
    """
    ws = _get_month_sheet(wb, target_month)
    if ws is None:
        return None, None, None, dict(COL)

    col_map = _detect_col_map(ws)
    needle = _normalize_doc_number(doc_number)
    doc_col = col_map['문서번호']
    date_col = col_map['배포일']

    for row in ws.iter_rows(min_row=2):
        cell = row[doc_col - 1]
        if cell.value is None:
            continue
        cell_val = _normalize_doc_number(str(cell.value))
        # Handle joint notifications: 'G/SPS/N/BDI/149,G/SPS/N/KEN/358,...'
        # Match if needle is any of the IDs in the cell
        cell_ids = [_normalize_doc_number(x) for x in re.split(r'[,;]', cell_val)]
        if needle in cell_ids or needle == cell_val:
            base_date = None
            date_cell = row[date_col - 1]
            if date_cell.value:
                from date_engine import parse_excel_date
                base_date = parse_excel_date(date_cell.value)
            return ws, cell.row, base_date, col_map

    return None, None, None, col_map


def write_fields(
    ws,
    row_idx: int,
    fields: dict,
    uncertain_fields: list,
    is_non_english: bool = False,
    col_map: dict = None,
):
    """
    Write computed fields to the matched Excel row.
    - Skips individual cells that already have a non-empty value (Korean already entered).
    - Applies yellow fill to uncertain fields.
    - Applies lime fill to 제목 and 내용 if source is non-English.
    """
    if col_map is None:
        col_map = COL
    for field_name in WRITABLE_FIELDS:
        if field_name not in fields:
            continue
        col_idx = col_map.get(field_name)
        if col_idx is None:
            continue

        cell = ws.cell(row=row_idx, column=col_idx)

        # Skip cells that already have a non-empty value
        if cell.value not in (None, ''):
            continue

        value = fields[field_name]
        if value is None:
            continue

        cell.value = value
        cell.fill = NO_FILL

        # Apply uncertainty highlighting
        if field_name in uncertain_fields:
            cell.fill = YELLOW_FILL
        elif is_non_english and field_name in ('제목', '내용'):
            cell.fill = LIME_FILL

    # Write reviewer notes to column 20 if there are flags
    if uncertain_fields:
        memo_col = col_map.get('검토메모', COL.get('검토메모'))
        if memo_col:
            note_cell = ws.cell(row=row_idx, column=memo_col)
            if note_cell.value in (None, ''):
                note_cell.value = '검토 필요: ' + ', '.join(uncertain_fields)

    return True


def load_and_process(excel_path: str, doc_number: str, fields: dict,
                     uncertain_fields: list, is_non_english: bool = False,
                     target_month: str = None) -> tuple:
    """
    High-level function: open workbook, find row, write fields, save.

    Returns (success: bool, error_msg: str, row_idx: int|None)
    """
    try:
        # Rolling backup — overwritten each run so it does not accumulate
        shutil.copy2(excel_path, excel_path + '.sps_bak')

        wb = load_workbook(excel_path)
        ws, row_idx, base_date, col_map = find_row(wb, doc_number, target_month)
        if ws is None or row_idx is None:
            return False, f'문서번호 {doc_number}을(를) Excel에서 찾을 수 없습니다.', None

        # Record row count before writing — must not change
        row_count_before = ws.max_row

        write_fields(ws, row_idx, fields, uncertain_fields, is_non_english, col_map)

        if ws.max_row != row_count_before:
            return False, 'Excel 행 수가 변경되어 저장을 중단했습니다. 백업 파일(.sps_bak)을 확인하세요.', None

        wb.save(excel_path)
        return True, '', row_idx

    except PermissionError:
        return False, 'Excel 파일이 다른 프로그램에서 열려 있습니다. 닫고 다시 시도해주세요.', None
    except Exception as e:
        return False, str(e), None


def get_base_date(excel_path: str, doc_number: str, target_month: str = None):
    """Get only the 배포일 for a given document number (for date calculations)."""
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        _, _, base_date, _ = find_row(wb, doc_number, target_month)
        return base_date
    except Exception:
        return None
