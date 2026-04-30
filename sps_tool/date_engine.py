"""
Date calculation engine for WTO SPS notification deadlines.
Resolves explicit dates, formula-based dates, and special values.
"""
import re
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta

MONTH_MAP = {
    'january': 1,  'february': 2,  'march': 3,     'april': 4,
    'may': 5,      'june': 6,      'july': 7,       'august': 8,
    'september': 9,'october': 10,  'november': 11,  'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7,
    'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
    # Spanish months
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5,
    'junio': 6, 'julio': 7, 'agosto': 8, 'septiembre': 9,
    'octubre': 10, 'noviembre': 11, 'diciembre': 12,
    # Portuguese months
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
    'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12,
}

NOT_APPLICABLE_PATTERNS = [
    'not applicable', 'n/a', 'na', 'no aplica', 'não aplicável',
    'does not apply', 'emergency', '-',
]
TBD_PATTERNS = [
    'to be determined', 'tbd', 'not yet determined', 'por determinar',
    'a determinar', 'a ser determinado', 'not specified',
]


def _parse_explicit_date(text: str):
    """
    Try to parse an explicit calendar date from text.
    Returns a date object or None.
    """
    text = text.strip()

    # DD Month YYYY  or  Month DD, YYYY
    patterns = [
        r'(\d{1,2})[/ .](\d{1,2})[/ .](\d{4})',          # DD/MM/YYYY or MM/DD/YYYY
        r'(\d{1,2})\s+(\w+)\s+(\d{4})',                    # 16 March 2026
        r'(\w+)\s+(\d{1,2}),?\s+(\d{4})',                  # March 16, 2026
        r'(\d{4})[/-](\d{2})[/-](\d{2})',                  # 2026-03-16 (ISO)
    ]

    for pat in patterns:
        m = re.search(pat, text)
        if not m:
            continue
        g = m.groups()
        try:
            # Pattern 1: numeric
            if re.match(r'\d+$', g[1]):
                day, month, year = int(g[0]), int(g[1]), int(g[2])
                # Distinguish DD/MM vs MM/DD by value range
                if month > 12:
                    day, month = month, day
                return date(year, month, day)
            # Pattern 2: DD Month YYYY
            if g[1].lower() in MONTH_MAP:
                return date(int(g[2]), MONTH_MAP[g[1].lower()], int(g[0]))
            # Pattern 3: Month DD, YYYY
            if g[0].lower() in MONTH_MAP:
                return date(int(g[2]), MONTH_MAP[g[0].lower()], int(g[1]))
            # Pattern 4: ISO
            return date(int(g[0]), int(g[1]), int(g[2]))
        except (ValueError, KeyError):
            continue
    return None


def _parse_formula(text: str, base_date: date):
    """
    Resolve formulas like '60 days from circulation', '6 months after publication'.
    Returns a date object or None.
    """
    text_lower = text.lower()

    # N days
    m = re.search(r'(\d+)\s+(?:calendar\s+)?days?\s+(?:from|after|following)', text_lower)
    if m:
        return base_date + timedelta(days=int(m.group(1)))

    # N months
    m = re.search(r'(\d+|one|two|three|four|five|six|seven|eight|nine|ten|twelve)\s+months?\s+(?:from|after|following|of)', text_lower)
    if m:
        word_to_num = {'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,
                       'seven':7,'eight':8,'nine':9,'ten':10,'twelve':12}
        raw = m.group(1)
        n = int(raw) if raw.isdigit() else word_to_num.get(raw, 0)
        if n:
            return base_date + relativedelta(months=n)

    # Approximately N days
    m = re.search(r'approximately\s+(\d+)\s+days?', text_lower)
    if m:
        return base_date + timedelta(days=int(m.group(1)))

    return None


def resolve_date(raw_text: str, base_date: date, is_emergency: bool = False) -> tuple:
    """
    Resolve a raw date string from a WTO SPS notification.

    Returns (formatted_date_str, source_expression):
    - formatted_date_str: DD/MM/YYYY string, '추후결정', or '-'
    - source_expression: the original text that was parsed
    """
    if not raw_text:
        raw_text = ''
    text = raw_text.strip()

    # Emergency notifications: comment deadline is always '-'
    if is_emergency and not text:
        return ('-', 'emergency notification')

    text_lower = text.lower()

    # Not applicable / dash
    if any(p in text_lower for p in NOT_APPLICABLE_PATTERNS) or text == '-':
        return ('-', text)

    # To be determined
    if any(p in text_lower for p in TBD_PATTERNS) or not text:
        return ('추후결정', text or 'not stated')

    # Try explicit date first
    explicit = _parse_explicit_date(text)
    if explicit:
        return (explicit.strftime('%d/%m/%Y'), text)

    # Try formula calculation
    if base_date:
        calculated = _parse_formula(text, base_date)
        if calculated:
            return (calculated.strftime('%d/%m/%Y'), text)

    # Cannot resolve — return 추후결정 with a flag
    return ('추후결정', f'[unresolved] {text}')


def parse_excel_date(cell_value) -> date | None:
    """Parse a date from an openpyxl cell value (may be date, datetime, or string)."""
    if cell_value is None:
        return None
    if isinstance(cell_value, date):
        return cell_value
    if hasattr(cell_value, 'date'):
        return cell_value.date()
    # Try string parsing
    if isinstance(cell_value, str):
        d = _parse_explicit_date(cell_value)
        return d
    return None
