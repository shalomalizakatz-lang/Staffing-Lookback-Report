"""
Agency invoice PDF parser.
Handles the Ageless Skye / similar format:
  DATE  POSITION  DESCRIPTION  QTY  RATE  AMOUNT
grouped under "Week starting MM/DD/YYYY" headers.
Returns: {week_start_date_str: {position: {total, ot}}}
"""

import re
from datetime import datetime, timedelta
from collections import defaultdict

try:
    from PyPDF2 import PdfReader
    _PDF_OK = True
except ImportError:
    _PDF_OK = False

POSITIONS = {'CNA', 'LPN', 'RN', 'HHA'}

# Matches "Week starting 04/07/2024" or "Week Starting ..."
WEEK_RE = re.compile(r'week\s+starting\s+(\d{1,2}/\d{1,2}/\d{4})', re.IGNORECASE)

# Matches a line-item row:
# 04/07/2024  CNA  CNA/Aniesa Khan  7.75  28.00  1,664.00
# The position token is one of CNA/LPN/RN/HHA (with optional suffix like "- OT")
# Note: description may run directly into qty with no space (e.g. "for 4/6)7.50")
LINE_RE = re.compile(
    r'(\d{1,2}/\d{1,2}/\d{4})\s+'       # date
    r'(CNA|LPN|RN|HHA)'                  # position
    r'(.*?)'                              # description (lazy)
    r'[)\s]+([\d,]+\.?\d*)'             # qty (hours) — allow ) before it
    r'\s+[\d,]+\.?\d*'                   # rate
    r'\s+[\d,]+\.?\d*'                   # amount
    r'\s*$',
    re.IGNORECASE
)

# Also catch lines where description wraps: just the position + name line
# and the numbers come on the next line — handled by joining lines first.

def _extract_text_pages(pdf_path: str) -> list[str]:
    if not _PDF_OK:
        raise RuntimeError("PyPDF2 not installed")
    reader = PdfReader(pdf_path)
    return [page.extract_text() or '' for page in reader.pages]


def _parse_page(text: str) -> dict:
    """Parse one invoice page. Returns {week_start: {pos: {total, ot}}}"""
    results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    current_week = None

    # Join continuation lines: if a line starts with a lowercase letter or
    # nothing useful, attach it to the previous line.
    raw_lines = text.splitlines()
    lines = []
    for line in raw_lines:
        stripped = line.strip()
        if not stripped:
            continue
        # If previous line exists and current looks like a continuation
        # (no date at start, no "Week", not a position, not a footer line), merge it.
        is_footer = bool(re.search(r'balance\s+due|staffing\s+services|invoice|bill\s+to|terms|due\s+date', stripped, re.I))
        if (lines and
                not is_footer and
                not re.match(r'\d{1,2}/\d{1,2}/\d{4}', stripped) and
                not re.match(r'week\s+starting', stripped, re.I) and
                not re.match(r'(CNA|LPN|RN|HHA)\b', stripped, re.I) and
                len(stripped) < 60):
            lines[-1] = lines[-1] + ' ' + stripped
        else:
            lines.append(stripped)

    for line in lines:
        # Check for week header
        m = WEEK_RE.search(line)
        if m:
            try:
                current_week = datetime.strptime(m.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')
            except ValueError:
                pass
            continue

        if current_week is None:
            continue

        # Try to match a line-item
        m = LINE_RE.match(line)
        if not m:
            continue

        pos = m.group(2).upper()
        desc = m.group(3)
        try:
            hours = float(m.group(4).replace(',', ''))
        except ValueError:
            continue

        is_ot = bool(re.search(r'\bOT\b', desc, re.IGNORECASE))

        results[current_week][pos]['total'] += hours
        if is_ot:
            results[current_week][pos]['ot'] += hours

    return results


def parse_invoice_pdf(pdf_path: str):
    """
    Parse an agency invoice PDF.
    Returns (data, agency_name, error)
      data: {week_start_date: {pos: {total, ot}}}
      agency_name: str
      error: str or None
    """
    try:
        pages = _extract_text_pages(pdf_path)
    except Exception as e:
        return None, None, str(e)

    combined = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    agency_name = None

    for text in pages:
        if not text.strip():
            continue

        # Try to extract agency name from first non-empty page
        if agency_name is None:
            # Agency name is on the first non-empty line, but may be preceded
            # by "Page X of Y" (sometimes merged on same line without newline)
            first_lines = [l.strip() for l in text.splitlines() if l.strip()]
            for line in first_lines[:4]:
                # Strip leading "Page N of N" prefix (may be merged with name)
                cleaned = re.sub(r'^Page\s+\d+\s+of\s+\d+\s*', '', line, flags=re.IGNORECASE).strip()
                if cleaned and not cleaned.lower().startswith('po box') and len(cleaned) > 4:
                    agency_name = cleaned
                    break
            if not agency_name:
                agency_name = 'Unknown Agency'

        page_data = _parse_page(text)
        for week, positions in page_data.items():
            for pos, hours in positions.items():
                combined[week][pos]['total'] += hours['total']
                combined[week][pos]['ot']    += hours['ot']

    if not combined:
        return None, agency_name, "No line items found — the PDF may use a different format"

    # Round all values
    result = {}
    for week, positions in combined.items():
        result[week] = {}
        for pos, hours in positions.items():
            result[week][pos] = {
                'total': round(hours['total'], 2),
                'ot':    round(hours['ot'], 2),
            }

    return result, agency_name, None


def match_to_payroll_weeks(invoice_data: dict, payroll_week_dates: list[str]) -> dict:
    """
    Map invoice week-start dates to payroll pay dates.
    Strategy: the pay date is typically 7–14 days after the work week starts.
    Returns {payroll_date: {pos: {total, ot}}} merged across all matching invoice weeks.
    """
    payroll_dts = [datetime.strptime(d, '%Y-%m-%d') for d in payroll_week_dates]
    mapped = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))

    for week_start_str, positions in invoice_data.items():
        week_start = datetime.strptime(week_start_str, '%Y-%m-%d')
        week_end   = week_start + timedelta(days=6)

        # Find the payroll date closest to this work week
        # Pay date should be AFTER the week ends, within ~14 days
        best_pay_date = None
        best_delta = None
        for pay_dt in payroll_dts:
            # Pay date should fall after week start, within 21 days
            delta = (pay_dt - week_start).days
            if 0 <= delta <= 21:
                if best_delta is None or delta < best_delta:
                    best_delta = delta
                    best_pay_date = pay_dt.strftime('%Y-%m-%d')

        if best_pay_date is None:
            # Fall back: nearest payroll date overall
            nearest = min(payroll_dts, key=lambda d: abs((d - week_start).days))
            best_pay_date = nearest.strftime('%Y-%m-%d')

        for pos, hours in positions.items():
            mapped[best_pay_date][pos]['total'] += hours['total']
            mapped[best_pay_date][pos]['ot']    += hours['ot']

    # Round
    result = {}
    for pay_date, positions in mapped.items():
        result[pay_date] = {
            pos: {'total': round(h['total'], 2), 'ot': round(h['ot'], 2)}
            for pos, h in positions.items()
        }
    return result
