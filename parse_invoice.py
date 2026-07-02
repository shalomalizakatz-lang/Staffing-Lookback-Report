"""
Agency invoice PDF parser — multi-format.

Supported formats (auto-detected):
  A  Ageless Skye style     "Week starting MM/DD/YYYY" + date/pos/name/qty/rate/amt lines
  B  Bayan Global style      "DESCRIPTION QTY RATE AMOUNT" header + "Name POS hrs rate amt"
  C  County Staffing style   "WeekendDate: MM/DD/YYYY" + 3-line employee blocks
  D  Per-shift style         "Classification Shift Units Bill Rate Period" header
  E  Week-range style        "Week Job Description Hours Rate Amount" header (position unknown)

Returns: (data, agency_name, error, unresolved)
  data:       {payroll_week_date: {pos: {total, ot}}}
  agency_name: str
  error:       str or None
  unresolved: list of {week, hours, description} items that couldn't be mapped to a position
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

# Map description fragments → position
POSITION_KEYWORDS = {
    'CNA': ['cna', 'certified nurse aide', 'nurse aide', 'sick cna'],
    'LPN': ['lpn', 'licensed practical nurse', 'practical nurse', 'sick lpn'],
    'RN':  ['rn', 'registered nurse', 'sick rn', 'rn - supervisor'],
    'HHA': ['hha', 'home health aide'],
}

def _pos_from_description(desc: str) -> str | None:
    d = desc.lower()
    for pos, keywords in POSITION_KEYWORDS.items():
        if any(k in d for k in keywords):
            return pos
    return None

def _to_float(s: str) -> float:
    try:
        return float(re.sub(r'[,$]', '', s))
    except (ValueError, TypeError):
        return 0.0

def _extract_text_pages(pdf_path: str) -> list[str]:
    if not _PDF_OK:
        raise RuntimeError("PyPDF2 not installed — run: pip install PyPDF2")
    reader = PdfReader(pdf_path)
    return [page.extract_text() or '' for page in reader.pages]

def _clean_lines(text: str) -> list[str]:
    return [l.strip() for l in text.splitlines() if l.strip()]

_SKIP_LINE_RE = re.compile(
    r'^(PO Box|INVOICE|BILL TO|DATE|TERMS|Please|Thank|Invoice|Employee|Description|'
    r'Reg Hrs|OT Hrs|OT Rate|Total|Amount|Net|Sub Total|Balance|Payment|Page \d|'
    r'P\.?\s*O\.?|Week|Classification|Shift|Units|Bill Rate|Period|Highland Care|'
    r'CNA\s*[-–]|LPN\s*[-–]|RN\s*[-–]|HHA\s*[-–]|Sick\s+(CNA|LPN|RN)|'
    r'Dietary|Certified Nurse|Licensed Practical|Registered Nurse|'
    r'DUE UPON|RECEIPT|Powered by|'
    r'\d{1,2}/\d{1,2}/\d{2,4}|\$)',
    re.I
)
_COMPANY_RE = re.compile(r'\b(LLC|Inc\.?|Corp\.?|Staffing|Recruiters|Companion|Resources)\b', re.I)

def _agency_name_from_lines(lines: list[str]) -> str:
    candidates = []
    for line in lines:
        cleaned = re.sub(r'^Page\s+\d+\s+of\s+\d+\s*', '', line, flags=re.I).strip()
        if not cleaned or len(cleaned) < 5:
            continue
        if _SKIP_LINE_RE.match(cleaned):
            continue
        if re.match(r'^\d', cleaned):
            continue
        # Skip address-style lines (city, state zip)
        if re.search(r',\s*[A-Z]{2}\s+[-–]?\s*\d{5}', cleaned):
            continue
        if re.search(r'\b(St|Ave|Blvd|Rd|Dr|Ln|Suite|PMB)\b.*\d', cleaned, re.I):
            continue
        candidates.append(cleaned)

    # Prefer lines with strong company name markers
    for c in candidates:
        if _COMPANY_RE.search(c):
            return c
    # Fall back to first reasonable candidate
    for c in candidates[:5]:
        if len(c) > 8:
            return c
    return 'Unknown Agency'

def _week_start_from_date(dt: datetime) -> datetime:
    """Return Sunday of the week containing dt."""
    return dt - timedelta(days=dt.weekday() + 1) if dt.weekday() != 6 else dt

def _nearest_payroll_date(week_start: datetime, payroll_dts: list[datetime]) -> str:
    best = min(payroll_dts, key=lambda d: abs((d - week_start).days))
    return best.strftime('%Y-%m-%d')

# ─── Format A: Ageless Skye ───────────────────────────────────────────────────
# "Week starting MM/DD/YYYY" headers + DATE POS DESCRIPTION QTY RATE AMOUNT

_WEEK_STARTING_RE = re.compile(r'week\s+starting\s+(\d{1,2}/\d{1,2}/\d{4})', re.I)
_LINE_ITEM_RE = re.compile(
    r'(\d{1,2}/\d{1,2}/\d{4})\s+'
    r'(CNA|LPN|RN|HHA)'
    r'(.*?)'
    r'[)\s]+([\d,]+\.?\d*)'
    r'\s+[\d,]+\.?\d*'
    r'\s+[\d,]+\.?\d*'
    r'\s*$',
    re.I
)

def _parse_format_a(pages: list[str]) -> tuple[dict, str]:
    results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    agency_name = _agency_name_from_lines(_clean_lines(pages[0]))
    _FOOTER_RE = re.compile(r'balance\s+due|staffing\s+services|invoice|bill\s+to|terms|due\s+date', re.I)
    current_week = None

    for text in pages:
        raw = text.splitlines()
        lines = []
        for line in raw:
            s = line.strip()
            if not s:
                continue
            is_footer = bool(_FOOTER_RE.search(s))
            if (lines and not is_footer
                    and not re.match(r'\d{1,2}/\d{1,2}/\d{4}', s)
                    and not re.match(r'week\s+starting', s, re.I)
                    and not re.match(r'(CNA|LPN|RN|HHA)\b', s, re.I)
                    and len(s) < 60):
                lines[-1] += ' ' + s
            else:
                lines.append(s)

        for line in lines:
            m = _WEEK_STARTING_RE.search(line)
            if m:
                try:
                    current_week = datetime.strptime(m.group(1), '%m/%d/%Y').strftime('%Y-%m-%d')
                except ValueError:
                    pass
                continue
            if current_week is None:
                continue
            m = _LINE_ITEM_RE.match(line)
            if not m:
                continue
            pos = m.group(2).upper()
            desc = m.group(3)
            hours = _to_float(m.group(4))
            is_ot = bool(re.search(r'\bOT\b', desc, re.I))
            results[current_week][pos]['total'] += hours
            if is_ot:
                results[current_week][pos]['ot'] += hours

    return results, agency_name


# ─── Format B: Bayan Global ───────────────────────────────────────────────────
# "DESCRIPTION QTY RATE AMOUNT" header
# Rows: "Firstname Lastname POS  hours  rate  amount"
# Date range appears as loose text e.g. "10/20 - 10/26"

_BAYAN_HEADER_RE = re.compile(r'DESCRIPTION\s+QTY\s+RATE\s+AMOUNT', re.I)
_DATE_RANGE_RE   = re.compile(r'(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s*[-–]\s*(\d{1,2}/\d{1,2}(?:/\d{2,4})?)')
_BAYAN_ITEM_RE   = re.compile(
    r'(.+?)\s+(CNA|LPN|RN|HHA)\s+([\d.]+)\s+[\d.]+\s+[\d,.]+\s*$', re.I
)
_INVOICE_DATE_RE = re.compile(r'DATE\s+(\d{1,2}/\d{1,2}/\d{4})', re.I)

def _parse_format_b(pages: list[str]) -> tuple[dict, str]:
    results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    all_text = '\n'.join(pages)
    lines = _clean_lines(all_text)
    agency_name = _agency_name_from_lines(lines)

    # Find invoice date as fallback week reference
    invoice_dt = None
    m = _INVOICE_DATE_RE.search(all_text)
    if m:
        try:
            invoice_dt = datetime.strptime(m.group(1), '%m/%d/%Y')
        except ValueError:
            pass

    # Find date range for the week
    week_key = None
    for line in lines:
        m = _DATE_RANGE_RE.search(line)
        if m:
            # Try to parse the end date of range as the week end
            end_str = m.group(2)
            # If year missing, use invoice year
            parts = end_str.split('/')
            if len(parts) == 2 and invoice_dt:
                end_str = f"{end_str}/{invoice_dt.year}"
            try:
                end_dt = datetime.strptime(end_str, '%m/%d/%Y')
                week_start = end_dt - timedelta(days=6)
                week_key = week_start.strftime('%Y-%m-%d')
                break
            except ValueError:
                pass

    if week_key is None and invoice_dt:
        # Fall back: use week containing invoice date
        week_start = _week_start_from_date(invoice_dt)
        week_key = week_start.strftime('%Y-%m-%d')

    if week_key is None:
        return results, agency_name

    in_items = False
    for line in lines:
        if _BAYAN_HEADER_RE.search(line):
            in_items = True
            continue
        if not in_items:
            continue
        if re.match(r'(BALANCE\s+DUE|\$[\d,]+\.?\d*\s*$)', line, re.I):
            break
        m = _BAYAN_ITEM_RE.match(line)
        if m:
            pos = m.group(2).upper()
            hours = _to_float(m.group(3))
            results[week_key][pos]['total'] += hours

    return results, agency_name


# ─── Format C: County Staffing ────────────────────────────────────────────────
# "WeekendDate: MM/DD/YYYY" + repeating column header + 3-line employee blocks:
#   Name (Lastname, Firstname)
#   Description (CNA - Certified Nurse Aide)
#   Reg Hrs, $Rate, OT Hrs, $OT Rate, $Total  (each on own line)

_WEEKEND_DATE_RE = re.compile(r'weekenddate[:\s]+(\d{1,2}/\d{1,2}/\d{4})', re.I)
_COUNTY_HEADER_TOKENS = {'Employee', 'Description', 'Reg Hrs', 'Rate', 'OT Hrs', 'OT Rate', 'Total'}
_NAME_RE  = re.compile(r'^[A-Z][a-zA-Z\-]+,\s+[A-Za-z]')   # Lastname, Firstname
_MONEY_RE = re.compile(r'^\$[\d,]+\.?\d*$')
_HOURS_RE = re.compile(r'^\d+\.?\d*$')

def _parse_format_c(pages: list[str]) -> tuple[dict, str]:
    results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    all_lines = [l for page in pages for l in _clean_lines(page)]
    agency_name = _agency_name_from_lines(all_lines)

    current_week = None

    for text in pages:
        lines = _clean_lines(text)
        i = 0
        while i < len(lines):
            line = lines[i]

            # Week header
            m = _WEEKEND_DATE_RE.search(line)
            if m:
                try:
                    # WeekendDate is the week ENDING date; week starts 6 days earlier
                    end_dt = datetime.strptime(m.group(1), '%m/%d/%Y')
                    week_start = end_dt - timedelta(days=6)
                    current_week = week_start.strftime('%Y-%m-%d')
                except ValueError:
                    pass
                i += 1
                continue

            # Skip column header rows
            if line in _COUNTY_HEADER_TOKENS or line in ('Employee', 'Please Pay'):
                i += 1
                continue

            # Employee block: name line
            if current_week and _NAME_RE.match(line) and i + 2 < len(lines):
                # Next meaningful lines: description, then numbers
                desc_line = lines[i + 1] if i + 1 < len(lines) else ''
                pos = _pos_from_description(desc_line)

                # Collect the following numeric lines (reg_hrs, rate, ot_hrs, ot_rate, total)
                nums = []
                j = i + 2
                while j < len(lines) and len(nums) < 5:
                    tok = lines[j]
                    if _HOURS_RE.match(tok) or _MONEY_RE.match(tok):
                        nums.append(tok)
                        j += 1
                    else:
                        break

                if pos and len(nums) >= 3:
                    reg_hrs = _to_float(nums[0])
                    ot_hrs  = _to_float(nums[2]) if len(nums) > 2 else 0.0
                    results[current_week][pos]['total'] += reg_hrs + ot_hrs
                    results[current_week][pos]['ot']    += ot_hrs
                    i = j
                    continue

            i += 1

    return results, agency_name


# ─── Format D: Per-shift ──────────────────────────────────────────────────────
# Header: "Classification Shift Units Bill Rate Period"
# Under "Employee: Name", rows: "$Amount POS shift-time $Rate Date Units"

_PERSHIFT_HEADER_RE = re.compile(r'Classification\s+Shift\s+Units\s+Bill\s+Rate\s+Period', re.I)
_EMPLOYEE_RE        = re.compile(r'^Employee:\s+(.+)$', re.I)
_PERSHIFT_ROW_RE    = re.compile(
    r'\$([\d,]+\.?\d*)\s+'          # amount
    r'(CNA|LPN|RN|HHA)\s+'          # position
    r'.+?\s+'                         # shift time (lazy)
    r'\$([\d,]+\.?\d*)\s+'           # rate
    r'(\d{1,2}/\d{1,2}/\d{4})\s+'   # date
    r'([\d.]+)',                       # units (hours)
    re.I
)

def _parse_format_d(pages: list[str]) -> tuple[dict, str]:
    results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    all_lines = [l for page in pages for l in _clean_lines(page)]
    agency_name = _agency_name_from_lines(all_lines)

    for text in pages:
        lines = _clean_lines(text)
        in_items = False
        for line in lines:
            if _PERSHIFT_HEADER_RE.search(line):
                in_items = True
                continue
            if not in_items:
                continue
            m = _PERSHIFT_ROW_RE.match(line)
            if not m:
                continue
            pos   = m.group(2).upper()
            date_str = m.group(4)
            hours = _to_float(m.group(5))
            try:
                dt = datetime.strptime(date_str, '%m/%d/%Y')
                week_start = _week_start_from_date(dt)
                week_key = week_start.strftime('%Y-%m-%d')
            except ValueError:
                continue
            results[week_key][pos]['total'] += hours

    return results, agency_name


# ─── Format E: Week-range style ──────────────────────────────────────────────
# Header: "Week Job Description Hours Rate Amount"
# Row:    "$Amount $Rate Hours DateRange JobDesc"
# Position often missing — flagged as unresolved

_WEEKRANGE_HEADER_RE = re.compile(r'Week\s+Job\s+Description\s+Hours\s+Rate\s+Amount', re.I)
_WEEKRANGE_ROW_RE    = re.compile(
    r'\$([\d,]+\.?\d*)\s+'             # amount
    r'\$([\d.]+)\s+'                   # rate
    r'([\d.]+)\s+'                     # hours
    r'(\d{1,2}/\d{1,2}/\d{4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{4})\s*'  # date range
    r'(.*)',                             # job description (may be empty or "OT")
    re.I
)

def _parse_format_e(pages: list[str]) -> tuple[dict, str, list]:
    results   = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    unresolved = []
    all_lines = [l for page in pages for l in _clean_lines(page)]
    agency_name = _agency_name_from_lines(all_lines)

    for text in pages:
        in_items = False
        for line in _clean_lines(text):
            if _WEEKRANGE_HEADER_RE.search(line):
                in_items = True
                continue
            if not in_items:
                continue
            m = _WEEKRANGE_ROW_RE.match(line)
            if not m:
                continue
            hours    = _to_float(m.group(3))
            start_str = m.group(4)
            job_desc  = m.group(6).strip()
            is_ot    = bool(re.search(r'\bOT\b', job_desc, re.I))

            try:
                week_start = datetime.strptime(start_str, '%m/%d/%Y')
                week_key   = week_start.strftime('%Y-%m-%d')
            except ValueError:
                continue

            pos = _pos_from_description(job_desc)
            if pos:
                results[week_key][pos]['total'] += hours
                if is_ot:
                    results[week_key][pos]['ot'] += hours
            else:
                unresolved.append({
                    'week': week_key,
                    'hours': hours,
                    'ot': is_ot,
                    'description': job_desc or '(blank)',
                    'rate': _to_float(m.group(2)),
                })

    return results, agency_name, unresolved


# ─── Format F: Meridian / per-date ───────────────────────────────────────────
# Header: "Date Worked  Service Type/Healthcare Professional  SHIFT  HOURS  RATE  AMOUNT"
# Row:    "MM/DD/YYYY   Name   POSITION shift-desc   hours   $rate   $amount"

_MERIDIAN_HEADER_RE = re.compile(r'Date\s+Worked\s+Service\s+Type', re.I)
_MERIDIAN_ROW_RE    = re.compile(
    r'(\d{1,2}/\d{1,2}/\d{4})\s+'       # date
    r'.+?\s+'                              # name (lazy)
    r'(CNA|LPN|RN|HHA)\b'                # position
    r'.+?'                                # shift desc (lazy)
    r'([\d.]+)\s+'                        # hours
    r'\$[\d,]+\.?\d*\s+'                  # rate
    r'\$[\d,]+\.?\d*\s*$',               # amount
    re.I
)

def _parse_format_f(pages: list[str]) -> tuple[dict, str]:
    results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))
    all_lines = []
    for text in pages:
        all_lines.extend(_clean_lines(text))
    agency_name = _agency_name_from_lines(all_lines)

    in_items = False
    for line in all_lines:
        if _MERIDIAN_HEADER_RE.search(line):
            in_items = True
            continue
        if not in_items:
            continue
        m = _MERIDIAN_ROW_RE.match(line)
        if not m:
            continue
        date_str = m.group(1)
        pos      = m.group(2).upper()
        hours    = _to_float(m.group(3))
        try:
            dt = datetime.strptime(date_str, '%m/%d/%Y')
            week_start = _week_start_from_date(dt)
            week_key   = week_start.strftime('%Y-%m-%d')
        except ValueError:
            continue
        results[week_key][pos]['total'] += hours

    return results, agency_name


# ─── Format detection ─────────────────────────────────────────────────────────

def _detect_format(pages: list[str]) -> str:
    all_text = '\n'.join(pages)
    if _WEEK_STARTING_RE.search(all_text):
        return 'A'
    if _BAYAN_HEADER_RE.search(all_text):
        return 'B'
    if _WEEKEND_DATE_RE.search(all_text):
        return 'C'
    if _PERSHIFT_HEADER_RE.search(all_text):
        return 'D'
    if _WEEKRANGE_HEADER_RE.search(all_text):
        return 'E'
    if _MERIDIAN_HEADER_RE.search(all_text):
        return 'F'
    return 'UNKNOWN'


# ─── Public API ───────────────────────────────────────────────────────────────

def parse_invoice_pdf(pdf_path: str):
    """
    Parse any supported agency invoice PDF.
    Returns (data, agency_name, error, unresolved)
      data:       {week_start_date: {pos: {total, ot}}}
      agency_name: str
      error:       str or None
      unresolved: list — items whose position could not be determined
    """
    try:
        pages = _extract_text_pages(pdf_path)
    except Exception as e:
        return None, None, str(e), []

    pages = [p for p in pages if p.strip()]
    if not pages:
        return None, None, 'PDF appears to be empty or image-based (no text layer)', []

    fmt = _detect_format(pages)

    unresolved = []
    try:
        if fmt == 'A':
            data, name = _parse_format_a(pages)
        elif fmt == 'B':
            data, name = _parse_format_b(pages)
        elif fmt == 'C':
            data, name = _parse_format_c(pages)
        elif fmt == 'D':
            data, name = _parse_format_d(pages)
        elif fmt == 'E':
            data, name, unresolved = _parse_format_e(pages)
        elif fmt == 'F':
            data, name = _parse_format_f(pages)
        else:
            return None, None, (
                'Invoice format not recognized. Supported formats: Ageless Skye, '
                'Bayan Global, County Staffing, per-shift, and week-range invoices.'
            ), []
    except Exception as e:
        return None, None, f'Parse error ({fmt}): {e}', []

    if not data and not unresolved:
        return None, name, 'No line items found — check that this is a staffing invoice', []

    # Round
    result = {}
    for week, positions in data.items():
        result[week] = {
            pos: {'total': round(h['total'], 2), 'ot': round(h['ot'], 2)}
            for pos, h in positions.items()
        }

    return result, name, None, unresolved


def match_to_payroll_weeks(invoice_data: dict, payroll_week_dates: list[str]) -> dict:
    """
    Map invoice week-start dates to payroll pay dates.
    The pay date is typically 7–14 days after the work week starts.
    Returns {payroll_date: {pos: {total, ot}}} merged across matching invoice weeks.
    """
    payroll_dts = [datetime.strptime(d, '%Y-%m-%d') for d in payroll_week_dates]
    mapped = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))

    for week_start_str, positions in invoice_data.items():
        week_start = datetime.strptime(week_start_str, '%Y-%m-%d')
        # Pay date should be 0–21 days after work week starts
        best_pay_date = None
        best_delta = None
        for pay_dt in payroll_dts:
            delta = (pay_dt - week_start).days
            if 0 <= delta <= 21:
                if best_delta is None or delta < best_delta:
                    best_delta = delta
                    best_pay_date = pay_dt.strftime('%Y-%m-%d')
        if best_pay_date is None:
            best_pay_date = _nearest_payroll_date(week_start, payroll_dts)

        for pos, hours in positions.items():
            mapped[best_pay_date][pos]['total'] += hours['total']
            mapped[best_pay_date][pos]['ot']    += hours['ot']

    result = {}
    for pay_date, positions in mapped.items():
        result[pay_date] = {
            pos: {'total': round(h['total'], 2), 'ot': round(h['ot'], 2)}
            for pos, h in positions.items()
        }
    return result
