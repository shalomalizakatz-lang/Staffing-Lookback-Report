"""
Highland Care Center - Staffing Report Automation
Reads the monthly lookback .xlsx (from back office) and generates a complete
staffing report with CNA/LPN/RN hours auto-populated from the payroll pivot cache.
Agency data is pulled from Agency Log sheet if present, otherwise left blank.

Usage:
    python generate_staffing_report.py <lookback_file.xlsx> [output_file.xlsx]
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, date
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00

# ─── Constants ───────────────────────────────────────────────────────────────

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# Earning codes that count toward productive hours
INCLUDE_EARNS = {'REGHRLY-Regular', 'OT-Overtime', 'HOLOT-Holiday OT', 'SALARY'}
# Earning codes that count as OT (excludes Holiday OT per manual convention)
OT_EARNS = {'OT-Overtime'}

# Job category → position mapping
CNA_JOBS = {'CNA'}
LPN_JOBS = {'LPN', 'LPN - Wound Care'}
RN_JOBS = {
    'RN',
    'RN - Supervisor / Unit Manager',
    'RN - Supervisor / Unit Manager (NA)',
}

# Position display order
POSITIONS = ['CNA', 'LPN', 'RN', 'HHA']

# Colors
COLORS = {
    'header_bg':  '1F3864',   # dark navy
    'header_font': 'FFFFFF',  # white
    'week_bg':    '2E75B6',   # medium blue
    'total_bg':   'D6E4F0',   # light blue
    'input_bg':   'FFF2CC',   # yellow — manual input cells
    'alt_row':    'EBF3FA',   # very light blue for alternating rows
    'white':      'FFFFFF',
    'border':     'B8CCE4',
}

THIN = Side(style='thin', color=COLORS['border'])
MED  = Side(style='medium', color='1F3864')
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_MED_TOP = Border(left=THIN, right=THIN, top=MED, bottom=THIN)

# ─── Pivot cache extraction ───────────────────────────────────────────────────

def extract_from_pivot_cache(xlsx_path: str):
    """
    Reads the Highland Care payroll pivot cache embedded in the .xlsx file.
    Returns: dict {pay_date_str: {position: {total, ot}}}
    """
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            files = z.namelist()

            # Find the cache that contains Highland Care data (exclude _rels files)
            cache_def_files = sorted([f for f in files
                                      if 'pivotCacheDefinition' in f and '_rels' not in f])
            cache_rec_files = sorted([f for f in files
                                      if 'pivotCacheRecords' in f and '_rels' not in f])

            if not cache_def_files:
                return None, "No pivot cache found in file"

            # Try each cache — find the one with Job Category field
            for def_file, rec_file in zip(cache_def_files, cache_rec_files):
                with z.open(def_file) as f:
                    def_root = ET.parse(f).getroot()

                fields = {}
                for i, cf in enumerate(def_root.findall(f'.//{{{NS}}}cacheField')):
                    items = []
                    for si in cf.findall(f'.//{{{NS}}}sharedItems/{{{NS}}}s'):
                        items.append(si.get('v'))
                    for si in cf.findall(f'.//{{{NS}}}sharedItems/{{{NS}}}d'):
                        items.append(si.get('v'))
                    fields[i] = {'name': cf.get('name'), 'items': items}

                # Check if this cache has Job Category
                has_job_cat = any(f['name'] == 'Job Category' for f in fields.values())
                if not has_job_cat:
                    continue

                # Map field indices by name
                field_by_name = {f['name']: i for i, f in fields.items()}

                cc_field_idx    = field_by_name.get('Cost Center (as of Calculation Entry Moment)')
                job_field_idx   = field_by_name.get('Job Category')
                earn_field_idx  = field_by_name.get('Earning')
                pd_field_idx    = field_by_name.get('Payment Date or Reversal Date')

                if None in (cc_field_idx, job_field_idx, earn_field_idx, pd_field_idx):
                    continue

                cost_centers = fields[cc_field_idx]['items']
                job_cats     = fields[job_field_idx]['items']
                earnings_map = fields[earn_field_idx]['items']
                pay_dates    = fields[pd_field_idx]['items']

                # Only keep June 2026 pay dates
                target_dates = {}
                for idx, pd_str in enumerate(pay_dates):
                    if pd_str and '2026-06' in pd_str:
                        target_dates[idx] = pd_str[:10]

                if not target_dates:
                    continue

                # Find field positions in records (sequential by index in field list)
                # Fields layout: 0..4 are strings, then cc, job, earn, dates, paydate, ...nums
                # Detect positions by counting elements
                with z.open(rec_file) as f:
                    rec_root = ET.parse(f).getroot()

                results = defaultdict(lambda: defaultdict(lambda: {'total': 0.0, 'ot': 0.0}))

                # Determine field positions by parsing a sample record
                sample_recs = list(rec_root.findall(f'{{{NS}}}r'))[:5]
                # Build position mapping by counting fields in def vs records
                num_fields = len(fields)

                for rec in rec_root.findall(f'{{{NS}}}r'):
                    elements = list(rec)
                    if len(elements) < num_fields:
                        continue

                    try:
                        def get_val(pos):
                            el = elements[pos]
                            return el.get('v')

                        cc_val    = get_val(cc_field_idx)
                        job_val   = get_val(job_field_idx)
                        earn_val  = get_val(earn_field_idx)
                        pd_val    = get_val(pd_field_idx)

                        # Hours is always near the end — find the numeric fields
                        # Heuristic: hours = second-to-last numeric field before rate
                        # Fields: ..., n(0), n(hours), n(amount), n(rate)
                        # Count backward from end to find hours
                        n_fields = [i for i, el in enumerate(elements)
                                    if el.tag == f'{{{NS}}}n']
                        if len(n_fields) < 2:
                            continue
                        hours_pos = n_fields[-3] if len(n_fields) >= 3 else n_fields[-2]
                        hours = float(elements[hours_pos].get('v') or 0)

                    except (IndexError, TypeError, ValueError):
                        continue

                    if cc_val is None or job_val is None or earn_val is None or pd_val is None:
                        continue

                    try:
                        cc_idx_v   = int(cc_val)
                        job_idx_v  = int(job_val)
                        earn_idx_v = int(earn_val)
                        pd_idx_v   = int(pd_val)
                    except (TypeError, ValueError):
                        continue

                    # Only Highland Care (20020315) cost centers
                    cc = cost_centers[cc_idx_v] if cc_idx_v < len(cost_centers) else ''
                    if not cc.startswith('20020315'):
                        continue

                    # Only target pay dates
                    if pd_idx_v not in target_dates:
                        continue

                    pay_date = target_dates[pd_idx_v]
                    job = job_cats[job_idx_v] if job_idx_v < len(job_cats) else ''
                    earn = earnings_map[earn_idx_v] if earn_idx_v < len(earnings_map) else ''

                    # Map to position
                    if job in CNA_JOBS:
                        pos = 'CNA'
                    elif job in LPN_JOBS:
                        pos = 'LPN'
                    elif job in RN_JOBS:
                        pos = 'RN'
                    else:
                        continue

                    if earn not in INCLUDE_EARNS or hours == 0:
                        continue

                    results[pay_date][pos]['total'] += hours
                    if earn in OT_EARNS:
                        results[pay_date][pos]['ot'] += hours

                return results, None

        return None, "No matching pivot cache found"

    except Exception as e:
        return None, str(e)


# ─── Styles ──────────────────────────────────────────────────────────────────

def make_font(bold=False, size=11, color='000000', italic=False):
    return Font(name='Arial', bold=bold, size=size, color=color, italic=italic)

def make_fill(hex_color):
    return PatternFill('solid', start_color=hex_color, fgColor=hex_color)

def center():
    return Alignment(horizontal='center', vertical='center')

def apply_header(ws, row, col, value, span=1, bg=None, fg='000000', bold=True,
                 size=11, border=True, wrap=False):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name='Arial', bold=bold, size=size, color=fg)
    if bg:
        cell.fill = make_fill(bg)
    cell.alignment = Alignment(horizontal='center', vertical='center',
                               wrap_text=wrap)
    if border:
        cell.border = BORDER_THIN
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    return cell


# ─── Workbook builder ────────────────────────────────────────────────────────

def build_staffing_report(payroll_data: dict, agency_data: dict,
                          weeks: list, census_map: dict,
                          lookback_wb=None, output_path='staffing_report_output.xlsx'):
    """
    Builds the complete staffing report workbook.

    payroll_data: {date_str: {pos: {total, ot}}}
    agency_data:  {date_str: {pos: {total, ot}}}
    weeks:        [(pay_date_str, label_str), ...]  sorted
    census_map:   {date_str: int}  — can be empty (will leave as input)
    """
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet
    wb.calculation.calcMode = 'auto'  # force Excel to recalculate all formulas on open

    _build_staffing_sheet(wb, payroll_data, agency_data, weeks, census_map)
    _build_agency_log(wb, weeks)
    _build_notes(wb, weeks)

    wb.save(output_path)
    return output_path


def _build_staffing_sheet(wb, payroll_data, agency_data, weeks, census_map):
    ws = wb.create_sheet('Staffing Report')
    ws.sheet_view.showGridLines = False

    # Column widths
    col_widths = {
        1: 18,   # A: Position
        2: 14,   # B: Total Hours
        3: 12,   # C: OT Hours
        4: 10,   # D: OT %
        5: 16,   # E: Total Agency Hrs
        6: 14,   # F: Agency OT Hrs
        7: 11,   # G: Agency %
        8: 11,   # H: Agency OT %
        9: 10,   # I: HPPD
        10: 2,   # J: spacer
        11: 12,  # K: Census
    }
    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    row = 1

    # Title
    ws.row_dimensions[row].height = 28
    cell = ws.cell(row=row, column=1, value='Highland Care Center — Monthly Staffing Report')
    cell.font = Font(name='Arial', bold=True, size=14, color=COLORS['header_font'])
    cell.fill = make_fill(COLORS['header_bg'])
    cell.alignment = Alignment(horizontal='left', vertical='center',
                               indent=1, wrap_text=False)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)

    month_label = ''
    if weeks:
        dt = datetime.strptime(weeks[0][0], '%Y-%m-%d')
        month_label = dt.strftime('%B %Y')
    ws.cell(row=row, column=1).value = (
        f'Highland Care Center — Staffing Report  |  {month_label}'
    )

    row += 1

    # Notes row
    ws.row_dimensions[row].height = 14
    note_cell = ws.cell(row=row, column=1,
        value='🟡 Yellow cells = manual input  |  All other cells auto-calculated from payroll data')
    note_cell.font = Font(name='Arial', italic=True, size=9, color='595959')
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
    row += 1

    # ── One block per week ──
    for week_date, week_label in weeks:
        payroll = payroll_data.get(week_date, {})
        agency  = agency_data.get(week_date, {})
        census  = census_map.get(week_date)

        ws.row_dimensions[row].height = 6
        row += 1

        # Week header
        ws.row_dimensions[row].height = 22
        week_cell = ws.cell(row=row, column=1,
                            value=f'  {week_label}')
        week_cell.font = Font(name='Arial', bold=True, size=12,
                              color=COLORS['header_font'])
        week_cell.fill = make_fill(COLORS['week_bg'])
        week_cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)

        # Census cell (right side, same row)
        census_label = ws.cell(row=row, column=10, value='Census:')
        census_label.font = Font(name='Arial', bold=True, size=10,
                                 color=COLORS['header_font'])
        census_label.fill = make_fill(COLORS['week_bg'])
        census_label.alignment = Alignment(horizontal='right', vertical='center')

        census_cell = ws.cell(row=row, column=11,
                              value=census if census else None)
        census_cell.font = make_font(bold=True, size=11)
        if not census:
            census_cell.fill = make_fill(COLORS['input_bg'])
        census_cell.alignment = center()
        census_cell.border = BORDER_THIN
        census_ref = census_cell.coordinate

        row += 1

        # Column headers
        ws.row_dimensions[row].height = 30
        headers = ['Position', 'Total Hours', 'OT Hours', 'OT %',
                   'Agency Hours', 'Agency OT Hrs', 'Agency %', 'Agency OT %', 'HPPD']
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=c, value=h)
            cell.font = Font(name='Arial', bold=True, size=9,
                             color=COLORS['header_font'])
            cell.fill = make_fill(COLORS['header_bg'])
            cell.alignment = Alignment(horizontal='center', vertical='center',
                                       wrap_text=True)
            cell.border = BORDER_THIN
        row += 1

        # Data rows
        pos_rows = {}
        for i, pos in enumerate(POSITIONS):
            ws.row_dimensions[row].height = 18
            is_alt = (i % 2 == 1)
            row_bg = COLORS['alt_row'] if is_alt else COLORS['white']

            p_data = payroll.get(pos, {'total': 0.0, 'ot': 0.0})
            a_data = agency.get(pos, {'total': 0.0, 'ot': 0.0})

            # A: Position
            c = ws.cell(row=row, column=1, value=pos)
            c.font = make_font(bold=True, size=11)
            c.fill = make_fill(row_bg)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = BORDER_THIN

            # B: Total Hours (payroll — auto)
            b_val = round(p_data['total'], 2) if p_data['total'] else 0
            c = ws.cell(row=row, column=2, value=b_val if b_val else None)
            c.font = make_font(size=11)
            c.fill = make_fill(row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            c.number_format = '#,##0.00'
            total_ref = c.coordinate

            # C: OT Hours (payroll — auto)
            c_val = round(p_data['ot'], 2) if p_data['ot'] else 0
            c = ws.cell(row=row, column=3, value=c_val if c_val else None)
            c.font = make_font(size=11)
            c.fill = make_fill(row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            c.number_format = '#,##0.00'
            ot_ref = c.coordinate

            # D: OT % — blank for HHA (agency-only, no payroll denominator)
            if pos == 'HHA' or not b_val:
                ot_pct_val = None
            else:
                ot_pct_val = f'=IF({total_ref}=0,"",{ot_ref}/{total_ref})'
            c = ws.cell(row=row, column=4, value=ot_pct_val)
            c.font = make_font(size=11)
            c.fill = make_fill(row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            if ot_pct_val:
                c.number_format = '0.0%'

            # E: Agency Hours (manual input — yellow if empty)
            a_val = round(a_data['total'], 2) if a_data.get('total') else None
            c = ws.cell(row=row, column=5, value=a_val)
            c.font = make_font(size=11)
            c.fill = make_fill(COLORS['input_bg'] if not a_val else row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            c.number_format = '#,##0.00'
            agency_ref = c.coordinate

            # F: Agency OT (manual input — yellow if empty)
            ao_val = round(a_data['ot'], 2) if a_data.get('ot') else None
            c = ws.cell(row=row, column=6, value=ao_val)
            c.font = make_font(size=11)
            c.fill = make_fill(COLORS['input_bg'] if not ao_val else row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            c.number_format = '#,##0.00'
            agency_ot_ref = c.coordinate

            # G: Agency % of payroll total (blank for HHA — no payroll denominator)
            if pos == 'HHA' or not b_val:
                agency_pct_val = None
            else:
                agency_pct_val = f'=IF({agency_ref}="","",IF({total_ref}=0,"",{agency_ref}/{total_ref}))'
            c = ws.cell(row=row, column=7, value=agency_pct_val)
            if agency_pct_val:
                c.number_format = '0.0%'
            c.font = make_font(size=11)
            c.fill = make_fill(row_bg)
            c.alignment = center()
            c.border = BORDER_THIN

            # H: Agency OT % — blank until agency hours are entered
            if a_val or ao_val:
                agt_ot_val = f'=IF({agency_ref}=0,"",{agency_ot_ref}/{agency_ref})'
            else:
                agt_ot_val = None
            c = ws.cell(row=row, column=8, value=agt_ot_val)
            c.font = make_font(size=11)
            c.fill = make_fill(row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            if agt_ot_val:
                c.number_format = '0.0%'

            # I: HPPD — blank until census is entered
            combined = f'({total_ref}+IF(ISNUMBER({agency_ref}),{agency_ref},0))'
            c = ws.cell(row=row, column=9,
                value=f'=IF({census_ref}="","",{combined}/{census_ref})')
            c.font = make_font(size=11)
            c.fill = make_fill(row_bg)
            c.alignment = center()
            c.border = BORDER_THIN
            c.number_format = '0.000'

            pos_rows[pos] = row
            row += 1

        # Total row
        ws.row_dimensions[row].height = 20
        c = ws.cell(row=row, column=1, value='Total')
        c.font = make_font(bold=True, size=11)
        c.fill = make_fill(COLORS['total_bg'])
        c.alignment = center()
        c.border = Border(left=MED, right=THIN, top=MED, bottom=MED)

        first_pos_row = pos_rows['CNA']
        last_pos_row  = pos_rows['HHA']

        for col in range(2, 10):
            col_letter = get_column_letter(col)
            c = ws.cell(row=row, column=col)
            c.fill = make_fill(COLORS['total_bg'])
            c.alignment = center()

            if col == 2:  # Total hours
                c.value = f'=SUM({col_letter}{first_pos_row}:{col_letter}{last_pos_row})'
                c.number_format = '#,##0.00'
                total_total_ref = c.coordinate
            elif col == 3:  # OT hours
                c.value = f'=SUM({col_letter}{first_pos_row}:{col_letter}{last_pos_row})'
                c.number_format = '#,##0.00'
                ot_total_ref = c.coordinate
            elif col == 4:  # OT %
                c.value = f'=IF({total_total_ref}=0,"",{ot_total_ref}/{total_total_ref})'
                c.number_format = '0.0%'
            elif col == 5:  # Agency total
                c.value = f'=SUM({col_letter}{first_pos_row}:{col_letter}{last_pos_row})'
                c.number_format = '#,##0.00'
                agency_total_ref = c.coordinate
            elif col == 6:  # Agency OT total
                c.value = f'=SUM({col_letter}{first_pos_row}:{col_letter}{last_pos_row})'
                c.number_format = '#,##0.00'
                agency_ot_total_ref = c.coordinate
            elif col == 7:  # Agency %
                c.value = f'=IF({agency_total_ref}=0,"",IF({total_total_ref}=0,"",{agency_total_ref}/{total_total_ref}))'
                c.number_format = '0.0%'
            elif col == 8:  # Agency OT %
                c.value = f'=IF({agency_total_ref}=0,"",IF({agency_ot_total_ref}=0,"",{agency_ot_total_ref}/{agency_total_ref}))'
                c.number_format = '0.0%'
            elif col == 9:  # HPPD total
                combined = f'({total_total_ref}+IF(ISNUMBER({agency_total_ref}),{agency_total_ref},0))'
                c.value = f'=IF({census_ref}="","",{combined}/{census_ref})'
                c.number_format = '0.000'

            c.font = make_font(bold=True, size=11)
            left_s = MED if col == 2 else THIN
            right_s = MED if col == 9 else THIN
            c.border = Border(left=left_s, right=right_s, top=MED, bottom=MED)

        row += 1

    # Freeze top 3 rows
    ws.freeze_panes = 'A4'


def _build_agency_log(wb, weeks):
    """Build a structured agency input sheet."""
    ws = wb.create_sheet('Agency Log')
    ws.sheet_view.showGridLines = False

    col_widths = [4, 22, 18, 10, 14, 14, 14]
    cols = ['', 'Week Ending', 'Position', 'Agency Name', 'Reg Hours',
            'OT Hours', 'Notes']
    for i, (w, h) in enumerate(zip(col_widths, cols), 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.row_dimensions[1].height = 26
    c = ws.cell(row=1, column=1, value='Agency Hours Log  —  Enter invoice data here')
    c.font = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    c.fill = make_fill('1F3864')
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.merge_cells('A1:G1')

    ws.row_dimensions[2].height = 14
    c = ws.cell(row=2, column=1,
        value='Position options: CNA, LPN, RN, HHA  |  Add one row per agency per position per week')
    c.font = Font(name='Arial', italic=True, size=9, color='595959')
    ws.merge_cells('A2:G2')

    # Headers
    ws.row_dimensions[3].height = 22
    for col, h in enumerate(cols, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        c.fill = make_fill('2E75B6')
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER_THIN

    # Pre-fill rows for each week × position
    r = 4
    for week_date, week_label in weeks:
        for pos in POSITIONS:
            ws.row_dimensions[r].height = 16
            bg = COLORS['alt_row'] if (r % 2 == 0) else COLORS['white']

            ws.cell(row=r, column=1).fill = make_fill(bg)
            c2 = ws.cell(row=r, column=2, value=week_label.split('(')[0].strip())
            c2.font = make_font(size=10)
            c2.fill = make_fill(bg)
            c2.alignment = Alignment(horizontal='center', vertical='center')
            c2.border = BORDER_THIN

            c3 = ws.cell(row=r, column=3, value=pos)
            c3.font = make_font(size=10)
            c3.fill = make_fill(bg)
            c3.alignment = Alignment(horizontal='center', vertical='center')
            c3.border = BORDER_THIN

            for col in [4, 5, 6, 7]:
                c = ws.cell(row=r, column=col)
                c.fill = make_fill(COLORS['input_bg'])
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = BORDER_THIN
                if col in (5, 6):
                    c.number_format = '#,##0.00'

            r += 1

        # Spacer row between weeks
        ws.row_dimensions[r].height = 6
        for col in range(1, 8):
            ws.cell(row=r, column=col).fill = make_fill('EBF3FA')
        r += 1

    ws.freeze_panes = 'B4'


def _build_notes(wb, weeks):
    ws = wb.create_sheet('Reference Notes')
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 55

    ws.row_dimensions[1].height = 26
    c = ws.cell(row=1, column=1, value='Reference Notes — Auto-population Rules')
    c.font = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    c.fill = make_fill('1F3864')
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.merge_cells('A1:C1')

    notes = [
        ('PAYROLL AUTO-POPULATION', ''),
        ('CNA Hours', 'Job Category = "CNA"  |  Earnings = REGHRLY + OT-Overtime + SALARY'),
        ('LPN Hours', 'Job Category = "LPN" or "LPN - Wound Care"  |  Same earnings'),
        ('RN Hours', 'Job Category = "RN", "RN - Supervisor/Unit Manager", '
                     '"RN - Supervisor/Unit Manager (NA)"  |  Same earnings'),
        ('HHA Hours', 'Not in payroll — agency only, enter in Agency Log'),
        ('OT Hours (payroll)', 'Only "OT-Overtime" code — Holiday OT (HOLOT) excluded per convention'),
        ('', ''),
        ('AGENCY', ''),
        ('Agency Hours', 'Enter in Agency Log sheet → weekly tables pull from there'),
        ('HHA', 'Always agency — enter in Agency Log as HHA'),
        ('', ''),
        ('HPPD', ''),
        ('Formula', '(Payroll Hours + Agency Hours) / Census'),
        ('Census', 'Manual entry — yellow cell in each week header'),
        ('', ''),
        ('COST CENTER FILTER', ''),
        ('Highland Care only', 'Cost centers starting with "20020315" — Achieve (20020738) excluded'),
    ]

    r = 2
    for label, detail in notes:
        ws.row_dimensions[r].height = 18
        is_section = detail == ''
        bg = COLORS['week_bg'] if is_section else (
            COLORS['alt_row'] if r % 2 == 0 else COLORS['white'])

        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(name='Arial', bold=is_section, size=10,
                      color='FFFFFF' if is_section else '000000')
        c.fill = make_fill(bg)
        c.alignment = Alignment(vertical='center', indent=1)
        c.border = BORDER_THIN

        c2 = ws.cell(row=r, column=3, value=detail)
        c2.font = Font(name='Arial', size=10)
        c2.fill = make_fill(bg)
        c2.alignment = Alignment(vertical='center', indent=1, wrap_text=True)
        c2.border = BORDER_THIN
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
        ws.cell(row=r, column=1).fill = make_fill(bg)

        r += 1


# ─── Entry point ─────────────────────────────────────────────────────────────

def main(input_path: str, output_path: str = None):
    print(f"Reading payroll data from: {input_path}")

    payroll_data, err = extract_from_pivot_cache(input_path)
    if err:
        print(f"Error extracting payroll data: {err}")
        sys.exit(1)

    if not payroll_data:
        print("No payroll data found for current month")
        sys.exit(1)

    # Sort weeks
    weeks = []
    for date_str in sorted(payroll_data.keys()):
        if '2026-06' not in date_str:
            continue
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        # Week label: "6/4/2026 (5/24 – 5/30)"
        weeks.append((date_str, dt.strftime('%-m/%-d/%Y')))

    if not weeks:
        print("No June 2026 data found")
        sys.exit(1)

    print(f"Found {len(weeks)} weeks: {[w[0] for w in weeks]}")

    # Print summary
    for date_str, label in weeks:
        d = payroll_data[date_str]
        print(f"\n  {label}:")
        for pos in ['CNA', 'LPN', 'RN']:
            p = d.get(pos, {})
            print(f"    {pos}: {p.get('total', 0):.2f} hrs  OT: {p.get('ot', 0):.2f}")

    if not output_path:
        month = datetime.strptime(weeks[0][0], '%Y-%m-%d').strftime('%B_%Y')
        output_path = f'Highland_Staffing_{month}.xlsx'

    output = build_staffing_report(
        payroll_data=payroll_data,
        agency_data={},  # empty — to be filled from Agency Log
        weeks=weeks,
        census_map={},   # empty — manual input
        output_path=output_path,
    )

    print(f"\nOutput written to: {output}")
    return output


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python generate_staffing_report.py <lookback.xlsx> [output.xlsx]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    main(input_file, output_file)
