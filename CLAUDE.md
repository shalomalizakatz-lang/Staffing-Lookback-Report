# Highland Care Center — Staffing Report Automation

## What This Project Does

Automates the monthly staffing report for Highland Care Center (Jamaica, Queens).

The back office sends a monthly **lookback report** as an `.xlsx` file. Historically, staff manually read through it and re-entered CNA/LPN/RN hours into a separate weekly breakdown table. This script eliminates that manual step entirely by reading the payroll data directly from the pivot cache embedded in the lookback file.

**One command in, clean Excel report out.**

---

## How to Run

```bash
python generate_staffing_report.py <lookback_file.xlsx> [output_file.xlsx]
```

**Example:**
```bash
python generate_staffing_report.py July_Lookback.xlsx Highland_Staffing_July_2026.xlsx
```

If no output filename is given, it auto-names the file using the month detected from the data (e.g. `Highland_Staffing_July_2026.xlsx`).

---

## What It Outputs

A 3-sheet Excel workbook:

| Sheet | Contents |
|---|---|
| **Staffing Report** | Weekly tables (one per paycheck date) with CNA/LPN/RN/HHA rows. Payroll hours auto-populated. Yellow cells = manual inputs (census, agency hours). All % and HPPD formula-driven. |
| **Agency Log** | Structured input table for agency invoice data. Fill this out after receiving invoices; the staffing report sheet references it. |
| **Reference Notes** | Documents exactly which earning codes and job categories map to each position. |

---

## Data Source & Key Logic

The lookback `.xlsx` from the back office contains an Excel pivot table whose **cache** holds the full employee-level payroll detail — including job title, earning code, hours, and pay date. The script reads this cache directly (no manual pivot expansion needed).

### Position Mapping

| Report Position | Payroll Job Categories Included |
|---|---|
| **CNA** | `CNA` |
| **LPN** | `LPN`, `LPN - Wound Care` |
| **RN** | `RN`, `RN - Supervisor / Unit Manager`, `RN - Supervisor / Unit Manager (NA)` |
| **HHA** | Not in payroll — agency only |

**Excluded from RN count (by convention):** ADON, MDS/RNAC, Director of Nursing, RN - Wound Care. These are admin/clinical roles not counted toward floor nursing HPPD.

### Earning Code Rules

| Column | Earning Codes Used |
|---|---|
| **Total Hours** | `REGHRLY-Regular`, `OT-Overtime`, `SALARY` |
| **OT Hours** | `OT-Overtime` only — **Holiday OT (`HOLOT`) is excluded** per existing convention |

### Facility Filter

Only cost centers beginning with `20020315` (Highland Care) are included.
Achieve (20020738) is automatically excluded.

---

## What Still Requires Manual Input

| Field | Where to Enter |
|---|---|
| **Census** (per week) | Yellow cell in each week header on Staffing Report sheet |
| **Agency Hours** (CNA/LPN/RN/HHA per week) | Agency Log sheet — one row per position per week |
| **Agency OT Hours** | Same — Agency Log sheet |

**Agency hours are the next automation target.** Invoices come in separately and currently require going invoice-by-invoice. Future work: PDF/Excel invoice parser that populates the Agency Log automatically.

---

## File Structure

```
/staffing_automation/
├── CLAUDE.md                        ← you are here
├── generate_staffing_report.py      ← main script
└── [lookback files dropped here]    ← input
```

---

## Validation Notes (from initial build — June 2026)

Extracted values were compared against 4 weeks of manually-entered data:
- CNA total hours: exact match on all 4 weeks
- LPN total hours: exact match on all 4 weeks
- RN total hours: exact match on 3 of 4 weeks; 0.5h rounding on one (negligible)
- CNA/LPN OT hours: exact match on all 4 weeks
- RN OT: $0 across all weeks (no floor RN overtime in June)

---

## Common Tasks for Claude Code

**"Run the report on this month's lookback"**
→ Drop the `.xlsx` file here, run the script, done.

**"Add a monthly summary tab"**
→ Modify `build_staffing_report()` in the script to add a summary sheet aggregating all weeks.

**"Flag weeks where LPN OT% exceeds 20%"**
→ Add conditional formatting logic in `_build_staffing_sheet()`.

**"Parse these agency invoices and populate the Agency Log"**
→ New function: reads invoice PDFs/Excel, extracts position + hours + week, writes to Agency Log sheet.

**"We added a new agency position (e.g. Restorative Aide)"**
→ Add the job category string to the appropriate set in the constants at the top of the script.

**"The back office changed the field name in the lookback"**
→ Update the `field_by_name.get(...)` strings in `extract_from_pivot_cache()`.
