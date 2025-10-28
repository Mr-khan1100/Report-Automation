# Discrepancy Automation

**Short:** Automated script to compare Blis, DCM and Celtra reports using a mapping sheet — produces per-placement/day comparison sheets and a summary.

---

## Purpose

This project automates discrepancy reporting across three advertising reports:

- **Blis** (creative-level metrics)
- **DCM** (placement-level metrics) --here any third-party report can be used (adform, sizmik, gimus, etc. with samne column values present in the sheet named 'DCM')
- **Celtra** (placement-level creative delivery metrics)

The script links records using the **Mapping** sheet and produces a consolidated, date-aggregated Excel report that shows where numbers differ between sources (impressions, clicks, etc.). It handles missing sources gracefully and fills missing dates with zeros.

---

## What to place in the input Excel (Template)

You must provide a single Excel workbook containing these sheets (sheet names are _case-insensitive_ but recommended to use exact names below):

### 1) `Mapping` (MANDATORY)

Columns (these column **names must exist** and should not be renamed):

- `Blis Creative ID` _(string / numeric — commas allowed in file but script normalizes)_
- `Celtra Placement ID` _(string — can be empty for some rows)_
- `DCM Placement ID` _(string — can be empty for some rows)_

> All three columns must be present in the sheet even if some values are blank. The mapping connects Blis ↔ Celtra ↔ DCM.

**Sample mapping row**

```
Blis Creative ID, Celtra Placement ID, DCM Placement ID
1617630,22538da5,426439118
1617631,22538da5,426439118
1617632,6471628a,426439118
```

### 2) `Blis` (MANDATORY)

Columns (the script expects these names – include them even if blank):

- `Date` (any parseable date/time)
- `Blis Creative ID`
- `Blis Requested Impression`
- `Blis Shown Impression`
- `Blis clicks`
- `Blis Raw clicks`
- `Blis win count`

Notes: numeric columns will be coerced to numbers and missing values become 0. Date column is normalized to `YYYY-MM-DD`.

### 3) `DCM` (OPTIONAL)

If present, expected columns (include them if you want DCM comparisons):

- `Date`
- `DCM Placement ID`
- `DCM impression`
- `DCM click`
- `DCM invalid Click` _(optional — if not present, related calculations are skipped)_

### 4) `Celtra` (OPTIONAL)

If present, expected columns:

- `Date`
- `Celtra Placement ID`
- `Celtra Requested Impression`
- `Celtra Loaded Impression`
- `Celtra Rendered Impression`
- `Clicks` _(becomes `Celtra Clicks` in output)_

---

## Sample links of report generation from celtra and imply:

**IMPLY :** https://imply.blis.com/pivot/d/c2e0bafd6e99f197ba/Cartographer_for_EMEA
**CELTRA :** https://blismedia.celtra.com/reports/#savedReportSpecId=8f322bea

---

## How the output looks

The script writes a single Excel workbook (same folder as your input) named:

```
discrepancyReport_YYYYMMDD_HHMMSS.xlsx
```

Inside the workbook:

- One sheet per mapping group (named `Group_<group_key>`). A group can be based on DCM id, Celtra id, or a single Blis creative (based on mapping).
  - **If the group key is a DCM id**: the sheet will show **one row per Celtra placement per date** (i.e., Celtra placements are _not_ collapsed into a single cell) — each row contains DCM + Celtra + Blis metrics for that date and that Celtra placement.
  - **If the group key is a Celtra id**: the sheet aggregates all Blis creatives mapped to that Celtra (single row per date).
  - **If the group key is a Blis creative id**: it shows that creative’s daily metrics and mapped placement IDs.
- Each per-group sheet contains the requested mandatory columns (only those present in source are included) and percent comparisons such as:
  - `blis vs DCM impression %`, `blis vs DCM click %`
  - `Blis raw click vs DCM invalid %` _(only if DCM invalid clicks exist)_
  - `blis impression vs celtra loaded %`, `blis click vs Celtra click %`
  - `celtra loaded vs DCM impression %`, `Celtra click vs DCM click %`
- A **Grand Total** row at the bottom of each sheet. When a DCM group produces multiple per-Celtra rows, DCM totals in the Grand Total are taken directly from the DCM source (to avoid double-counting).
- A **Summary** sheet listing each group and whether Blis / DCM / Celtra data were present plus high-level totals and percentages.

---

## Checks & Behavior

- Mapping IDs are **normalized** (commas removed and whitespace trimmed) so `1,613,502` → `1613502`.
- If a creative maps to multiple placements (unexpected), the script writes a `Mapping_Warnings` sheet.
- Missing sheets (`DCM` or `Celtra`) are handled gracefully — related comparisons are skipped and columns are not shown.
- Missing dates in any source are filled with `0` so time series align across sources.
- Safe percent calculation: divide-by-zero yields `NaN` (not infinity) to avoid misleading percentages.

---

## Requirements

- Python 3.8+ recommended
- Libraries:
  - `pandas`
  - `numpy`
  - `openpyxl`

Install via pip:

```bash
pip install pandas numpy openpyxl
```

---

## How to run

1. Place your input workbook (with `Mapping` + `Blis` and optional `DCM` / `Celtra`) inside the project folder.
2. Run the script from terminal / PowerShell (example):

```bash
python descrepancy_automation_script.py descrepancyTemplate.xlsx
```

Output file will be created in the same folder, e.g.:

```
discrepancyReport_20251028_223054.xlsx
```

> Tip: Use a Python virtual environment to keep dependencies isolated.

---

## Troubleshooting

- **`KeyError: 'date'`** — ensure your source sheets have a date column and it’s parseable. Column must be named `Date` or be the first column.
- **All Celtra values zero** — check `Mapping` for correct Celtra IDs and that `Celtra` sheet has `Celtra Placement ID` values that match mapping (script normalizes commas/whitespace).
- **Script exits with usage message** — you ran the script without the input filename. Provide the Excel file path as the first argument.

---

## Optional improvements (ideas)

- Add `--output` argument to choose the output path. (Can be added using `argparse`.)
- Save outputs into an `output/` folder automatically.
- Add logging to a file with `--verbose` flag.
