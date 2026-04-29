# MRC Business Review Dashboard

Interactive financial dashboard and PowerPoint generator for monthly business reviews.

---

## Quick Start

```bash
# 1. Activate virtual environment
.venv\Scripts\activate          # Windows

# 2. Place Excel files in any subfolder (see naming convention below)

# 3. Launch the dashboard
streamlit run mbr_dashboard.py
```

Open your browser at `http://localhost:8501`.

---

## System Components

| File | Purpose |
|------|---------|
| `mbr_dashboard.py` | Streamlit web dashboard — all interactive views |
| `generate_full_mbr_stacked.py` | Automated PowerPoint slide generator |
| `config_loader.py` | YAML config reader (dot-notation access) |
| `config/config.yaml` | All row/column indices, colors, thresholds |

---

## Excel File Naming Convention

```
YYMMDD_Project_Budget_Analysis_MonthName_vX.Y.xlsx
Example: 260331_Project_Budget_Analysis_March_v1.6.xlsx
```

- The dashboard scans the entire script folder recursively for files matching this pattern.
- Multiple versions of the same month are supported — the highest version wins.
- Files are auto-sorted by year and month.

### Required Sheets

| Sheet | Used For |
|-------|---------|
| `Summary TL` | Net Sales, EBIT, BU-level NS and EBIT data |
| `Order Intake` | OI totals by BU and individual project list |
| `WIP & BL` | Work-in-Progress project list |
| `EBIT Calc.` | YTD EBIT %, Budget EBIT %, Net Fees % rates |
| `Ext. Prod.` | Monthly production history per project (optional — Project History tab) |
| `Gross Margin` | ABNS (Awarded But Not Signed) projects (optional) |

---

## Business Units

| Code | Full Name | Color |
|------|-----------|-------|
| ENG | Engineering | `#0EA5E9` (blue) |
| MC | Management Consultancy | `#10B981` (green) |
| T&SI | Technology & Smart Infrastructures | `#FAB611` (amber) |
| NUC | Nuclear | `#93C572` (light green) |

---

## Global Controls

### Business Unit Radio (top of page)
Applies to every tab simultaneously.
- **Company Total** — shows aggregated company-wide figures
- **ENG / MC / T&SI / NUC** — filters all views to that single BU

### Sidebar — Period
- **Year** selector → filters available months to that year
- **Compare months** checkbox → multi-select up to any number of months (shown side-by-side or as sub-tabs if > 3)
- **Compare years** checkbox → loads the same month(s) from a second year for YoY comparison

### Sidebar — Generate Report
Runs `generate_full_mbr_stacked.py` as a subprocess for the selected month, then offers a PPTX download and optional email delivery.

---

## Tab Reference

---

### 1. Net Sales

**Data source:** `Summary TL` sheet — columns 1–16 (pandas), rows configured via `Excel_Mapping.Rows`.

| Row config key | Default Excel row | Content |
|---|---|---|
| `Categories_NS_kTL` | 14 (0-indexed) | Month labels (Jan, Feb … Dec, target) |
| `Slide2_Contract` | 15 | Contract NS — cumulative YTD (kTL) |
| `Slide2_WP` | 16 | Weighted Proposals NS — cumulative YTD (kTL) |
| `Slide2_WO` | 17 | Weighted Opportunities NS — cumulative YTD (kTL) |

**BU view data source:** Same sheet, columns `BU_NS_Start`–`BU_NS_End` (default cols 52–65). Each BU has three rows (Order / Offer / Opp). Row numbers are configured via `BU_NS_ENG_Order` etc.

**Metric cards:**
- **Monthly** = `YTD[current month] − YTD[previous month]` (derived by scanning backwards for the nearest month column)
- **YTD** = raw cumulative value at the current month column
- **EOY** = value at the December column (or the year column if December is absent)

**Chart:** Stacked bar — each tier (Contract / WP / WO for Company Total; Order / Offer / Opp for BU view) is a stack segment. Values are converted to per-month deltas before plotting so bars show monthly amounts, not cumulative totals.

**ABNS section (Company Total only):** Reads `Gross Margin` sheet. Filters rows where `Status = ABNS` and `Item = GR`, sums the `GM_TOTAL` column (default col 25).

---

### 2. EBIT

**Data source:** `Summary TL` sheet — columns 1–15.

| Row config key | Default Excel row | Content |
|---|---|---|
| `Categories_EBIT_kTL` | 36 | Month labels |
| `Slide4_Contract` | 37 | Contract EBIT — cumulative YTD (kTL) |
| `Slide4_Contract_WP` | 38 | Contract + WP EBIT — cumulative YTD (kTL) |
| `Slide4_Contract_WP_WO` | 39 | Full Outlook EBIT — cumulative YTD (kTL) |

**BU view data source:** Same sheet, columns `BU_EBIT_Start`–`BU_EBIT_End` (default cols 52–65). Each BU has a Total row and a Pct (EBIT %) row.

| BU EBIT config key | Default row |
|---|---|
| `BU_EBIT_ENG` (Total) | 100 (1-indexed in config, stored as `value − 1` at read time) |
| `BU_EBIT_P_ENG` (Pct) | 108 (0-indexed) |
| Same pattern for MC / TSI / NUC | |

**Company Total metric cards:** Monthly and YTD values derived same way as Net Sales (delta from cumulative arrays).

**BU view metric cards (compact row):** Monthly | YTD | EOY | EBIT % — all four shown in one row to save space.

**EBIT % section (Company Total):**
- Source: `EBIT Calc.` sheet, row headers in row 2 (pandas row 1)
- `EBIT_Pct_Actual` → config `EBIT_Pct_Actual` (default row 66, 0-indexed)
- `EBIT_Pct_Budget` → config `EBIT_Pct_Budget` (default row 67)
- `Net_Fees_Pct` → config `Net_Fees_Pct` (default row 68)
- Values are raw decimals (0.084 = 8.4%); multiplied by 100 for display.

**BU EBIT % trend:** Same column range as BU EBIT Total, using the Pct row. Shown as a line chart below the bar chart.

---

### 3. Order Intake

**Data source:** `Order Intake` sheet.

**Row offset detection:** The dashboard scans from the bottom of the sheet upward to find the categories row, looking for a cell containing a 4-digit year followed by "Jan"/"Feb" two columns to the right. This auto-corrects for files with varying numbers of project rows.

| Data point | Location |
|---|---|
| Categories row | `Categories_OI_kTL` (default row 543) + detected offset |
| ENG monthly total | `Slide6_ENG` (default row 544) + offset |
| MC monthly total | `Slide6_MC` (default row 545) + offset |
| T&SI monthly total | `Slide6_TSI` (default row 546) + offset |
| NUC monthly total | `Slide6_NUC` (default row 547) + offset |

**Month column selection:** `OI_Col_Base + month_index` where `OI_Col_Base = 6` and `month_index` = 0 for January, 1 for February, etc. (parsed from the filename date prefix).

**Project list:** Rows 2 through `Slide6_ENG + offset`. Each row: project name (col 3), client (col 2), BU (col 4), monthly value (col `OI_Col_Base + month_index`). Rows with zero value, "total"/"subtotal" in name or client are skipped.

**Metric cards (Company Total):**
- Monthly = value at `oi_idx` column per BU
- YTD = sum of monthly columns from `MON_OI_Base` (default 2) through `oi_idx`

---

### 4. WIP (Work in Progress)

**Data source:** `WIP & BL` sheet, starting at `WIP_DataStartRow` (default row 3).

**Row filtering rules** (mirrors the PPTX reporting code exactly):
1. Only rows where `Type` (col 3) = `"GROSS FEES"` are kept
2. Rows where `Client` (col 4) is blank or contains "akkuyu" are excluded
3. Nuclear BU rows are excluded from the main list
4. Entries where `WIP_TL` is between `WIP_NEG_THR` (−1,000 TL) and `WIP_MIN_TL` (1,000,000 TL) are dropped — keeps meaningful positives and significant negatives only

**Key columns read:**

| Config key | Default col (0-indexed) | Content |
|---|---|---|
| `WIP_Col_Name` | 2 | Project name |
| `WIP_Col_Type` | 3 | Cost type (must be "GROSS FEES") |
| `WIP_Col_Client` | 4 | Client name |
| `WIP_Col_BU` | 10 | Business unit |
| `WIP_Col_OrigCurrency` | 11 | Original currency code |
| `WIP_Col_TotalInvoiceOC` | 93 | Total invoice in original currency |
| `WIP_Col_TotalProdOC` | 94 | Total production in original currency |
| `WIP_Col_WIP_TL` | 98 | WIP balance in TL |

**Metric cards:**
- **Total WIP ≥1M** — sum of displayed (filtered) projects
- **Grand Total WIP** — sum of ALL non-zero GROSS FEES entries before the 1M threshold filter (matches the reporting code grand total)

**Sort order:** Positive WIP sorted descending, negative WIP sorted ascending (most negative last), appended after positives.

---

### 5. Project History

**Data source:** `Ext. Prod.` sheet (loaded separately via `load_margin_data()`).

**Sheet structure detected dynamically:**
- Header row: found by scanning for a row containing both "PROJECT CODE" and "PROJECT NAME"
- Month labels: row immediately above the header containing month names
- Date row: one row below the month label row (used to disambiguate year)

**What is read per project row:**
- Monthly production values for each month column where the field label = "PRODUCTION"
- Special columns: `BL (Year Prod)`, `BO (End Year Total)`, `Next Year`, `TD (To Date)` — detected by regex scanning

**Project key:** `project_code|currency` — a project with the same code but different currencies is kept as separate entries.

**Type breakdown:** All cost types (GROSS FEES, SUBCON 1, ASSOCIATE 1, etc.) are accumulated separately in `type_data`. The "Margin (All Types)" view sums all types together.

**WIP TL history view:** For each selected month file, the dashboard looks up the project name (case-insensitive) in the WIP project list and records the WIP TL balance.

**Local BU filter:** Automatically locks to the global BU radio selection. If global = "Company Total", the filter is free; if global = "ENG", only ENG projects appear and the filter is disabled.

**File-to-file comparison:** Uses the current month file and the previous month file (by position in the sorted file list) to show Δ Month and Δ EoY Forecast per cost type in the breakdown table.

---

### 6. Scenarios

Four sub-tabs:

#### BU Comparison
- Reads `bu_ns` and `bu_ebit` from `Summary TL` for all selected months
- Shows NS by BU (total and tier breakdown) and EBIT by BU as bar charts

#### Top 10 Clients
- **Net Sales (Production)** source: `Ext. Prod.` sheet — sums all monthly GROSS FEES production YTD per client (excludes special columns: BL, BO, Next Year, TD)
- **WIP** source: `WIP & BL` sheet — sums positive WIP TL per client

#### Portfolio Mix
- Interactive BU toggle (on/off per BU)
- Computes selected vs excluded NS and EBIT amounts and share percentages
- Source: same `bu_ns` / `bu_ebit` arrays as BU Comparison

#### What-If Analysis
- **Baseline:** Monthly Contract NS and monthly Contract EBIT from `Summary TL`; EBIT % from `EBIT Calc.` sheet row 66
- **Levers:** NS growth % (slider −50 → +100), additional wins (kTL, additive), EBIT margin adjustment (pp, −20 → +20)
- **Formula:**
  ```
  Projected NS   = Actual NS × (1 + growth%/100) + additional_wins
  Projected EBIT% = Actual EBIT% + margin_adj_pp/100
  Projected EBIT  = Projected NS × Projected EBIT%
  ```
- Waterfall chart shows NS movement (Growth bar and New Wins bar only appear when non-zero)

---

### 7. Executive Summary

Four sections selectable via radio buttons:

#### Overview
- **NS KPIs:** Monthly, YTD, EOY forecast, Target — from `Summary TL` Contract NS row
- **EBIT KPIs:** Same structure from Contract EBIT row
- **Margins & OI:** EBIT % actual/budget/Net Fees from `EBIT Calc.` sheet; OI monthly and YTD from `Order Intake` sheet
- **Business Unit Snapshot table:** Per-BU NS monthly, NS YTD, NS Target, % vs Target, EBIT monthly, EBIT YTD — source is `bu_ns` / `bu_ebit` arrays from `Summary TL`. NS Target reads the column at index `BU_Target_Col_Index` (default = 1 within the BU NS column range, i.e. the annual plan column)
- **MoM / YoY Δ:** Shown when two months are loaded. Compares monthly and YTD values between the two periods

#### YTD vs Target
- Reads the column whose category label contains "target" (case-insensitive) from NS and EBIT arrays
- Progress bar = `YTD value / Target value`
- Coverage % shown alongside abbreviated YTD and Target numbers (hover for full precision)
- Margin Rates from `EBIT Calc.` shown at the bottom

#### Rolling 12-Month Net Sales
- Builds a 12-month window ending on the currently selected month
- **File selection per slot:**
  - Current year → exact month file, falls back to latest available file for that year
  - Previous years → December file only (fully finalized annual data), falls back to latest available
- Monthly value = `YTD[month] − YTD[month−1]` (cumulative-to-delta conversion)
- Views: Company Total (Contract NS only) | By Business Unit (all tiers summed) | By Project / By Client (from `Ext. Prod.` sheet, GROSS FEES only)

#### Pipeline Health
- Source: `bu_ns` arrays — Order / Offer / Opp values at the current month column
- Year-end target: sum of the three tiers at `BU_Target_Col_Index`
- Metrics per BU: YTD pipeline total, year-end target, % of target, Secured (Order%) and Secured+Offer%

---

## Number Formats

All monetary values are in **kTL** (thousands of Turkish Lira) unless stated otherwise. The dashboard uses Turkish locale formatting:
- Thousands separator: `.` (period)
- Decimal separator: `,` (comma)
- Example: `1.234,56` = 1,234.56 kTL

WIP values are in raw TL (not kTL) and use abbreviated format (`human_tl`): M for millions, B for billions.

---

## Data Loading & Caching

```
load_month(filepath)          → @st.cache_data(ttl=3600)
  Reads: Summary TL, Order Intake, WIP & BL, EBIT Calc., Gross Margin

load_margin_data(filepath)    → @st.cache_resource(ttl=3600)
  Reads: Ext. Prod. sheet only

merge_margin_data(months)     → @st.cache_resource(ttl=3600)
  Merges load_margin_data() results across selected months

fetch_all_data(months_tuple)  → @st.cache_data(ttl=3600)
  Calls load_month() for each month in the selection
```

Cache TTL is 1 hour. Use the **Refresh Files** button in the sidebar to clear cache immediately after updating an Excel file.

---

## Configuration Reference (`config/config.yaml`)

### Excel_Mapping.Rows (0-indexed unless noted)

| Key | Default | Meaning |
|-----|---------|---------|
| `Categories_NS_kTL` | 14 | NS month label row in Summary TL |
| `Slide2_Contract` | 15 | NS Contract YTD row |
| `Slide2_WP` | 16 | NS Weighted Proposals YTD row |
| `Slide2_WO` | 17 | NS Weighted Opportunities YTD row |
| `Categories_EBIT_kTL` | 36 | EBIT month label row |
| `Slide4_Contract` | 37 | EBIT Contract YTD row |
| `Slide4_Contract_WP` | 38 | EBIT Contract+WP YTD row |
| `Slide4_Contract_WP_WO` | 39 | EBIT Full Outlook YTD row |
| `Categories_OI_kTL` | 543 | OI categories row (base; offset applied at runtime) |
| `Slide6_ENG` | 544 | OI ENG monthly total row |
| `Slide6_MC` | 545 | OI MC monthly total row |
| `Slide6_TSI` | 546 | OI T&SI monthly total row |
| `Slide6_NUC` | 547 | OI NUC monthly total row |
| `WIP_DataStartRow` | 3 | First data row in WIP & BL sheet |
| `BU_NS_ENG_Order` | 19 | BU NS ENG Order row (1-indexed) |
| `BU_EBIT_ENG` | 100 | BU EBIT ENG Total row (1-indexed) |
| `BU_EBIT_P_ENG` | 108 | BU EBIT ENG Pct row (0-indexed) |
| `EBIT_Pct_Actual` | 66 | EBIT Calc. actual % row (0-indexed) |
| `EBIT_Pct_Budget` | 67 | EBIT Calc. budget % row (0-indexed) |
| `Net_Fees_Pct` | 68 | EBIT Calc. Net Fees % row (0-indexed) |

### Excel_Mapping.Columns (0-indexed)

| Key | Default | Meaning |
|-----|---------|---------|
| `WIP_Col_Name` | 2 | Project name in WIP & BL |
| `WIP_Col_Type` | 3 | Cost type (must be "GROSS FEES") |
| `WIP_Col_Client` | 4 | Client name |
| `WIP_Col_BU` | 10 | Business unit code |
| `WIP_Col_OrigCurrency` | 11 | Original currency |
| `WIP_Col_WIP_TL` | 98 | WIP balance in TL |
| `BU_NS_Start` | 52 | First column of BU NS block |
| `BU_NS_End` | 66 | Last column (exclusive) of BU NS block |
| `OI_kTL_Start` | 4 | OI categories start column |

### Excel_Mapping.Offsets

| Key | Default | Meaning |
|-----|---------|---------|
| `MON_NS_Base` | 3 | First month column index in NS array (col 1 + 3 = Jan) |
| `MON_EBIT_Base` | 3 | Same for EBIT |
| `MON_OI_Base` | 2 | First month column in OI BU arrays |
| `OI_Col_Base` | 6 | Column offset for OI monthly project values |

### Dashboard Settings

| Key | Default | Meaning |
|-----|---------|---------|
| `KTL_Decimal_Places` | 2 | Decimal places for kTL display |
| `WIP_Neg_Threshold` | −1000 | Minimum negative WIP TL to include |
| `Chart_Top_N_Projects` | 30 | Max bars in project charts |
| `Client_Top_N` | 10 | Top N clients in Scenarios |
| `Chart_Label_Min_Pct` | 0.03 | Minimum segment size to show inside-bar label |
| `BU_Target_Col_Index` | 1 | Column index within BU NS block for annual target |

### WIP_Table

| Key | Default | Meaning |
|-----|---------|---------|
| `WIP_Min_TL` | 1,000,000 | Minimum absolute WIP TL to display |

---

## Folder Structure

```
MRC Finance/
├── config/
│   └── config.yaml                    # All row/column/threshold settings
├── <YYMMDD>_<MonthName>/              # One folder per month (any depth)
│   └── YYMMDD_Project_Budget_Analysis_MonthName_vX.Y.xlsx
├── mbr_dashboard.py                   # Streamlit dashboard
├── generate_full_mbr_stacked.py       # PowerPoint generator
├── config_loader.py                   # YAML loader utility
└── README.md                          # This file
```

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| Tab shows no data | Excel file not found or sheet missing | Check filename pattern; verify sheet names match required list |
| OI data misaligned | Row offset detection failed | A warning appears in the OI tab; check that the OI categories row has a 4-digit year + "Jan"/"Feb" nearby |
| WIP Grand Total differs from table sum | Expected — Grand Total includes sub-threshold projects | No fix needed; this matches the PPTX reporting logic |
| EBIT % shows 0 | `EBIT Calc.` sheet absent or row indices wrong | Check sheet name and update `EBIT_Pct_Actual` row in config |
| Project History empty | `Ext. Prod.` sheet missing or header row not found | Verify sheet exists; check that "PROJECT CODE" and "PROJECT NAME" appear in a header row within the first 10 rows |
| Cache showing stale data | 1-hour TTL not expired | Click **Refresh Files** in sidebar |
| Numbers look wrong after Excel update | Old cache still active | Click **Refresh Files** in sidebar |
