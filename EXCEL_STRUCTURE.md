# MRC Finance Excel — Source Data Structure Reference

**Verified against:** `260228_Project_Budget_Analysis_February_v1.2.xlsx`  
**File naming pattern:** `YYMMDD_Project_Budget_Analysis_MonthName_vX.Y.xlsx`  
**Total sheets:** 18 (only 5 are read by the dashboard/generator)

---

## Sheet Overview

| Sheet | Used By | Purpose |
|-------|---------|---------|
| **Summary TL** | Dashboard + PPTX | NS, EBIT, BU totals (in kTL and k€) |
| **Order Intake** | Dashboard + PPTX | Project-level Order Intake by month |
| **Ext. Prod.** | Dashboard (Project History) | Monthly production by project and cost type |
| **WIP & BL** | Dashboard + PPTX | WIP and Backlog by project |
| **Gross Margin** | PPTX only | Project-level margin analysis (kTL) |
| *Other 13 sheets* | — | Intermediate calculations, not read |

---

## Sheet 1: Summary TL

**Shape:** 120 rows × 67 columns (pandas 0-based)

### Column Layout (row 3 = header row for most sections)
| Col index | Label | Notes |
|-----------|-------|-------|
| 1 | 2024 | Full year actual |
| 2 | 2025 | Full year actual |
| 3 | 2026 Target | Annual target |
| 4 | Jan … | Monthly cumulative YTD (col 4–15) |
| 15 | Dec | Last monthly column |

> **Important:** NS/EBIT sections insert a **BL** column at col 4, shifting Jan to col 5.  
> At the NS category label row (row 14), col 4 = `BL`, col 5 = `Jan`.  
> This matches the config comment: `[0]=2025, [1]=2026T, [2]=BL, [3]=Jan`.

### Row Map — left-side sections (cols 1–16)

| Pandas row | Label (col 0) | Config key | Section |
|------------|---------------|------------|---------|
| 1 | *(section header)* Order Intake Progress (kTL) | — | OI |
| 3 | *(column headers)* 2024, 2025, 2026 Target, Jan… | — | OI |
| 4 | Contract | `Slide2_Contract` → 15 ⚠️ | OI |
| 5 | Weighted Proposals | `Slide2_WP` → 16 ⚠️ | OI |
| 6 | Weighted Opportunities | `Slide2_WO` → 17 ⚠️ | OI |
| 12 | *(section header)* Gross Revenue Progress (kTL) | — | NS TL |
| 14 | *(category labels)* 2024, 2025, 2026T, BL, Jan… | `Categories_NS_kTL` = 14 ✓ | NS TL |
| 15 | Contract | `Slide2_Contract` = 15 ✓ | NS TL |
| 16 | Weighted Proposals | `Slide2_WP` = 16 ✓ | NS TL |
| 17 | Weighted Opportunities | `Slide2_WO` = 17 ✓ | NS TL |
| 34 | *(section header)* EBIT (kTL) | — | EBIT TL |
| 37 | Contract | `Slide4_Contract` = 37 ✓ | EBIT TL |
| 38 | Contract+WP | `Slide4_Contract_WP` = 38 ✓ | EBIT TL |
| 39 | Contract+WP+WO | `Slide4_Contract_WP_WO` = 39 ✓ | EBIT TL |
| 45 | *(section header)* Gross Revenue Progress (k€) | — | NS EUR |
| 47 | *(category labels)* | `Categories_NS_kEUR` = 47 ✓ | NS EUR |
| 48 | Contract | `Slide3_Contract` = 48 ✓ | NS EUR |
| 53 | *(section header)* EBIT (k€) | — | EBIT EUR |
| 56 | Contract | `Slide5_Contract` = 56 ✓ | EBIT EUR |

### Row Map — BU section (cols 52–65, 14 values each)

**BU column headers** are at row 14, cols 52–65:  
`2025 | 2026 Target | Jan | Feb | Mar | Apr | May | Jun | Jul | Aug | Sep | Oct | Nov | Dec`

> ℹ️ **Config values are Excel 1-based row numbers** (not pandas 0-based).  
> The PPTX generator correctly converts with `df.iloc[xrow - 1, ...]`.  
> The dashboard was missing this `-1` offset (fixed in `mbr_dashboard.py`).

#### BU NS (NS by Business Unit) — verified layout

| BU | Tier | Pandas row | Config value (Excel row) | Config key |
|----|------|-----------|--------------------------|------------|
| MC | Order | 15 | 16 ✓ | `BU_NS_MC_Order` |
| MC | Offer | 16 | 17 ✓ | `BU_NS_MC_Offer` |
| MC | Opp | 17 | 18 ✓ | `BU_NS_MC_Opp` |
| ENG | Order | 18 | 19 ✓ | `BU_NS_ENG_Order` |
| ENG | Offer | 19 | 20 ✓ | `BU_NS_ENG_Offer` |
| ENG | Opp | 20 | 21 ✓ | `BU_NS_ENG_Opp` |
| T&SI | Order | 21 | 22 ✓ | `BU_NS_TSI_Order` |
| T&SI | Offer | 22 | 23 ✓ | `BU_NS_TSI_Offer` |
| T&SI | Opp | 23 | 24 ✓ | `BU_NS_TSI_Opp` |
| NUC | Order | 24 | 25 ✓ | `BU_NS_NUC_Order` |
| NUC | Offer | 25 | 26 ✓ | `BU_NS_NUC_Offer` |
| NUC | Opp | 26 | 27 ✓ | `BU_NS_NUC_Opp` |

#### BU EBIT — verified layout

| BU | Pandas row | Config value (Excel row) | Config key |
|----|-----------|--------------------------|------------|
| MC | 98 | 99 ✓ | `BU_EBIT_MC` |
| ENG | 99 | 100 ✓ | `BU_EBIT_ENG` |
| T&SI | 100 | 101 ✓ | `BU_EBIT_TSI` |
| NUC | 101 | 102 ✓ | `BU_EBIT_NUC` |

#### BU EBIT % — verified layout

| BU | Pandas row | Config value (Excel row) | Config key |
|----|-----------|--------------------------|------------|
| MC | 106 | 107 ✓ | `BU_EBIT_P_MC` |
| ENG | 107 | 108 ✓ | `BU_EBIT_P_ENG` |
| T&SI | 108 | 109 ✓ | `BU_EBIT_P_TSI` |
| NUC | 109 | 110 ✓ | `BU_EBIT_P_NUC` |

---

## Sheet 2: Order Intake

**Shape:** 552 rows × 25 columns

### Header
- **Row 0:** empty  
- **Row 1:** `Ref# | Customer | Project | Business Unit | Status | Jan | Feb | … | Dec`  
- **Row 2+:** project records

### Monthly columns
| Col index | 6 | 7 | 8 | … | 17 |
|-----------|---|---|---|---|----|
| Label | Jan | Feb | Mar | … | Dec |

### Summary rows (bottom of sheet)
Rows ~539–551 hold BU-level aggregate totals used by dashboard and PPTX:

| Config key | Value | Content |
|------------|-------|---------|
| `OI_Month_Header_Row` | 539 | Month label row |
| `Categories_OI_kTL` | 543 | Category label row (auto-probed ±5 rows at runtime) |
| `Slide6_ENG` | 544 | ENG BU total row |
| `Slide6_MC` | 545 | MC BU total row |
| `Slide6_TSI` | 546 | T&SI BU total row |
| `Slide6_NUC` | 547 | NUC BU total row |

> Note: these row indices are within the **Order Intake** sheet, not Summary TL.  
> The dashboard applies a runtime offset (`_oi_offset`) if the rows have shifted.

---

## Sheet 3: Ext. Prod.

**Shape:** 546 rows × 90 columns

### Header rows
- **Row 0:** Exchange rate (col 6), period labels (`PREVIOUS YEARS`, `JANUARY`, etc.)  
- **Row 1:** Period end dates  
- **Row 2:** Column headers  
- **Row 3+:** Data rows (one row per project × cost type)

### Identifier columns (0-based)
| Col | Header | Config key |
|-----|--------|------------|
| 1 | Project Code | `EXT_PROD_CODE` |
| 2 | Project Name | `EXT_PROD_NAME` |
| 3 | Type | `EXT_PROD_TYPE` (filter: GROSS FEES only) |
| 4 | Client/Subc/Associate Name | `EXT_PROD_CLIENT` |
| 6 | BU | `EXT_PROD_BU` |
| 7 | Original Currency | — (detected via "ORIGINAL CURRENCY" header) |
| 8 | Total Budget (Original Currency) | — |
| 9 | Probability | — |

### Monthly production columns (4-col repeating pattern)
Each calendar month occupies 4 columns:

| Offset | Header | Notes |
|--------|--------|-------|
| +0 | PRODUCTION OC | In original currency |
| +1 | AVG TRL/OC | Exchange rate |
| +2 | PRODUCTION | **In TL** — this is what the dashboard uses |
| +3 | PRODUCTION € | In EUR |

| Period | Start col (0-based) |
|--------|---------------------|
| Previous Years | 11 |
| January | 14 |
| February | 18 |
| March | 22 |
| April | 26 |
| May | 30 |
| June | 34 |
| July | 38 |
| August | 42 |
| September | 46 |
| October | 50 |
| November | 54 |
| December | 58 |

**PRODUCTION (TL)** for each month = `col_start + 2` (e.g. Jan TL = col 16, Feb TL = col 20).

### Summary columns (end of row)
| Col | Header | Used as |
|-----|--------|---------|
| 62 | 2026 PRODUCTION OC | — |
| 63 | 2026 PRODUCTION | 2026 total in TL |
| 65 | END OF 2026 TOTAL (OC) | Year-end projection OC |
| 66 | END OF 2026 TOTAL | Year-end projection TL — "End of Year" metric |
| 68 | Next Year | Next-year projection TL |
| 72 | YTD | Year-to-date TL |
| 74 | TD | To-date (all-time) TL |

> The dashboard auto-detects `next_year_col` and `td_col` by scanning for "NEXT YEAR" and "TD" headers, so these column positions are resilient to minor file changes.

---

## Sheet 4: WIP & BL

**Shape:** 578 rows × 102 columns

### Header rows
- **Rows 0–1:** empty / period markers  
- **Row 2:** Column headers  
- **Row 3+:** Data rows (one row per project × cost type)

### Identifier columns (0-based)
| Col | Header | Config key |
|-----|--------|------------|
| 1 | Project Code | `WIP_Col_Code` |
| 2 | Project Name | `WIP_Col_Name` |
| 3 | Type | `WIP_Col_Type` (filter: GROSS FEES only) |
| 4 | Client | `WIP_Col_Client` |
| 10 | BU | `WIP_Col_BU` |
| 11 | Original Currency | `WIP_Col_OrigCurrency` |

### Monthly columns (6-col repeating pattern)
Starting from col 15 (2024-12-31), each period has 6 columns:  
`INVOICING | PRODUCTION | AVG TRL/OC | WIP | BACKLOG | CLOS TRL/OC`

Months covered: Dec-2024, Jan-2025 … Dec-2025, then 2026 months.

### Summary columns (end of row) — all config values verified ✓
| Col | Header | Config key |
|-----|--------|------------|
| 91 | 2025 INVOICE OC | — |
| 92 | 2025 PRODUCTION OC | — |
| 93 | TOTAL INVOICE OC | `WIP_Col_TotalInvoiceOC` ✓ |
| 94 | TOTAL PRODUCTION OC | `WIP_Col_TotalProdOC` ✓ |
| 96 | WIP OC | — |
| 97 | BL OC | — |
| 98 | WIP TL | `WIP_Col_WIP_TL` ✓ |
| 99 | BL TL | — |

---

## Sheet 5: Gross Margin

**Shape:** 857 rows × 32 columns (PPTX generator only, not dashboard)

### Header rows
- **Row 0:** Exchange rates (45.5, 55.2, 1000…)  
- **Row 1:** Filter label ("MAIN")  
- **Row 2:** SUBTOTAL formulas  
- **Row 3:** Column headers  
- **Row 4+:** Data

### Key columns (0-based)
| Col | Header |
|-----|--------|
| 1 | Ref # (project code) |
| 2 | Customer |
| 3 | Business Unit |
| 4 | Short Project Name |
| 5 | Issue Date |
| 6 | Status |
| 7 | Item (GR / SUBCON / ASSOCIATE / PE / PEXP / CM) |
| 8 | Value (kTL) |
| 13–24 | Monthly 2026 data (Jan–Dec) |
| 25 | 2026 TOTAL |
| 28 | TD (Total to Date) |
| 29 | End of 2026 |

---

## Data Type Notes

- **NS/EBIT values** in Summary TL are **cumulative YTD** (each month includes all prior months).  
  The dashboard calls `monthly_val()` to extract per-month figures by subtracting the previous month.
- **BU NS/EBIT** in Summary TL are also **cumulative YTD** (same logic).
- **Ext. Prod. PRODUCTION (TL)** columns are **already in TL** — safe to sum across rows with same Project Code + Currency.
- **WIP_TL** (col 98) is in TL.
- All OI monthly values are **non-cumulative** (individual month amounts).

---

## Config Verification Summary

| Config section | Status |
|----------------|--------|
| NS kTL row indices (Slide2_*) | ✓ Correct |
| NS kEUR row indices (Slide3_*) | ✓ Correct |
| EBIT kTL row indices (Slide4_*) | ✓ Correct |
| EBIT kEUR row indices (Slide5_*) | ✓ Correct |
| NS column ranges (cols 1–16) | ✓ Correct |
| EBIT column ranges (cols 1–16) | ✓ Correct |
| BU NS/EBIT column range (52–65) | ✓ Correct |
| WIP column indices (93, 94, 98) | ✓ Correct |
| Ext. Prod. Next Year (col 68), TD (col 74) | ✓ Correct |
| BU_NS_* row indices (all 12 keys) | ✓ Correct (Excel 1-based; dashboard now applies −1) |
| BU_EBIT_* row indices (4 keys) | ✓ Correct (Excel 1-based; dashboard now applies −1) |
| BU_EBIT_P_* row indices (4 keys) | ✓ Correct (Excel 1-based; PPTX generator applies −1) |
