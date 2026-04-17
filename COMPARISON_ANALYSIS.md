# MRC Finance: Dashboard vs PPTX Generator Comparison

**Date:** 2026-04-17  
**Files Compared:**
- `mbr_dashboard.py` (Streamlit interactive dashboard)
- `generate_full_mbr_stacked.py` (PPTX presentation generator)

---

## EXECUTIVE SUMMARY

The dashboard and PPTX generator use **similar data loading patterns** but have **critical differences** in:
1. **Row offset handling** – dashboard misses dynamic offset detection
2. **BU row indices** – dashboard uses wrong rows for MC/T&SI/NUC
3. **Monthly data selection** – dashboard sums all months; PPTX dynamically selects month column
4. **Config usage** – dashboard loads config but doesn't use offset/column logic

**Impact:** Dashboard may show **incorrect BU data** for MC/T&SI/NUC BUs, and **inflated project values** by summing across all months.

---

## DETAILED COMPARISON

### 1. ORDER INTAKE: BU SUMMARY DATA

#### Loading Strategy
| Aspect | Dashboard | PPTX Generator |
|--------|-----------|----------------|
| BU Summary Rows | Fixed: `R["Slide6_ENG"]=544`, `R["Slide6_MC"]=545`, etc. | Same config keys, BUT with offset detection |
| Row Offset Detection | **NONE** – uses raw config value | **YES** – `_oi_offset = _oi_cats_row - _oi_cats_cfg` (line 1955) |
| Offset Application | N/A | Applied to all row reads: `R.get(key) + _oi_offset` |

**Risk:** If your Excel file has extra/fewer rows than the standard 586K lines, dashboard will misalign.

#### BU Data Arrays

**Dashboard (lines 267-270):**
```python
"ENG":  [safe_float(df_oi.iloc[R["Slide6_ENG"], c]) for c in range(4, 18)],
"MC":   [safe_float(df_oi.iloc[R["Slide6_MC"],  c]) for c in range(4, 18)],
"T&SI": [safe_float(df_oi.iloc[R["Slide6_TSI"], c]) for c in range(4, 18)],
"NUC":  [safe_float(df_oi.iloc[R["Slide6_NUC"], c]) for c in range(4, 18)],
```
- Defaults: 544, 545, 546, 547 (pandas indices = Excel rows 545, 546, 547, 548)

**PPTX Generator (lines 1963-1966):**
```python
s6_eng = _oi_row("Slide6_ENG", 544)
s6_mc  = _oi_row("Slide6_MC",  545)
s6_tsi = _oi_row("Slide6_TSI", 546)
s6_nuc = _oi_row("Slide6_NUC", 547)
```
- Same defaults BUT wrapped in `_oi_row(key, default)` which applies `_oi_offset`
- Definition (line 1959-1961):
  ```python
  def _oi_row(key, default):
      return [safe_float(df_oi.iloc[R.get(key, default) + _oi_offset, c])
              for c in range(_oi_col_start, _oi_col_end)]
  ```

**Column Range:** Both use `range(4, 18)` = 14 values [2025_actual, Target, Jan, Feb, …, Dec]
- **CORRECT** – both match the structure of BU summary rows

**Monthly Indexing:**
- Dashboard: Uses `cats_oi.index(month_abbr)` label lookup → finds index 2 for "Jan"
- PPTX: Uses `MON_OI = 2 + month_idx` arithmetic lookup → also gives index 2 for January
- **Both equivalent** ✓

---

### 2. ORDER INTAKE: INDIVIDUAL PROJECTS

#### Dashboard (lines 172-187)

```python
oi_projects = []
for ri in range(2, R["Slide6_ENG"]):  # Loop rows 2 to 544 (all projects)
    row = df_oi.iloc[ri]
    name = str(row.iloc[3]).strip()  # Col D = project name
    if not name or name.lower() == "nan":
        continue
    # **SUMS ALL 14 COLUMNS** (includes 2025, Target, all months)
    total_val = sum(safe_float(row.iloc[c]) for c in range(4, 18))
    if total_val > 0:
        bu_raw = str(row.iloc[4]).strip()  # Col E = BU name
        oi_projects.append({...})
```

**Column interpretation for project rows:**
- Col C (2) = client
- Col D (3) = project name
- Col E (4) = BU name (text → safe_float = 0)
- Col F (5) = ??? (historical data? annual target?)
- Cols G-R (6-17) = monthly values

**Problem:** `total_val = sum(range(4, 18))` includes col F (5th element), which likely isn't a monthly value. If col F is 2023 actual or 2024 target (like in BU summary rows), this inflates project totals.

#### PPTX Generator (lines 2353, 2369-2376)

```python
oi_col = O.get("OI_Col_Base", 6) + month_idx  # Current month only!
for idx in range(2, oi_data_end):  # Row 2 to Categories_OI_kTL (543)
    row = df_oi.iloc[idx]
    val = safe_float(row[oi_col])  # **SINGLE MONTH COLUMN ONLY**
    if val <= 0:
        continue
    ...
    oi_projects.append((project, client, val, bu))
```

**Differences:**
| Aspect | Dashboard | PPTX |
|--------|-----------|------|
| Project Value | Sum of cols 5-17 (14 columns) | Current month only (col 6+month_idx) |
| Interpretation | Yearly/cumulative total | Monthly signed OI |
| Data Range | All months | Current month |
| Count | ALL projects ever | Only projects with >0 in current month |

**Impact:** Dashboard shows persistent list of projects; PPTX shows month-specific projects.

---

### 3. BUSINESS UNIT (BU) NODE SUMMARIES

#### Dashboard BU_NS Structure (lines 227-238)

```python
bu_ns = {}
for b_key, b_ord, b_off, b_opp in [
    ("ENG",  "BU_NS_ENG_Order", "BU_NS_ENG_Offer", "BU_NS_ENG_Opp"),
    ("MC",   "BU_NS_MC_Order",  "BU_NS_MC_Offer",  "BU_NS_MC_Opp"),
    ("T&SI", "BU_NS_TSI_Order", "BU_NS_TSI_Offer", "BU_NS_TSI_Opp"),
    ("NUC",  "BU_NS_NUC_Order", "BU_NS_NUC_Offer", "BU_NS_NUC_Opp"),
]:
    bu_ns[b_key] = {
        "cats": cats_bu_ns,
        "Order": row_vals(df, _r_cfg.get(b_ord, 19), c_ns_s, c_ns_e),
        "Offer": row_vals(df, _r_cfg.get(b_off, 20), c_ns_s, c_ns_e),
        "Opp":   row_vals(df, _r_cfg.get(b_opp, 21), c_ns_s, c_ns_e),
    }
```

**Row Defaults:**
- ENG: Order=19, Offer=20, Opp=21
- MC: Order=19 **(?wrong!)**,  Offer=20, Opp=21
- T&SI: Order=19 **(?wrong!)**, Offer=20, Opp=21
- NUC: Order=19 **(?wrong!)**, Offer=20, Opp=21

#### PPTX Generator BU_NS Structure (lines 2062-2084)

```python
def bu_ns_row(xrow):
    return [safe_float(df.iloc[xrow - 1, c]) 
            for c in range(C.get("BU_NS_Start", 51), C.get("BU_NS_End", 66))]

bu_ns = {
    "ENG": {
        "Order": bu_ns_row(R.get("BU_NS_ENG_Order", 19)),    # Row 19
        "Offer": bu_ns_row(R.get("BU_NS_ENG_Offer", 20)),    # Row 20
        "Opp": bu_ns_row(R.get("BU_NS_ENG_Opp", 21)),        # Row 21
    },
    "MC": {
        "Order": bu_ns_row(R.get("BU_NS_MC_Order", 16)),     # Row 16 ← DIFFERENT!
        "Offer": bu_ns_row(R.get("BU_NS_MC_Offer", 17)),     # Row 17
        "Opp": bu_ns_row(R.get("BU_NS_MC_Opp", 18)),         # Row 18
    },
    "T&SI": {
        "Order": bu_ns_row(R.get("BU_NS_TSI_Order", 22)),    # Row 22 ← DIFFERENT!
        "Offer": bu_ns_row(R.get("BU_NS_TSI_Offer", 23)),    # Row 23
        "Opp": bu_ns_row(R.get("BU_NS_TSI_Opp", 24)),        # Row 24
    },
    "NUC": {
        "Order": bu_ns_row(R.get("BU_NS_NUC_Order", 25)),    # Row 25 ← DIFFERENT!
        "Offer": bu_ns_row(R.get("BU_NS_NUC_Offer", 26)),    # Row 26
        "Opp": bu_ns_row(R.get("BU_NS_NUC_Opp", 27)),        # Row 27
    },
}
```

**Critical Mismatch:**
- Dashboard hardcodes MC/T&SI/NUC to use rows 19,20,21 (ENG data!)
- PPTX uses correct rows: MC=16-18, T&SI=22-24, NUC=25-27

**Impact:** Dashboard BU rankings for MC/T&SI/NUC are **showing ENG data** instead of their own data.

---

### 4. CONFIGURATION & OFFSET SYSTEM

#### Dashboard Config Loading (lines 35-50, 53-62)

```python
R = _r_cfg.get("Excel_Mapping.Rows", {}) if _r_cfg else {}
# Inline defaults for every key:
R.get("Categories_OI_kTL", 543)
R.get("Slide6_ENG", 544)
R.get("Slide6_MC", 545)
# etc.

C = _c_cfg.get("Excel_Mapping.Columns", {}) if _c_cfg else {}
# Inline defaults:
C.get("OI_kTL_Start", 4)
C.get("OI_kTL_End", 18)
# etc.
```

**Missing:** No offset dictionary (O)
```python
# Dashboard does NOT load:
O = config.get("Excel_Mapping.Offsets", {})
```

#### PPTX Generator Config Loading (lines 1990-1992)

```python
R = config.get("Excel_Mapping.Rows", {})
C = config.get("Excel_Mapping.Columns", {})
O = config.get("Excel_Mapping.Offsets", {})  # ← DASHBOARD MISSING THIS
```

**Offset Usages in PPTX (unavailable to dashboard):**
```python
MON_NS = O.get("MON_NS_Base", 3) + month_idx       # Offset = 3
MON_EBIT = O.get("MON_EBIT_Base", 4) + month_idx   # Offset = 4
MON_BU = O.get("MON_BU_Base", 2) + month_idx       # Offset = 2
MON_OI = O.get("MON_OI_Base", 2) + month_idx       # Offset = 2
OI_Col_Base = O.get("OI_Col_Base", 6) + month_idx  # Base = 6
```

---

### 5. HARDCODED vs CONFIG-DRIVEN

#### Dashboard Hardcoded Values (Lines 20-28, 64-67)

```python
BU_COLORS = {
    "ENG": "#0EA5E9",    # Blue
    "MC": "#10B981",     # Green
    "T&SI": "#F59E0B",   # Amber
    "NUC": "#EF4444",    # Red
}

BU_ABBREV = {
    "Engineering": "ENG",
    "MC": "MC",
    "T&SI": "T&SI",
    "TSI": "T&SI",
    "Nuclear": "NUC",
}

MONTH_ORDER = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]
```

**Status:** Hardcoded in Python code
- **Maintenance:** Requires code edit + redeployment
- **Config:** Would need additions to config.yaml

#### PPTX Generator Hardcoded Values (lines 50-57)

```python
# Loaded from config:
COLORS = {
    "NS_Color": colors.get("NS_Color", "#0EA5E9"),
    "EBIT_Color": colors.get("EBIT_Color", "#10B981"),
    # etc.
}

BU_COLORS = {
    "ENG": bu_colors.get("ENG", "#0EA5E9"),
    # etc.
}
```

**Status:** Config-driven with hardcoded fallback
- **Maintenance:** Edit config.yaml only (no code change needed)
- **Flexibility:** Can change colors per PPTX generation

---

### 6. CACHING STRATEGY

| Aspect | Dashboard | PPTX |
|--------|-----------|------|
| **Cache Type** | Streamlit `@st.cache_data(ttl=3600)` | No caching (inline execution) |
| **Cache Duration** | 1 hour | N/A – static/fresh each run |
| **Cache Scope** | Per-month data dictionary | N/A |
| **Risk** | Stale data if Excel updated during hour | None – always fresh |
| **Benefit** | Fast page loads/interactivity | Data is always current |

**Implication:** Dashboard users see up to 1-hour-old data; PPTX always reads fresh.

---

## CRITICAL ISSUES TO FIX

### Issue 1: BU Row Indices Wrong for MC/T&SI/NUC ⚠️ HIGH PRIORITY

**Location:** `mbr_dashboard.py` lines 227-237

**Current (WRONG):**
```python
for b_key, b_ord, b_off, b_opp in [
    ("ENG",  "BU_NS_ENG_Order", "BU_NS_ENG_Offer", "BU_NS_ENG_Opp"),
    ("MC",   "BU_NS_MC_Order",  "BU_NS_MC_Offer",  "BU_NS_MC_Opp"),     # Uses config key BU_NS_MC_Order → default 19
    ("T&SI", "BU_NS_TSI_Order", "BU_NS_TSI_Offer", "BU_NS_TSI_Opp"),    # Uses config key BU_NS_TSI_Order → default 19
    ("NUC",  "BU_NS_NUC_Order", "BU_NS_NUC_Offer", "BU_NS_NUC_Opp"),    # Uses config key BU_NS_NUC_Order → default 19
]:
```

**Expected (from PPTX):**
```python
# MC should use config keys that default to 16, 17, 18 (not 19, 20, 21)
# T&SI should use config keys that default to 22, 23, 24
# NUC should use config keys that default to 25, 26, 27
```

**Fix:** Verify config.yaml has correct row keys and that dashboard code reads them. Check if dashboard defaults are overridden by actual config values.

---

### Issue 2: OI Project Total Value Includes Non-Monthly Columns ⚠️ MEDIUM PRIORITY

**Location:** `mbr_dashboard.py` line 178

**Current:**
```python
total_val = sum(safe_float(row.iloc[c]) for c in range(4, 18))  # 14 columns
```

**Why Wrong:** Sums all columns including col F (index 5), which for project rows may not be a monthly value.

**Comparison to PPTX (line 2353):**
```python
oi_col = O.get("OI_Col_Base", 6) + month_idx  # Only current month
val = safe_float(row[oi_col])
```

**Options:**
1. Sum only monthly columns: `sum(range(6, 18))` (12 columns, Jan-Dec only)
2. Use current month column: `row.iloc[6 + month_idx]` (match PPTX logic)
3. Clarify intent: Is this yearly total or monthly value?

---

### Issue 3: No Row Offset Detection ⚠️ MEDIUM PRIORITY

**Location:** `mbr_dashboard.py` lines 172-173 (missing offset logic from PPTX lines 1947-1957)

**Dashboard (current):**
```python
for ri in range(2, R["Slide6_ENG"]):
```

**PPTX (robust):**
```python
_oi_offset = _oi_cats_row - _oi_cats_cfg  # Detect file variation
_oi_row = R.get(key, default) + _oi_offset  # Apply offset
```

**Risk:** If your Excel sheets vary in row count, dashboard will misalign data.

**Fix:** Implement PPTX's offset detection in dashboard `load_month()` function.

---

### Issue 4: Missing Offset Dictionary (O) ⚠️ LOW PRIORITY (for current use)

**Location:** `mbr_dashboard.py` (not loaded at all)

**PPTX uses (lines 2098-2103):**
```python
MON_OI = O.get("MON_OI_Base", 2) + month_idx
```

**Dashboard equivalent:**
```python
# Uses label lookup instead: cats_oi.index(month_abbr)
# Works, but is less flexible than offset approach
```

**Impact:** Current dashboard design works, but cannot adopt PPTX's dynamic offset pattern without refactoring.

---

## RECOMMENDATIONS

### Immediate (This Week)

1. **Fix BU Row Indices** – Check config.yaml to verify MC/T&SI/NUC row keys are correct and defaults match PPTX (16-18, 22-24, 25-27)
2. **Clarify OI Project Value** – Decide if `total_val` should use single month or sum all months; update logic
3. **Test BU Data** – Compare dashboard BU_NS values for MC/T&SI/NUC against PPTX output to verify they match

### Short Term (This Sprint)

4. **Implement Row Offset Detection** – Apply PPTX's `_oi_offset` logic to dashboard (copy lines 1947-1957)
5. **Load Offset Dictionary** – Add `O = config.get("Excel_Mapping.Offsets", {})` if future dynamic month selection is needed
6. **Config-Driven Colors** – Move `BU_COLORS` to config.yaml (less urgent if rarely changed)

### Testing Checklist

- [ ] Dashboard OI BU totals match PPTX slide 6 for Jan/Feb/etc.
- [ ] Dashboard BU_NS Order values match PPTX for all 4 BUs
- [ ] OI project list aligns between dashboard and PPTX (or document intentional differences)
- [ ] Cache invalidation works if Excel file changes mid-hour
- [ ] Config.yaml has all required keys; defaults are safe fallbacks

---

## APPENDIX: Side-by-Side Line References

| Task | Dashboard | PPTX Generator |
|------|-----------|----------------|
| Config Loading | Lines 35-50, 53-62 | Lines 1959-1963, 1990-1992 |
| OI Sheet Read | Line 163 | Line 1876 |
| Row Offset Detection | **MISSING** | Lines 1947-1957 |
| OI Categories | Line 168 | Lines 1915, 1939-1954 |
| BU Summary Reads | Lines 267-270 | Lines 1963-1966 |
| BU Node Extraction | Lines 227-238 | Lines 2062-2084 |
| Individual Projects | Lines 172-187 | Lines 2350-2376 |
| Monthly Offset | Label lookup (637) | Arithmetic: `MON_OI = 2 + month_idx` (2098) |
| Caching | Lines 153, 280 | None |
| Colors | Line 20 (hardcoded) | Lines 50-57 (config) |
| BU Filters | Lines 498-507 | Lines 2360, 2394 |

---

**End of Comparison Analysis**
ccnnnn.