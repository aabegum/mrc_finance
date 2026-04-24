import colorsys
import io
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from config_loader import ConfigLoader

st.set_page_config(page_title="MRC Business Review", layout="wide", page_icon="📊")

SCRIPT_DIR = Path(__file__).parent


# ── Config ─────────────────────────────────────────────────────────────────────
_cfg    = ConfigLoader(str(SCRIPT_DIR / "config"))
_r_cfg  = _cfg.get("Excel_Mapping.Rows")    or {}
_c_cfg  = _cfg.get("Excel_Mapping.Columns") or {}
_o_cfg  = _cfg.get("Excel_Mapping.Offsets") or {}

# ── Colors (sourced from config.yaml) ─────────────────────────────────────────
PRIMARY  = _cfg.get("THEMES.premium.Colors.Dark.Hex") or "#1E3A8A"
_bu_cfg  = _cfg.get("BU_Colors") or {}
BU_COLORS = {
    "ENG":  _bu_cfg.get("ENG",  "#0EA5E9"),
    "MC":   _bu_cfg.get("MC",   "#10B981"),
    "T&SI": _bu_cfg.get("TSI",  "#FAB611"),
    "NUC":  _bu_cfg.get("NUC",  "#93C572"),
}

def bu_tier_colors(bu_name: str) -> list[str]:
    """Return [Order, Offer, Opp] hex colors for a BU using BU_Shade_Tiers config."""
    base_hex  = BU_COLORS.get(bu_name, "#0EA5E9").lstrip("#")
    r, g, b   = (int(base_hex[i:i+2], 16) / 255.0 for i in (0, 2, 4))
    h, s, v   = colorsys.rgb_to_hsv(r, g, b)

    shade_cfg = _cfg.get("BU_Shade_Tiers") or {}
    v         = max(v, shade_cfg.get("Base_Brightness_Floor", 0.62))

    def _hsv_hex(s_mult, v_mult):
        ns, nv = min(1.0, s * s_mult), min(1.0, v * v_mult)
        rr, gg, bb = colorsys.hsv_to_rgb(h, ns, nv)
        return "#{:02X}{:02X}{:02X}".format(int(rr * 255), int(gg * 255), int(bb * 255))

    ord_cfg = shade_cfg.get("Order", {})
    off_cfg = shade_cfg.get("Offer", {})
    opp_hex = (shade_cfg.get("Opp") or {}).get("Color", "#BBBBBB")

    return [
        _hsv_hex(ord_cfg.get("Saturation_Multiplier", 1.0), ord_cfg.get("Brightness_Multiplier", 1.0)),
        _hsv_hex(off_cfg.get("Saturation_Multiplier", 0.45), off_cfg.get("Brightness_Multiplier", 1.18)),
        opp_hex,
    ]


# Raw BU names from Excel → abbreviation used everywhere internally
BU_ABBREV = {
    "Engineering": "ENG",
    "MC":          "MC",
    "T&SI":        "T&SI",
    "TSI":         "T&SI",
    "Nuclear":     "NUC",
}

# ── Excel row/column/offset indices from config ────────────────────────────────
R = {
    "Categories_NS_kTL":       _r_cfg.get("Categories_NS_kTL", 14),
    "Categories_EBIT_kTL":     _r_cfg.get("Categories_EBIT_kTL", 36),
    "Categories_OI_kTL":       _r_cfg.get("Categories_OI_kTL", 543),
    "Slide2_Contract":         _r_cfg.get("Slide2_Contract", 15),
    "Slide2_WP":               _r_cfg.get("Slide2_WP", 16),
    "Slide2_WO":               _r_cfg.get("Slide2_WO", 17),
    "Slide4_Contract":         _r_cfg.get("Slide4_Contract", 37),
    "Slide4_Contract_WP":      _r_cfg.get("Slide4_Contract_WP", 38),
    "Slide4_Contract_WP_WO":   _r_cfg.get("Slide4_Contract_WP_WO", 39),
    "Slide6_ENG":              _r_cfg.get("Slide6_ENG", 544),
    "Slide6_MC":               _r_cfg.get("Slide6_MC", 545),
    "Slide6_TSI":              _r_cfg.get("Slide6_TSI", 546),
    "Slide6_NUC":              _r_cfg.get("Slide6_NUC", 547),
    "WIP_DataStartRow":        _r_cfg.get("WIP_DataStartRow", 3),
}

O = {
    "MON_NS_Base":   _o_cfg.get("MON_NS_Base",   3),
    "MON_EBIT_Base": _o_cfg.get("MON_EBIT_Base", 3),
    "MON_BU_Base":   _o_cfg.get("MON_BU_Base",   2),
    "MON_OI_Base":   _o_cfg.get("MON_OI_Base",   2),
    "OI_Col_Base":   _o_cfg.get("OI_Col_Base",   6),
}

C = {
    "WIP_Col_Name":           _c_cfg.get("WIP_Col_Name", 2),
    "WIP_Col_Type":           _c_cfg.get("WIP_Col_Type", 3),
    "WIP_Col_Client":         _c_cfg.get("WIP_Col_Client", 4),
    "WIP_Col_BU":             _c_cfg.get("WIP_Col_BU", 10),
    "WIP_Col_OrigCurrency":   _c_cfg.get("WIP_Col_OrigCurrency", 11),
    "WIP_Col_TotalInvoiceOC": _c_cfg.get("WIP_Col_TotalInvoiceOC", 93),
    "WIP_Col_TotalProdOC":    _c_cfg.get("WIP_Col_TotalProdOC", 94),
    "WIP_Col_WIP_TL":         _c_cfg.get("WIP_Col_WIP_TL", 98),
}

# WIP minimum threshold — from config, falls back to 1 M TL
WIP_MIN_TL = int((_cfg.get("WIP_Table.WIP_Min_TL") or 1_000_000))

MONTH_ORDER = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


# ── Formatters ─────────────────────────────────────────────────────────────────
def safe_float(val):
    try:
        return 0.0 if pd.isna(val) else float(val)
    except Exception:
        return 0.0


def clean_label(val):
    """Return clean string for category labels (avoids '2025.0' style)."""
    if pd.isna(val):
        return ""
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
    except Exception:
        pass
    return str(val).strip()


def human_k(v):
    """Format kTL values (already in thousands)."""
    sign = "-" if v < 0 else ""
    av = abs(v)
    if av >= 1_000_000:
        return f"{sign}{av / 1_000_000:.1f}M"
    if av >= 1_000:
        return f"{sign}{av / 1_000:.0f}k"
    return f"{sign}{av:,.0f}"


def human_tl(v):
    """Format raw TL values."""
    sign = "-" if v < 0 else ""
    av = abs(v)
    if av >= 1_000_000_000:
        return f"{sign}{av / 1_000_000_000:.1f}B"
    if av >= 1_000_000:
        return f"{sign}{av / 1_000_000:.1f}M"
    if av >= 1_000:
        return f"{sign}{av / 1_000:.1f}k"
    return f"{sign}{av:,.0f}"


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def _parse_margin_year(date_text: str):
    if not date_text or str(date_text).strip().upper() in ["NAN", "NONE"]:
        return None
    txt = str(date_text).strip()
    year_match = re.search(r"\b(20\d{2}|19\d{2})\b", txt)
    if year_match:
        return int(year_match.group(1))
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(txt, fmt).year
        except Exception:
            pass
    return None


def _margin_month_label(month_name: str, date_text: str):
    year = _parse_margin_year(date_text)
    return f"{month_name} {year}" if year else month_name


def row_vals(df, row, col_start, col_end):
    return [safe_float(df.iloc[row, c]) for c in range(col_start, col_end)]


_MONTH_ABBREVS = {"jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"}


def monthly_val(arr, cats, idx):
    """Return singular per-month value from a cumulative YTD array.

    Finds the previous monthly column in cats by walking backwards and
    subtracts it. January (no prior month column) returns the raw value.
    Returns 0.0 if idx is -1.
    """
    if idx == -1:
        return 0.0
    prev = next(
        (i for i in range(idx - 1, -1, -1)
         if str(cats[i]).strip().lower()[:3] in _MONTH_ABBREVS),
        -1,
    )
    return arr[idx] - (arr[prev] if prev != -1 else 0.0)


def to_monthly_deltas(arr, cats):
    """Convert a cumulative YTD array to per-month values.

    Identifies monthly columns by matching category labels against month
    abbreviations. The first monthly column (January) is kept as-is.
    Each subsequent month becomes arr[i] - arr[i-1].
    Non-month columns (2025, 2026T, BL) are left unchanged.
    """
    monthly_idx = [
        i for i, c in enumerate(cats[:len(arr)])
        if str(c).strip().lower()[:3] in _MONTH_ABBREVS
    ]
    out = list(arr)
    for pos, idx in enumerate(monthly_idx):
        if pos > 0:
            out[idx] = arr[idx] - arr[monthly_idx[pos - 1]]
    return out


def _cat_index(cats, label):
    """Case-insensitive index lookup into a category list. Returns -1 if not found."""
    label_l = str(label).lower()
    for i, c in enumerate(cats):
        if str(c).lower() == label_l:
            return i
    return -1


# ── File scanning ──────────────────────────────────────────────────────────────
def _parse_ver(v):
    """Parse version string like 'v1.10' into a comparable tuple."""
    return tuple(int(n) for n in re.findall(r"\d+", v))


def scan_excel_files():
    pattern = re.compile(
        r"(\d{6})_Project_Budget_Analysis_\w+_(v[\d.]+)\.xlsx", re.IGNORECASE
    )
    found = {}
    for f in SCRIPT_DIR.rglob("*.xlsx"):
        if f.name.startswith("~"):
            continue
        m = pattern.match(f.name)
        if not m:
            continue
        date_part, ver = m.group(1), m.group(2)
        try:
            year  = int("20" + date_part[:2])
            month = int(date_part[2:4])
            label = datetime(year, month, 1).strftime("%B %Y")
            if label not in found or _parse_ver(ver) > _parse_ver(found[label][1]):
                found[label] = (f, ver)
        except Exception:
            pass
    return {k: v[0] for k, v in sorted(found.items())}


EXCEL_FILES = scan_excel_files()


# ── Data loading ───────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def load_month(filepath: str):
    xl = pd.ExcelFile(filepath)
    required = ["Summary TL", "Order Intake", "WIP & BL"]
    missing = [s for s in required if s not in xl.sheet_names]
    if missing:
        st.error(f"Missing sheets: {missing}")
        return None

    df     = pd.read_excel(xl, sheet_name="Summary TL",  header=None)
    df_oi  = pd.read_excel(xl, sheet_name="Order Intake", header=None)
    df_wip = pd.read_excel(xl, sheet_name="WIP & BL",    header=None)

    # ── Row Offset Detection ── dynamic adjustment for file variations ──────────
    _oi_offset       = 0
    _oi_offset_found = False
    expected_cat_row = R["Categories_OI_kTL"]
    _oi_col_start    = _c_cfg.get("OI_kTL_Start", 4)
    for probe_offset in range(-10, 25):
        probe_row = expected_cat_row + probe_offset
        if 0 <= probe_row < len(df_oi):
            _c4 = str(df_oi.iloc[probe_row, _oi_col_start]).strip()
            if _oi_col_start + 2 < len(df_oi.columns):
                _c6 = str(df_oi.iloc[probe_row, _oi_col_start + 2]).strip()
                if (_c4.split(".")[0].isdigit() and len(_c4.split(".")[0]) == 4) or "2025" in _c4:
                    if "jan" in _c6.lower() or "feb" in _c6.lower():
                        _oi_offset       = probe_offset
                        _oi_offset_found = True
                        break
    if not _oi_offset_found:
        st.warning(
            f"OI row offset detection failed for **{Path(filepath).name}**. "
            f"Using default row {expected_cat_row}. Order Intake data may be misaligned."
        )

    cats_ns   = [clean_label(df.iloc[R["Categories_NS_kTL"], c])    for c in range(1, 17)]
    cats_ebit = [clean_label(df.iloc[R["Categories_EBIT_kTL"], c])  for c in range(1, 17)]
    cats_oi   = [clean_label(df_oi.iloc[R["Categories_OI_kTL"] + _oi_offset, c]) for c in range(4, 18)]

    # Extract month index from filename to select correct OI column
    month_idx = 0
    match = re.search(r"(\d{6})", str(filepath))
    if match:
        try:
            month_idx = int(match.group(1)[2:4]) - 1  # Jan=0, Feb=1, …
        except Exception:
            month_idx = 0

    oi_col      = O["OI_Col_Base"] + month_idx
    _oi_data_end = R["Slide6_ENG"] + _oi_offset
    oi_projects = []
    for ri in range(2, _oi_data_end):
        row  = df_oi.iloc[ri]
        name = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        if not name or name.lower() == "nan" or "total" in name.lower() or "subtotal" in name.lower():
            continue
        client = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if "total" in client.lower() or "subtotal" in client.lower():
            continue
        month_val = safe_float(row.iloc[oi_col])
        if month_val > 0:
            bu_raw = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
            oi_projects.append({
                "project": name,
                "client":  client,
                "bu":      BU_ABBREV.get(bu_raw, bu_raw),
                "value":   month_val,
            })
    oi_projects.sort(key=lambda x: x["value"], reverse=True)

    # WIP projects — Nuclear excluded per business rule
    wip_projects = []
    for ri in range(R["WIP_DataStartRow"], len(df_wip)):
        row       = df_wip.iloc[ri]
        proj_type = str(row.iloc[C["WIP_Col_Type"]]).strip()   if pd.notna(row.iloc[C["WIP_Col_Type"]])   else ""
        if proj_type != "GROSS FEES":
            continue
        client = str(row.iloc[C["WIP_Col_Client"]]).strip() if pd.notna(row.iloc[C["WIP_Col_Client"]]) else ""
        if not client or "akkuyu" in client.lower():
            continue
        bu_raw = str(row.iloc[C["WIP_Col_BU"]]).strip() if pd.notna(row.iloc[C["WIP_Col_BU"]]) else ""
        bu     = BU_ABBREV.get(bu_raw, bu_raw)
        if bu == "NUC":
            continue
        wip_tl = safe_float(row.iloc[C["WIP_Col_WIP_TL"]])
        if wip_tl < WIP_MIN_TL:
            continue
        wip_projects.append({
            "name":          str(row.iloc[C["WIP_Col_Name"]]).strip()          if pd.notna(row.iloc[C["WIP_Col_Name"]])          else "",
            "client":        client,
            "bu":            bu,
            "orig_currency": str(row.iloc[C["WIP_Col_OrigCurrency"]]).strip()  if pd.notna(row.iloc[C["WIP_Col_OrigCurrency"]])  else "",
            "inv_oc":        safe_float(row.iloc[C["WIP_Col_TotalInvoiceOC"]]),
            "prod_oc":       safe_float(row.iloc[C["WIP_Col_TotalProdOC"]]),
            "wip_tl":        wip_tl,
        })
    wip_projects.sort(key=lambda x: x["wip_tl"], reverse=True)

    # ── BU-level data ──────────────────────────────────────────────────────────
    c_ns_s   = _c_cfg.get("BU_NS_Start",   52)
    c_ns_e   = _c_cfg.get("BU_NS_End",     66)
    c_ebit_s = _c_cfg.get("BU_EBIT_Start", 52)
    c_ebit_e = _c_cfg.get("BU_EBIT_End",   66)

    cats_bu_ns   = [clean_label(df.iloc[_r_cfg.get("Categories_BU_NS",   14), c]) for c in range(c_ns_s,   c_ns_e)]
    cats_bu_ebit = [clean_label(df.iloc[_r_cfg.get("Categories_BU_EBIT", 97), c]) for c in range(c_ebit_s, c_ebit_e)]

    bu_ns = {}
    for b_key, b_ord, b_off, b_opp in [
        ("ENG",  "BU_NS_ENG_Order", "BU_NS_ENG_Offer", "BU_NS_ENG_Opp"),
        ("MC",   "BU_NS_MC_Order",  "BU_NS_MC_Offer",  "BU_NS_MC_Opp"),
        ("T&SI", "BU_NS_TSI_Order", "BU_NS_TSI_Offer", "BU_NS_TSI_Opp"),
        ("NUC",  "BU_NS_NUC_Order", "BU_NS_NUC_Offer", "BU_NS_NUC_Opp"),
    ]:
        bu_ns[b_key] = {
            "cats":  cats_bu_ns,
            "Order": row_vals(df, _r_cfg.get(b_ord, 19 if b_key == "ENG" else (16 if b_key == "MC" else (22 if b_key == "T&SI" else 25))) - 1, c_ns_s, c_ns_e),
            "Offer": row_vals(df, _r_cfg.get(b_off, 20 if b_key == "ENG" else (17 if b_key == "MC" else (23 if b_key == "T&SI" else 26))) - 1, c_ns_s, c_ns_e),
            "Opp":   row_vals(df, _r_cfg.get(b_opp, 21 if b_key == "ENG" else (18 if b_key == "MC" else (24 if b_key == "T&SI" else 27))) - 1, c_ns_s, c_ns_e),
        }

    bu_ebit = {}
    for b_key, b_row in [
        ("ENG",  "BU_EBIT_ENG"),
        ("MC",   "BU_EBIT_MC"),
        ("T&SI", "BU_EBIT_TSI"),
        ("NUC",  "BU_EBIT_NUC"),
    ]:
        bu_ebit[b_key] = {
            "cats":  cats_bu_ebit,
            "Total": row_vals(df, _r_cfg.get(b_row, 100) - 1, c_ebit_s, c_ebit_e),
        }

    # ── Build raw dicts ────────────────────────────────────────────────────────
    ns_out = {
        "cats":     cats_ns,
        "Contract": row_vals(df, R["Slide2_Contract"], 1, 17),
        "WP":       row_vals(df, R["Slide2_WP"],       1, 17),
        "WO":       row_vals(df, R["Slide2_WO"],       1, 17),
    }
    ebit_out = {
        "cats":           cats_ebit,
        "Contract":       row_vals(df, R["Slide4_Contract"],       1, 16),
        "Contract+WP":    row_vals(df, R["Slide4_Contract_WP"],    1, 16),
        "Contract+WP+WO": row_vals(df, R["Slide4_Contract_WP_WO"], 1, 16),
    }

    return {
        "ns":  ns_out,
        "ebit": ebit_out,
        "oi": {
            "cats": cats_oi,
            "ENG":  [safe_float(df_oi.iloc[R["Slide6_ENG"] + _oi_offset, c]) for c in range(4, 18)],
            "MC":   [safe_float(df_oi.iloc[R["Slide6_MC"]  + _oi_offset, c]) for c in range(4, 18)],
            "T&SI": [safe_float(df_oi.iloc[R["Slide6_TSI"] + _oi_offset, c]) for c in range(4, 18)],
            "NUC":  [safe_float(df_oi.iloc[R["Slide6_NUC"] + _oi_offset, c]) for c in range(4, 18)],
        },
        "oi_projects":  oi_projects,
        "wip_projects": wip_projects,
        "bu_ns":        bu_ns,
        "bu_ebit":      bu_ebit,
    }


@st.cache_data(ttl=3600)
def fetch_all_data(months_tuple):
    res = {}
    for sm in months_tuple:
        fp = EXCEL_FILES.get(sm)
        if fp and fp.exists():
            res[sm] = load_month(str(fp))
    return res


@st.cache_data(ttl=3600)
def load_margin_data(filepath: str):
    """Load margin history from 'Ext. Prod.' sheet."""
    try:
        xl = pd.ExcelFile(filepath)
        if "Ext. Prod." not in xl.sheet_names:
            return None

        df_ext = pd.read_excel(xl, sheet_name="Ext. Prod.", header=None, dtype=str)

        header_row = None
        for hr in range(min(10, len(df_ext))):
            h_row = [str(c).strip().upper() for c in df_ext.iloc[hr]]
            if "PROJECT CODE" in h_row and "PROJECT NAME" in h_row:
                header_row = hr
                break
        if header_row is None:
            st.error("Could not locate the header row in 'Ext. Prod.' sheet.")
            return None

        month_row = None
        for rr in range(header_row - 1, max(header_row - 6, -1), -1):
            row = [str(c).strip().upper() for c in df_ext.iloc[rr]]
            if any(m.upper() in " ".join(row) for m in MONTH_ORDER):
                month_row = rr
                break
        if month_row is None:
            st.warning(
                "Could not detect month header row in 'Ext. Prod.' sheet. "
                "Month labels will be inferred from available columns."
            )
            month_row = max(header_row - 2, 0)

        row_months = [str(c).strip() for c in df_ext.iloc[month_row]]
        row_dates  = [str(c).strip() for c in df_ext.iloc[month_row + 1]] if month_row + 1 < len(df_ext) else ["" for _ in row_months]
        row_fields = [str(c).strip().upper() for c in df_ext.iloc[header_row]]

        month_cols    = []
        current_month = None
        for col in range(len(row_fields)):
            cell_month = row_months[col]
            if cell_month:
                month_match = next(
                    (m for m in MONTH_ORDER if m.upper() in cell_month.upper()), None
                )
                if month_match:
                    current_month = month_match
                elif "PREVIOUS YEARS" in cell_month.upper():
                    current_month = None
            if not current_month:
                continue
            if row_fields[col] == "PRODUCTION":
                date_str   = row_dates[col] if col < len(row_dates) else ""
                month_label = _margin_month_label(current_month, date_str)
                month_year  = _parse_margin_year(date_str)
                try:
                    month_index = MONTH_ORDER.index(current_month) + 1
                except ValueError:
                    month_index = 1
                sort_date = datetime(month_year, month_index, 1) if month_year else datetime(1900, month_index, 1)
                month_cols.append({
                    "label":    month_label,
                    "col_idx":  col,
                    "date":     date_str,
                    "sort_key": sort_date,
                })

        bl_col = bo_col = bl_year = bo_year = None
        next_year_col = td_col = None
        for rr in range(min(10, len(df_ext))):
            for col in range(len(df_ext.columns)):
                val      = str(df_ext.iloc[rr, col]).strip().upper()
                bl_match = re.search(r'\b(20\d{2})\s+PRODUCTION\b', val)
                if bl_col is None and bl_match and "END OF" not in val:
                    bl_col  = col
                    bl_year = bl_match.group(1)
                bo_match = re.search(r'\bEND OF\s+(20\d{2})\s+TOTAL PRODUCTION\b', val)
                if bo_col is None and bo_match:
                    bo_col  = col
                    bo_year = bo_match.group(1)
                if next_year_col is None and val == "NEXT YEAR":
                    next_year_col = col
                # "TD" exact match (not TD €, TD $, etc.) — TL To-Date total production
                if td_col is None and val == "TD":
                    td_col = col

        ext_cols = {
            "code":     _c_cfg.get("EXT_PROD_CODE",   1),
            "name":     _c_cfg.get("EXT_PROD_NAME",   2),
            "type":     _c_cfg.get("EXT_PROD_TYPE",   3),
            "client":   _c_cfg.get("EXT_PROD_CLIENT", 4),
            "bu":       _c_cfg.get("EXT_PROD_BU",     6),
            "currency": 7,  # "Original Currency" column
        }
        for hr in range(5):
            if hr < len(df_ext):
                h_row = [str(c).strip().upper() for c in df_ext.iloc[hr]]
                if "PROJECT CODE" in h_row:
                    ext_cols["code"] = h_row.index("PROJECT CODE")
                    if "PROJECT NAME" in h_row:                      ext_cols["name"]     = h_row.index("PROJECT NAME")
                    if "TYPE" in h_row:                              ext_cols["type"]     = h_row.index("TYPE")
                    if "CLIENT/SUBC/ASSOCIATE NAME" in h_row:        ext_cols["client"]   = h_row.index("CLIENT/SUBC/ASSOCIATE NAME")
                    elif "CLIENT" in h_row:                          ext_cols["client"]   = h_row.index("CLIENT")
                    if "BU" in h_row:                                ext_cols["bu"]       = h_row.index("BU")
                    if "ORIGINAL CURRENCY" in h_row:                 ext_cols["currency"] = h_row.index("ORIGINAL CURRENCY")
                    break

        _EXT_ROW_LIMIT = 5000
        if len(df_ext) > _EXT_ROW_LIMIT:
            st.warning(
                f"'Ext. Prod.' sheet has {len(df_ext)} rows but only the first "
                f"{_EXT_ROW_LIMIT} are read."
            )

        projects          = {}
        project_start_row = min(header_row + 1, len(df_ext))
        for ri in range(project_start_row, min(len(df_ext), _EXT_ROW_LIMIT)):
            row       = df_ext.iloc[ri]
            proj_code = str(row.iloc[ext_cols["code"]]).strip() if pd.notna(row.iloc[ext_cols["code"]]) else ""
            proj_name = str(row.iloc[ext_cols["name"]]).strip() if pd.notna(row.iloc[ext_cols["name"]]) else ""
            if not proj_code or not proj_name or proj_name.lower() == "nan" or "total" in proj_name.lower():
                continue
            proj_type = str(row.iloc[ext_cols["type"]]).strip().upper() if ext_cols["type"] < len(row) else ""
            if not proj_type or proj_type == "NAN":
                proj_type = "GROSS FEES"
            client   = str(row.iloc[ext_cols["client"]]).strip()   if pd.notna(row.iloc[ext_cols["client"]])   else ""
            bu_raw   = str(row.iloc[ext_cols["bu"]]).strip()        if pd.notna(row.iloc[ext_cols["bu"]])       else ""
            bu       = BU_ABBREV.get(bu_raw, bu_raw)
            currency = str(row.iloc[ext_cols["currency"]]).strip()  if ext_cols["currency"] < len(row) and pd.notna(row.iloc[ext_cols["currency"]]) else ""

            # Key by (project_code, currency) to keep distinct contracts separate.
            proj_key = f"{proj_code}|{currency}"

            margin_vals = {}
            for month_info in month_cols:
                col_idx = month_info["col_idx"]
                val     = safe_float(row.iloc[col_idx]) if col_idx < len(row) else 0
                margin_vals[month_info["label"]] = val

            bl_val       = safe_float(row.iloc[bl_col])       if bl_col       is not None and bl_col       < len(row) else 0
            bo_val       = safe_float(row.iloc[bo_col])       if bo_col       is not None and bo_col       < len(row) else 0
            next_yr_val  = safe_float(row.iloc[next_year_col]) if next_year_col is not None and next_year_col < len(row) else 0
            td_val       = safe_float(row.iloc[td_col])       if td_col       is not None and td_col       < len(row) else 0
            if bl_year:
                margin_vals[f"BL ({bl_year} Prod)"]      = bl_val
            if bo_year:
                margin_vals[f"BO (End {bo_year} Total)"] = bo_val
            if next_year_col is not None:
                margin_vals["Next Year"]                 = next_yr_val
            if td_col is not None:
                margin_vals["TD (To Date)"]              = td_val

            if proj_key not in projects:
                projects[proj_key] = {
                    "code":        proj_code,
                    "name":        proj_name,
                    "client":      client,
                    "bu":          bu,
                    "currency":    currency,
                    "margin_data": {},   # GROSS FEES only — backward compat
                    "type_data":   {},   # all types keyed by type name
                }

            # Accumulate into type_data for every type
            if proj_type not in projects[proj_key]["type_data"]:
                projects[proj_key]["type_data"][proj_type] = {}
            for k, v in margin_vals.items():
                projects[proj_key]["type_data"][proj_type][k] = (
                    projects[proj_key]["type_data"][proj_type].get(k, 0) + v
                )

            # Keep margin_data as GROSS FEES only (used by WIP view and existing merge logic)
            if proj_type == "GROSS FEES":
                for k, v in margin_vals.items():
                    projects[proj_key]["margin_data"][k] = (
                        projects[proj_key]["margin_data"].get(k, 0) + v
                    )

        return {
            "projects":     projects,
            "month_cols":   month_cols,
            "bl_col":       bl_col,
            "bo_col":       bo_col,
            "bl_year":      bl_year,
            "bo_year":      bo_year,
            "next_year_col": next_year_col,
            "td_col":       td_col,
        }
    except Exception as e:
        import traceback
        st.error(f"Error loading margin data from {Path(filepath).name}:\n{e}\n\n{traceback.format_exc()}")
        return None


def merge_margin_data(selected_months):
    all_projects   = {}
    all_month_cols = []
    bl_year = bo_year = None

    for sm in selected_months:
        margin_file = EXCEL_FILES.get(sm)
        if not margin_file or not margin_file.exists():
            continue
        margin_data = load_margin_data(str(margin_file))
        if not margin_data or not margin_data.get("projects"):
            continue
        for proj_key, proj_info in margin_data["projects"].items():
            if proj_key not in all_projects:
                all_projects[proj_key] = {
                    "code":     proj_info["code"],
                    "name":     proj_info["name"],
                    "client":   proj_info["client"],
                    "bu":       proj_info["bu"],
                    "currency": proj_info["currency"],
                    "margin_data": {},
                    "type_data":   {},
                }
            all_projects[proj_key]["margin_data"].update(proj_info["margin_data"])
            for t_name, t_vals in proj_info.get("type_data", {}).items():
                if t_name not in all_projects[proj_key]["type_data"]:
                    all_projects[proj_key]["type_data"][t_name] = {}
                all_projects[proj_key]["type_data"][t_name].update(t_vals)
        all_month_cols.extend(margin_data["month_cols"])
        if margin_data.get("bl_year"):
            bl_year = margin_data["bl_year"]
        if margin_data.get("bo_year"):
            bo_year = margin_data["bo_year"]

    seen_labels       = set()
    unique_month_cols = []
    for mc in all_month_cols:
        if mc["label"] not in seen_labels:
            seen_labels.add(mc["label"])
            unique_month_cols.append(mc)

    return {"projects": all_projects, "month_cols": unique_month_cols, "bl_year": bl_year, "bo_year": bo_year}


# ── Chart helpers ──────────────────────────────────────────────────────────────
def chart_stacked(data, title, keys, color_palette=None):
    n    = min(len(data["cats"]), min(len(data[k]) for k in keys))
    df   = pd.DataFrame({"Category": data["cats"][:n]})
    for k in keys:
        df[k] = data[k][:n]
    df_m = df.melt("Category", var_name="Tier", value_name="Value")
    colors = color_palette or (
        [BU_COLORS.get(k, "#0EA5E9") for k in keys]
        if any(k in BU_COLORS for k in keys)
        else ["#0EA5E9", "#10B981", "#F59E0B", "#E11D48"]
    )
    fig = px.bar(
        df_m, x="Category", y="Value", color="Tier", title=title,
        color_discrete_sequence=colors, barmode="stack",
        hover_data={"Value": ":,.0f"},
    )
    fig.update_layout(
        plot_bgcolor="white",
        font=dict(family="Segoe UI", color=PRIMARY),
        xaxis_tickangle=-45,
        margin=dict(t=80, b=20),
        uniformtext_minsize=8, uniformtext_mode="hide",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        legend_title_text="",
    )
    fig.update_traces(texttemplate="%{y:,.0f}", textposition="inside", insidetextanchor="middle")
    return fig


def chart_grouped(data, title, keys, color_palette=None):
    n    = min(len(data["cats"]), min(len(data[k]) for k in keys))
    df   = pd.DataFrame({"Category": data["cats"][:n]})
    for k in keys:
        df[k] = data[k][:n]
    df_m   = df.melt("Category", var_name="Tier", value_name="Value")
    colors = color_palette or ["#0EA5E9", "#10B981", "#F59E0B"]
    fig    = px.bar(
        df_m, x="Category", y="Value", color="Tier", title=title,
        color_discrete_sequence=colors,
        barmode="group", hover_data={"Value": ":,.0f"},
    )
    fig.update_layout(
        plot_bgcolor="white",
        font=dict(family="Segoe UI", color=PRIMARY),
        xaxis_tickangle=-45,
        margin=dict(t=80, b=20),
        uniformtext_minsize=8, uniformtext_mode="hide",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        legend_title_text="",
    )
    fig.update_traces(texttemplate="%{y:,.0f}", textposition="inside", insidetextanchor="middle")
    return fig


def chart_projects(projects, title):
    if not projects:
        return None
    df          = pd.DataFrame(projects[:30])
    df["label"] = df["value"].apply(human_k)
    fig = px.bar(
        df, x="value", y="project", color="bu", orientation="h",
        title=title, text="label",
        color_discrete_map=BU_COLORS,
        category_orders={"project": df["project"].tolist()},
        hover_data={"client": True, "value": ":,.0f"},
    )
    fig.update_layout(
        plot_bgcolor="white",
        font=dict(family="Segoe UI", color=PRIMARY),
        yaxis_title="", xaxis_title="kTL",
        margin=dict(t=60, r=130),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        height=max(400, len(df) * 28),
    )
    fig.update_traces(textposition="outside", cliponaxis=False)
    return fig


def chart_wip(wip_list, title):
    if not wip_list:
        return None
    df          = pd.DataFrame(wip_list[:30])
    df["label"] = df["wip_tl"].apply(human_tl)
    fig = px.bar(
        df, x="wip_tl", y="name", color="bu", orientation="h",
        title=title, text="label",
        color_discrete_map=BU_COLORS,
        category_orders={"name": df["name"].tolist()},
        hover_data={"client": True, "orig_currency": True, "wip_tl": ":,.0f"},
    )
    fig.update_layout(
        plot_bgcolor="white",
        font=dict(family="Segoe UI", color=PRIMARY),
        yaxis_title="", xaxis_title="TL",
        margin=dict(t=60, r=130),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        height=max(400, len(df) * 28),
    )
    fig.update_traces(textposition="outside", cliponaxis=False)
    return fig


# ── App ────────────────────────────────────────────────────────────────────────
st.title("MRC Business Review Dashboard")

if not EXCEL_FILES:
    st.error(
        "No Excel files found. Place `*_Project_Budget_Analysis_*_v*.xlsx` files "
        "in the same folder as this script."
    )
    st.stop()

# ── Sidebar 1: Period ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Period")
    all_years    = sorted(set(k.split()[1] for k in EXCEL_FILES))
    selected_year = st.selectbox("Year", all_years, index=len(all_years) - 1)

    months_for_year = sorted(
        [k for k in EXCEL_FILES if k.endswith(selected_year)],
        key=lambda x: MONTH_ORDER.index(x.split()[0]),
    )

    compare_mode      = st.checkbox("Compare months", value=False)
    year_compare_mode = st.checkbox("Compare years",  value=False)

    # Year comparison selector — always shown when "Compare years" is checked
    comparison_year     = None
    months_for_comparison = []
    if year_compare_mode:
        compare_year_options = sorted([y for y in all_years if y != selected_year], reverse=True)
        if compare_year_options:
            comparison_year = st.selectbox("Compare with year:", compare_year_options)
            months_for_comparison = sorted(
                [k for k in EXCEL_FILES if k.endswith(comparison_year)],
                key=lambda x: MONTH_ORDER.index(x.split()[0]),
            )
        else:
            st.info(f"No other year available to compare with {selected_year}.")

    # Month selection
    if compare_mode:
        default_sel = (
            [months_for_year[-2], months_for_year[-1]]
            if len(months_for_year) > 1
            else [months_for_year[-1]]
        )
        selected_months_raw = st.multiselect(
            "Select Months", months_for_year, default=default_sel,
            format_func=lambda x: x.split()[0],
        )
        selected_months = sorted(
            selected_months_raw, key=lambda x: MONTH_ORDER.index(x.split()[0])
        )
        # If year comparison is also active, add same months from the comparison year
        if year_compare_mode and comparison_year:
            for m in list(selected_months):
                month_name  = m.split()[0]
                comp_matches = [cm for cm in months_for_comparison if cm.startswith(month_name)]
                if comp_matches:
                    if comp_matches[0] not in selected_months:
                        selected_months.append(comp_matches[0])
                else:
                    st.warning(f"No {month_name} file found for {comparison_year}.")
            # Sort by (year, month) so same-year months are grouped together
            selected_months = sorted(
                selected_months,
                key=lambda x: (x.split()[1], MONTH_ORDER.index(x.split()[0])),
            )
    else:
        single_month = st.selectbox(
            "Month", months_for_year,
            index=len(months_for_year) - 1,
            format_func=lambda x: x.split()[0],
        )
        selected_months = [single_month] if single_month else []

        # If year comparison enabled, also add same month from comparison year
        if year_compare_mode and comparison_year and selected_months:
            comp_month_name = selected_months[0].split()[0]
            comp_months     = [m for m in months_for_comparison if m.startswith(comp_month_name)]
            if comp_months:
                selected_months.append(comp_months[0])
            else:
                st.warning(
                    f"No {comp_month_name} file found for {comparison_year}. "
                    "Year comparison cannot be displayed."
                )

if not selected_months:
    st.warning("Please select at least one month from the sidebar.")
    st.stop()

with st.spinner("Loading data..."):
    loaded_data = fetch_all_data(tuple(selected_months))

# Build project/client sets for tab-level filter widgets (kept separate by tab)
all_oi_projs   = set()
all_oi_clients = set()
all_wip_projs  = set()
all_wip_clients = set()
for data in loaded_data.values():
    if not data:
        continue
    for p in data.get("oi_projects", []):
        if p.get("project"): all_oi_projs.add(p["project"])
        if p.get("client"):  all_oi_clients.add(p["client"])
    for p in data.get("wip_projects", []):
        if p.get("name"):   all_wip_projs.add(p["name"])
        if p.get("client"): all_wip_clients.add(p["client"])

# ── Sidebar 2: Generate / Refresh ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("### Generate Report")
    st.caption("Select a single month to generate PPTX:")
    gen_month = st.selectbox(
        "PPTX Month",
        selected_months if selected_months else months_for_year,
        format_func=lambda x: x.split()[0],
    )
    sel_file = EXCEL_FILES.get(gen_month)
    st.caption(f"File: {sel_file.name if sel_file else '—'}")
    run_btn = st.button("▶ Generate PPTX", type="primary", use_container_width=True)

    st.markdown("---")
    if st.button("🔄 Refresh Files", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ── Run report ─────────────────────────────────────────────────────────────────
if run_btn and sel_file:
    st.markdown("---")
    with st.spinner(f"Generating PPTX for {gen_month}…"):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_DIR / "generate_full_mbr_stacked.py"), str(sel_file)],
            capture_output=True, text=True, cwd=str(SCRIPT_DIR),
        )
    if result.returncode == 0:
        st.success(f"Report generated successfully for {gen_month}!")
        done_line = next(
            (l.strip() for l in result.stdout.splitlines() if l.strip().startswith("[DONE]")),
            None,
        )
        if done_line:
            st.info(done_line.removeprefix("[DONE]").strip())
        if result.stdout:
            with st.expander("Show generation log", expanded=False):
                st.code(result.stdout, language="")
    else:
        st.error("Generator encountered an error — see details below.")
        if result.stdout:
            st.code(result.stdout, language="")
        if result.stderr:
            with st.expander("Error details", expanded=True):
                st.code(result.stderr, language="")
    st.markdown("---")


# ── Filter helpers ─────────────────────────────────────────────────────────────
def filter_oi(projects, proj_sel=None, proj_text="", cli_sel=None, cli_text=""):
    """Filter OI project list. Multiselect (exact) takes priority over text (substring)."""
    out = projects
    if proj_sel:
        out = [p for p in out if p.get("project") in proj_sel]
    elif proj_text:
        out = [p for p in out if proj_text.lower() in (p.get("project") or "").lower()]
    if cli_sel:
        out = [p for p in out if p.get("client") in cli_sel]
    elif cli_text:
        out = [p for p in out if cli_text.lower() in (p.get("client") or "").lower()]
    return out


def filter_wip(wip, proj_sel=None, proj_text="", cli_sel=None, cli_text=""):
    """Filter WIP project list. Multiselect (exact) takes priority over text (substring)."""
    out = wip
    if proj_sel:
        out = [p for p in out if p.get("name") in proj_sel]
    elif proj_text:
        out = [p for p in out if proj_text.lower() in (p.get("name") or "").lower()]
    if cli_sel:
        out = [p for p in out if p.get("client") in cli_sel]
    elif cli_text:
        out = [p for p in out if cli_text.lower() in (p.get("client") or "").lower()]
    return out


# ── Render helpers ─────────────────────────────────────────────────────────────
def render_ns_view(sm, data, prev_sm, prev_data, is_total, bu_name):
    if is_total:
        ns     = data["ns"]
        k_list = ["Contract", "WP", "WO"]
        p_ns   = prev_data["ns"] if prev_data else None
    else:
        ns     = data["bu_ns"][bu_name]
        k_list = ["Order", "Offer", "Opp"]
        p_ns   = prev_data["bu_ns"][bu_name] if prev_data else None

    month_abbr = sm.split()[0][:3]
    ns_idx     = _cat_index(ns["cats"], month_abbr)

    # Metric cards show per-month delta; charts use the full cumulative array
    totals = {k: monthly_val(ns[k], ns["cats"], ns_idx) for k in k_list}
    totals["Total"] = sum(totals.values())

    prev_totals = {}
    if p_ns and prev_sm:
        p_abbr = prev_sm.split()[0][:3]
        p_idx  = _cat_index(p_ns["cats"], p_abbr)
        if p_idx != -1:
            prev_totals = {k: monthly_val(p_ns[k], p_ns["cats"], p_idx) for k in k_list}
            prev_totals["Total"] = sum(prev_totals.values())

    all_metric_keys = k_list + ["Total"]
    mcols = st.columns(len(all_metric_keys))
    for idx_m, k in enumerate(all_metric_keys):
        val   = totals[k]
        delta = val - prev_totals.get(k, 0) if prev_totals else None
        mcols[idx_m].metric(k, human_k(val), delta=human_k(delta) if delta is not None else None)

    palette = None if is_total else bu_tier_colors(bu_name)
    st.plotly_chart(
        chart_stacked(ns, "", k_list, palette),
        use_container_width=True, key=f"ns_{bu_name}_{sm}",
    )


def render_ebit_view(sm, data, prev_sm, prev_data, is_total, bu_name):
    month_abbr = sm.split()[0][:3]
    if is_total:
        ebit   = data["ebit"]
        k_list = ["Contract", "Contract+WP", "Contract+WP+WO"]
        p_ebit = prev_data["ebit"] if prev_data else None
    else:
        ebit   = data["bu_ebit"][bu_name]
        k_list = ["Total"]
        p_ebit = prev_data["bu_ebit"][bu_name] if prev_data else None

    ebit_idx      = _cat_index(ebit["cats"], month_abbr)
    prev_ebit_idx = -1
    if p_ebit and prev_sm:
        prev_abbr     = prev_sm.split()[0][:3]
        prev_ebit_idx = _cat_index(p_ebit["cats"], prev_abbr)

    # Metric cards show per-month delta; charts use the full cumulative array
    mcols = st.columns(min(3, len(k_list)))
    for idx_m, k in enumerate(k_list):
        val   = monthly_val(ebit[k], ebit["cats"], ebit_idx)
        p_val = monthly_val(p_ebit[k], p_ebit["cats"], prev_ebit_idx) if p_ebit and prev_ebit_idx != -1 else None
        delta = val - p_val if p_val is not None else None
        mcols[idx_m % len(mcols)].metric(k, human_k(val), delta=human_k(delta) if delta is not None else None)

    ebit_palette = None if is_total else [BU_COLORS.get(bu_name, "#0EA5E9")]
    st.plotly_chart(
        chart_grouped(ebit, "", k_list, ebit_palette),
        use_container_width=True, key=f"ebit_{bu_name}_{sm}",
    )


def render_oi_view(sm, data, prev_sm, prev_data, is_total, bu_name,
                   proj_sel=None, proj_text="", cli_sel=None, cli_text=""):
    if is_total:
        oi         = data["oi"]
        keys_oi    = ["ENG", "MC", "T&SI", "NUC"]
        month_abbr = sm.split()[0][:3]
        oi_idx     = _cat_index(oi["cats"], month_abbr)

        prev_oi     = prev_data["oi"] if prev_data else None
        prev_oi_idx = -1
        if prev_oi and prev_sm:
            prev_abbr   = prev_sm.split()[0][:3]
            prev_oi_idx = _cat_index(prev_oi["cats"], prev_abbr)

        # 4-column metric grid (one per BU)
        mcols = st.columns(4)
        for idx_m, k in enumerate(keys_oi):
            val   = oi[k][oi_idx] if oi_idx != -1 else 0
            delta = val - prev_oi[k][prev_oi_idx] if prev_oi and prev_oi_idx != -1 else None
            mcols[idx_m].metric(k, human_k(val), delta=human_k(delta) if delta is not None else None)

        # Stacked OI chart (all BUs combined over time)
        st.plotly_chart(
            chart_stacked(oi, "", keys_oi, [BU_COLORS.get(k, "#888") for k in keys_oi]),
            use_container_width=True, key=f"oi_chart_total_{sm}",
        )
        bg_oi = filter_oi(data["oi_projects"], proj_sel, proj_text, cli_sel, cli_text)
    else:
        bg_oi    = [p for p in filter_oi(data["oi_projects"], proj_sel, proj_text, cli_sel, cli_text) if p["bu"] == bu_name]
        total_oi = sum(p["value"] for p in bg_oi)
        delta_oi = None
        if prev_data:
            p_bg_oi  = [p for p in filter_oi(prev_data["oi_projects"], proj_sel, proj_text, cli_sel, cli_text) if p["bu"] == bu_name]
            delta_oi = total_oi - sum(p["value"] for p in p_bg_oi)
        st.metric(
            f"Total Order Intake ({len(bg_oi)} projs)",
            human_k(total_oi),
            delta=human_k(delta_oi) if delta_oi is not None else None,
        )
        fig_bu_oi = chart_projects(bg_oi, "")
        if fig_bu_oi:
            st.plotly_chart(fig_bu_oi, use_container_width=True, key=f"oi_bu_fig_{bu_name}_{sm}")

    if bg_oi:
        df_oi_disp = pd.DataFrame(bg_oi)
        df_oi_disp["Value (kTL)"] = df_oi_disp["value"].apply(human_k)
        cols_show = (
            ["project", "client", "bu", "Value (kTL)"]
            if is_total
            else ["project", "client", "Value (kTL)"]
        )
        st.dataframe(
            df_oi_disp[cols_show], hide_index=True,
            use_container_width=True, key=f"oi_df_{bu_name}_{sm}",
        )
        st.download_button(
            "⬇ Export to Excel",
            df_to_excel_bytes(df_oi_disp[cols_show]),
            file_name=f"OI_{sm.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"oi_export_{bu_name}_{sm}",
        )
    else:
        st.info("No projects match the current filters.")


def render_wip_view(sm, data, prev_sm, prev_data, is_total, bu_name,
                    proj_sel=None, proj_text="", cli_sel=None, cli_text=""):
    if is_total:
        wip       = filter_wip(data["wip_projects"], proj_sel, proj_text, cli_sel, cli_text)
        total_wip = sum(p["wip_tl"] for p in wip)
        delta_wip = None
        if prev_data:
            prev_total = sum(p["wip_tl"] for p in filter_wip(prev_data["wip_projects"], proj_sel, proj_text, cli_sel, cli_text))
            delta_wip  = human_tl(total_wip - prev_total)
    else:
        wip       = [p for p in filter_wip(data["wip_projects"], proj_sel, proj_text, cli_sel, cli_text) if p["bu"] == bu_name]
        total_wip = sum(p["wip_tl"] for p in wip)
        delta_wip = None
        if prev_data:
            prev_wip  = [p for p in filter_wip(prev_data["wip_projects"], proj_sel, proj_text, cli_sel, cli_text) if p["bu"] == bu_name]
            delta_wip = human_tl(total_wip - sum(p["wip_tl"] for p in prev_wip))

    st.metric(f"Total WIP ({len(wip)} projs)", human_tl(total_wip), delta=delta_wip)

    fig = chart_wip(wip, "")
    if fig:
        st.plotly_chart(fig, use_container_width=True, key=f"wip_fig_{bu_name}_{sm}")
    if wip:
        df_disp          = pd.DataFrame(wip)
        df_disp["WIP TL"] = df_disp["wip_tl"].apply(human_tl)
        cols_show = (
            ["name", "client", "bu", "orig_currency", "WIP TL"]
            if is_total
            else ["name", "client", "orig_currency", "WIP TL"]
        )
        st.dataframe(
            df_disp[cols_show], hide_index=True,
            use_container_width=True, key=f"wip_df_{bu_name}_{sm}",
        )
        st.download_button(
            "⬇ Export to Excel",
            df_to_excel_bytes(df_disp[cols_show]),
            file_name=f"WIP_{sm.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"wip_export_{bu_name}_{sm}",
        )

        with st.expander("💱 WIP by Currency", expanded=False):
            cur_totals = {}
            for p in wip:
                cur = p.get("orig_currency") or "Unknown"
                cur_totals[cur] = cur_totals.get(cur, 0) + p["wip_tl"]
            df_cur = pd.DataFrame([
                {"Currency": k, "WIP TL": v, "Formatted": human_tl(v)}
                for k, v in sorted(cur_totals.items(), key=lambda x: -x[1])
            ])
            cc1, cc2 = st.columns(2)
            with cc1:
                fig_cpie = px.pie(
                    df_cur, values="WIP TL", names="Currency",
                    title="WIP share by currency",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig_cpie.update_traces(textinfo="label+percent")
                fig_cpie.update_layout(margin=dict(t=40, b=0))
                st.plotly_chart(fig_cpie, use_container_width=True, key=f"cur_pie_{bu_name}_{sm}")
            with cc2:
                fig_cbar = px.bar(
                    df_cur, x="WIP TL", y="Currency", orientation="h",
                    text="Formatted", title="WIP TL by currency",
                    color="Currency",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig_cbar.update_layout(
                    plot_bgcolor="white", showlegend=False,
                    font=dict(family="Segoe UI", color=PRIMARY),
                    margin=dict(t=40, b=0),
                )
                fig_cbar.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(fig_cbar, use_container_width=True, key=f"cur_bar_{bu_name}_{sm}")
            st.dataframe(
                df_cur[["Currency", "Formatted"]].rename(columns={"Formatted": "WIP TL"}),
                hide_index=True, use_container_width=True,
            )


def _month_cols_iter(months, data_map):
    """Yield (col, sm, data, prev_sm, prev_data) for each month column."""
    if len(months) <= 3:
        cols = st.columns(len(months))
        for i, sm in enumerate(months):
            data = data_map.get(sm)
            if not data:
                continue
            prev_sm   = months[i - 1] if i > 0 else None
            prev_data = data_map.get(prev_sm) if prev_sm else None
            yield cols[i], sm, data, prev_sm, prev_data
    else:
        tab_names  = [sm.split()[0] for sm in months]
        month_tabs = st.tabs(tab_names)
        for i, (tab, sm) in enumerate(zip(month_tabs, months)):
            data = data_map.get(sm)
            if not data:
                continue
            prev_sm   = months[i - 1] if i > 0 else None
            prev_data = data_map.get(prev_sm) if prev_sm else None
            with tab:
                yield None, sm, data, prev_sm, prev_data


# ── Main UI ───────────────────────────────────────────────────────────────────
global_bu_view = st.radio(
    "Business Unit View", ["Company Total", "ENG", "MC", "T&SI", "NUC"], horizontal=True,
    help="Applies to all tabs. 'Company Total' shows aggregated figures; selecting a BU filters to that unit's data.",
)
st.markdown("---")
is_total = global_bu_view == "Company Total"

# Show year in subheaders whenever multiple years are loaded simultaneously
_multi_year = len(set(sm.split()[1] for sm in selected_months if len(sm.split()) > 1)) > 1
_sm_label   = lambda sm: sm if _multi_year else sm.split()[0]

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Net Sales", "EBIT", "Order Intake", "WIP", "Project History", "Scenarios"])

# ── Tab 1: Net Sales ──────────────────────────────────────────────────────────
with tab1:
    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        if col:
            with col:
                st.subheader(f"{_sm_label(sm)} Net Sales")
                render_ns_view(sm, data, prev_sm, prev_data, is_total, global_bu_view)
        else:
            st.subheader(f"{_sm_label(sm)} Net Sales")
            render_ns_view(sm, data, prev_sm, prev_data, is_total, global_bu_view)

    if len(selected_months) == 2:
        _sm1, _sm2   = selected_months[0], selected_months[1]
        _d1, _d2     = loaded_data.get(_sm1), loaded_data.get(_sm2)
        if _d1 and _d2:
            _is_yoy  = _sm1.split()[1] != _sm2.split()[1]
            _tag     = "YoY" if _is_yoy else "MoM"
            st.markdown(f"**{_tag} Δ% — {_sm1} → {_sm2}**")
            _ns1, _ns2 = _d1["ns"], _d2["ns"]
            _i1 = _cat_index(_ns1["cats"], _sm1.split()[0][:3])
            _i2 = _cat_index(_ns2["cats"], _sm2.split()[0][:3])
            _k_ns = ["Contract", "WP", "WO"]
            _dcols = st.columns(len(_k_ns) + 1)
            for _ci, _k in enumerate(_k_ns + ["Total"]):
                _v1 = sum(monthly_val(_ns1[kk], _ns1["cats"], _i1) for kk in _k_ns) if _k == "Total" else monthly_val(_ns1[_k], _ns1["cats"], _i1)
                _v2 = sum(monthly_val(_ns2[kk], _ns2["cats"], _i2) for kk in _k_ns) if _k == "Total" else monthly_val(_ns2[_k], _ns2["cats"], _i2)
                _pct = (_v2 - _v1) / abs(_v1) * 100 if _v1 else None
                _dcols[_ci].metric(_k, f"{_pct:+.1f}%" if _pct is not None else "—",
                                   delta=human_k(_v2 - _v1) if _v1 else None)

# ── Tab 2: EBIT ───────────────────────────────────────────────────────────────
with tab2:
    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        if col:
            with col:
                st.subheader(f"{_sm_label(sm)} EBIT")
                render_ebit_view(sm, data, prev_sm, prev_data, is_total, global_bu_view)
        else:
            st.subheader(f"{_sm_label(sm)} EBIT")
            render_ebit_view(sm, data, prev_sm, prev_data, is_total, global_bu_view)

    if len(selected_months) == 2:
        _sm1, _sm2   = selected_months[0], selected_months[1]
        _d1, _d2     = loaded_data.get(_sm1), loaded_data.get(_sm2)
        if _d1 and _d2:
            _is_yoy  = _sm1.split()[1] != _sm2.split()[1]
            _tag     = "YoY" if _is_yoy else "MoM"
            st.markdown(f"**{_tag} Δ% — {_sm1} → {_sm2}**")
            _eb1, _eb2  = _d1["ebit"], _d2["ebit"]
            _i1 = _cat_index(_eb1["cats"], _sm1.split()[0][:3])
            _i2 = _cat_index(_eb2["cats"], _sm2.split()[0][:3])
            _k_ebit = ["Contract", "Contract+WP", "Contract+WP+WO"]
            _dcols  = st.columns(len(_k_ebit))
            for _ci, _k in enumerate(_k_ebit):
                _v1 = monthly_val(_eb1[_k], _eb1["cats"], _i1)
                _v2 = monthly_val(_eb2[_k], _eb2["cats"], _i2)
                _pct = (_v2 - _v1) / abs(_v1) * 100 if _v1 else None
                _dcols[_ci].metric(_k, f"{_pct:+.1f}%" if _pct is not None else "—",
                                   delta=human_k(_v2 - _v1) if _v1 else None)

# ── Tab 3: Order Intake ───────────────────────────────────────────────────────
with tab3:
    with st.expander("🔍 Filter by Project / Client", expanded=False):
        fc1, fc2 = st.columns(2)
        with fc1:
            oi_proj_text = st.text_input(
                "Search Project", placeholder="Type to narrow the list…", key="oi_proj_text",
                help="Type any part of the project name to narrow the dropdown below.",
            )
            oi_proj_opts = (
                [p for p in sorted(all_oi_projs) if oi_proj_text.lower() in p.lower()]
                if oi_proj_text else sorted(all_oi_projs)
            )
            oi_proj_sel = st.multiselect(
                "Pin Projects", oi_proj_opts, key="oi_proj_ms",
                help="Select one or more projects to show only those rows. Overrides the text search above.",
            )
        with fc2:
            oi_cli_text = st.text_input(
                "Search Client", placeholder="Type to narrow the list…", key="oi_cli_text",
                help="Type any part of the client name to narrow the dropdown below.",
            )
            oi_cli_opts = (
                [c for c in sorted(all_oi_clients) if oi_cli_text.lower() in c.lower()]
                if oi_cli_text else sorted(all_oi_clients)
            )
            oi_cli_sel = st.multiselect(
                "Pin Clients", oi_cli_opts, key="oi_cli_ms",
                help="Select one or more clients to show only their rows. Overrides the text search above.",
            )

    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        if col:
            with col:
                st.subheader(f"{_sm_label(sm)} Order Intake")
                render_oi_view(sm, data, prev_sm, prev_data, is_total, global_bu_view,
                               proj_sel=oi_proj_sel, proj_text=oi_proj_text,
                               cli_sel=oi_cli_sel, cli_text=oi_cli_text)
        else:
            st.subheader(f"{_sm_label(sm)} Order Intake")
            render_oi_view(sm, data, prev_sm, prev_data, is_total, global_bu_view,
                           proj_sel=oi_proj_sel, proj_text=oi_proj_text,
                           cli_sel=oi_cli_sel, cli_text=oi_cli_text)

# ── Tab 4: WIP ────────────────────────────────────────────────────────────────
with tab4:
    with st.expander("🔍 Filter by Project / Client", expanded=False):
        wc1, wc2 = st.columns(2)
        with wc1:
            wip_proj_text = st.text_input(
                "Search Project", placeholder="Type to narrow the list…", key="wip_proj_text",
                help="Type any part of the project name to narrow the dropdown below.",
            )
            wip_proj_opts = (
                [p for p in sorted(all_wip_projs) if wip_proj_text.lower() in p.lower()]
                if wip_proj_text else sorted(all_wip_projs)
            )
            wip_proj_sel = st.multiselect(
                "Pin Projects", wip_proj_opts, key="wip_proj_ms",
                help="Select one or more projects to show only those rows. Overrides the text search above.",
            )
        with wc2:
            wip_cli_text = st.text_input(
                "Search Client", placeholder="Type to narrow the list…", key="wip_cli_text",
                help="Type any part of the client name to narrow the dropdown below.",
            )
            wip_cli_opts = (
                [c for c in sorted(all_wip_clients) if wip_cli_text.lower() in c.lower()]
                if wip_cli_text else sorted(all_wip_clients)
            )
            wip_cli_sel = st.multiselect(
                "Pin Clients", wip_cli_opts, key="wip_cli_ms",
                help="Select one or more clients to show only their rows. Overrides the text search above.",
            )

    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        if col:
            with col:
                st.subheader(f"{_sm_label(sm)} WIP")
                render_wip_view(sm, data, prev_sm, prev_data, is_total, global_bu_view,
                                proj_sel=wip_proj_sel, proj_text=wip_proj_text,
                                cli_sel=wip_cli_sel, cli_text=wip_cli_text)
        else:
            st.subheader(f"{_sm_label(sm)} WIP")
            render_wip_view(sm, data, prev_sm, prev_data, is_total, global_bu_view,
                            proj_sel=wip_proj_sel, proj_text=wip_proj_text,
                            cli_sel=wip_cli_sel, cli_text=wip_cli_text)

# ── Tab fragments ─────────────────────────────────────────────────────────────
# Each fragment reruns independently on internal widget changes so the outer
# st.tabs() selection is never reset when the user interacts with filters.

@st.fragment
def _render_project_history(loaded_data, selected_months, global_bu_view, multi_year):
    st.markdown("### Project History")
    st.caption(
        "Track project-based monthly Production (TL) and project WIP TL history. "
        "Production values are loaded from 'Ext. Prod.' and WIP values from the monthly WIP data."
    )

    with st.spinner("Loading margin data…"):
        margin_data = merge_margin_data(selected_months)

    if margin_data and margin_data.get("projects"):
        projects    = margin_data["projects"]
        month_cols  = margin_data["month_cols"]
        bl_year     = margin_data.get("bl_year")
        bo_year     = margin_data.get("bo_year")

        st.markdown("#### Project Filters")
        pf1, pf2 = st.columns(2)
        with pf1:
            ph_bu_filter = st.multiselect(
                "Filter by Business Unit",
                ["ENG", "MC", "T&SI", "NUC"],
                default=["ENG", "MC", "T&SI", "NUC"],
                key="ph_bu_filter",
            )
        with pf2:
            _all_ph_clients = sorted(
                {info["client"] for info in projects.values()
                 if info["client"] and info["client"].lower() not in ("", "nan")}
            )
            ph_cli_text = st.text_input(
                "Search Client", placeholder="Filter by client…", key="ph_cli_text",
            )
            _cli_opts = [c for c in _all_ph_clients if ph_cli_text.lower() in c.lower()] if ph_cli_text else _all_ph_clients
            ph_cli_filter = st.multiselect("Filter by Client", _cli_opts, key="ph_cli_filter")

        filtered_projects = {
            key: info for key, info in projects.items()
            if info["bu"] in ph_bu_filter
            and (not ph_cli_filter or info["client"] in ph_cli_filter)
        }

        # Build display labels: "Project Name" or "Project Name (CURRENCY)" when the same name
        # appears under different currencies (different contract lines).
        _name_counts: dict[str, int] = {}
        for info in filtered_projects.values():
            _name_counts[info["name"]] = _name_counts.get(info["name"], 0) + 1

        def _display_label(info: dict) -> str:
            name = info["name"]
            return f"{name} ({info['currency']})" if _name_counts.get(name, 0) > 1 else name

        # Map display_label → internal key for selectbox
        label_to_key  = {_display_label(v): k for k, v in filtered_projects.items()}
        all_labels    = sorted(label_to_key.keys())

        sel_label = None
        if not all_labels:
            st.warning("No projects found for the selected Business Units.")
        else:
            ph_proj_text = st.text_input(
                "🔍 Search Project", placeholder="Type to narrow the list…", key="ph_proj_text",
                help="Type any part of the project name to filter the dropdown below.",
            )
            visible_labels = (
                [l for l in all_labels if ph_proj_text.lower() in l.lower()]
                if ph_proj_text else all_labels
            )
            sel_label = st.selectbox("Select Project", visible_labels if visible_labels else all_labels, key="margin_proj")

        sel_proj  = label_to_key.get(sel_label) if sel_label else None

        if sel_proj and sel_proj in filtered_projects:
            proj_info = filtered_projects[sel_proj]
            st.markdown(
                f"**Code:** {proj_info['code']} | **Client:** {proj_info['client']} "
                f"| **BU:** {proj_info['bu']} | **Currency:** {proj_info['currency']}"
            )

            _type_order = [
                "GROSS FEES", "REIMBURSABLES", "SUBCON 1", "SUBCON 2",
                "ASSOCIATE 1", "ENG. COST", "PROJECT EXP", "PROJECT EXP ACCR",
            ]
            _available_types = list(proj_info.get("type_data", {}).keys())
            _available_types_sorted = (
                [t for t in _type_order if t in _available_types]
                + [t for t in _available_types if t not in _type_order]
            )
            if not _available_types_sorted:
                _available_types_sorted = ["GROSS FEES"]

            ph_view_col, ph_type_col = st.columns([2, 2])
            with ph_view_col:
                ph_view = st.radio(
                    "View",
                    ["Production (Ext. Prod.)", "WIP TL"],
                    horizontal=True,
                    key="ph_view",
                )
            with ph_type_col:
                ph_cost_type = st.selectbox(
                    "Cost Type",
                    _available_types_sorted,
                    index=0,
                    key="ph_cost_type",
                    help="Filter production data by cost type. GROSS FEES is the main project total.",
                ) if ph_view == "Production (Ext. Prod.)" else "GROSS FEES"

            _active_margin_data = proj_info.get("type_data", {}).get(
                ph_cost_type, proj_info["margin_data"]
            )

            # ── Build WIP history (needed by both views for combined table) ────
            wip_history = []
            for _sm in selected_months:
                d = loaded_data.get(_sm)
                if not d:
                    continue
                wip_projs = filter_wip(d["wip_projects"])
                if global_bu_view != "Company Total":
                    wip_projs = [p for p in wip_projs if p["bu"] == global_bu_view]
                total_wip = sum(
                    p["wip_tl"] for p in wip_projs
                    if str(p["name"]).strip().lower() == str(proj_info["name"]).strip().lower()
                )
                wip_history.append({"Period": _sm, "WIP TL": total_wip})

            if ph_view == "Production (Ext. Prod.)":
                month_cols_sorted = sorted(month_cols, key=lambda m: m["sort_key"])

                margin_rows = []
                for month_info in month_cols_sorted:
                    label = month_info["label"]
                    val   = _active_margin_data.get(label, 0)
                    margin_rows.append({"Period": label, "Production (TL)": val})

                bl_label = f"BL ({bl_year} Prod)"      if bl_year else "BL (Year Prod)"
                bo_label = f"BO (End {bo_year} Total)" if bo_year else "BO (End Year Total)"
                bl_val   = _active_margin_data.get(bl_label, 0)
                bo_val   = _active_margin_data.get(bo_label, 0)

                if margin_rows:
                    df_margin          = pd.DataFrame(margin_rows)
                    df_margin["Label"] = df_margin["Production (TL)"].apply(human_tl)

                    _type_suffix = f" [{ph_cost_type}]" if ph_cost_type != "GROSS FEES" else ""
                    fig_margin = px.line(
                        df_margin, x="Period", y="Production (TL)", markers=True,
                        title=f"{sel_label}{_type_suffix} — Monthly Production (TL)",
                        text="Label",
                    )
                    fig_margin.update_traces(
                        textposition="top center",
                        line=dict(color=PRIMARY, width=2),
                        marker=dict(size=8),
                    )
                    fig_margin.update_layout(
                        plot_bgcolor="white",
                        font=dict(family="Segoe UI", color=PRIMARY),
                        yaxis_title="Production (TL)",
                        xaxis_title="",
                        xaxis_tickangle=-45,
                        margin=dict(t=60, b=20),
                    )
                    st.plotly_chart(fig_margin, use_container_width=True, key=f"margin_chart_{sel_proj}_{ph_cost_type}")

                    next_yr_val = _active_margin_data.get("Next Year", None)
                    td_val      = _active_margin_data.get("TD (To Date)", None)

                    metric_cols = st.columns(4)
                    metric_cols[0].metric(f"{bl_year or 'Year'} Production (TL)", human_tl(bl_val))
                    metric_cols[1].metric(f"End {bo_year or 'Year'} Total (TL)",  human_tl(bo_val))
                    if next_yr_val is not None:
                        metric_cols[2].metric("Next Year (TL)", human_tl(next_yr_val))
                    if td_val is not None:
                        metric_cols[3].metric("TD To Date (TL)", human_tl(td_val))

                    df_display = pd.DataFrame([
                        {"Period": row["Period"], "Production (TL)": human_tl(row["Production (TL)"])}
                        for row in margin_rows
                    ])
                    if bl_year:
                        df_display = pd.concat([df_display, pd.DataFrame([
                            {"Period": f"{bl_year} Production (TL)", "Production (TL)": human_tl(bl_val)}
                        ])], ignore_index=True)
                    if bo_year:
                        df_display = pd.concat([df_display, pd.DataFrame([
                            {"Period": f"End {bo_year} Total (TL)", "Production (TL)": human_tl(bo_val)}
                        ])], ignore_index=True)
                    st.dataframe(df_display, hide_index=True, use_container_width=True)
                    st.download_button(
                        "⬇ Export Production to Excel",
                        df_to_excel_bytes(df_display),
                        file_name=f"Production_{proj_info['code']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"prod_export_{sel_proj}",
                    )

            else:  # WIP TL view
                if wip_history:
                    df_wip_hist           = pd.DataFrame(wip_history)
                    latest_wip            = df_wip_hist.iloc[-1]["WIP TL"]
                    prev_wip_val          = df_wip_hist.iloc[-2]["WIP TL"] if len(df_wip_hist) > 1 else None
                    delta_wip             = latest_wip - prev_wip_val if prev_wip_val is not None else None

                    st.metric(
                        f"WIP TL — {sel_label}",
                        human_tl(latest_wip),
                        delta=human_tl(delta_wip) if delta_wip is not None else None,
                    )

                    df_wip_hist["Label"] = df_wip_hist["WIP TL"].apply(human_tl)
                    fig_wip = px.line(
                        df_wip_hist, x="Period", y="WIP TL", markers=True,
                        title=f"{sel_label} — WIP TL History",
                        text="Label",
                    )
                    fig_wip.update_traces(
                        textposition="top center",
                        line=dict(color="#10B981", width=2),
                        marker=dict(size=8),
                    )
                    fig_wip.update_layout(
                        plot_bgcolor="white",
                        font=dict(family="Segoe UI", color=PRIMARY),
                        yaxis_title="WIP TL",
                        xaxis_title="",
                        xaxis_tickangle=-45,
                        margin=dict(t=60, b=20),
                    )
                    st.plotly_chart(fig_wip, use_container_width=True, key=f"wip_chart_{sel_proj}")

                    st.dataframe(
                        df_wip_hist[["Period", "Label"]].rename(columns={"Label": "WIP TL"}),
                        hide_index=True, use_container_width=True,
                    )
                else:
                    st.info("No WIP TL data found for this project in the selected months.")
        else:
            st.info("Please select a project to view history.")
    else:
        st.error("Could not load margin data from 'Ext. Prod.' sheet. Verify the sheet exists and has the correct structure.")


@st.fragment
def _render_scenarios(loaded_data, selected_months, global_bu_view, multi_year):
    _sm_label  = lambda sm: sm if multi_year else sm.split()[0]
    st.markdown("### Scenarios")
    sc_tabs = st.tabs(["BU Comparison", "YTD vs Target", "Pipeline Health", "Top 10 Clients", "BU Mix"])

    _sm_latest   = selected_months[-1]
    _d_latest    = loaded_data.get(_sm_latest)

    # ── S2: BU Comparison ─────────────────────────────────────────────────────
    with sc_tabs[0]:
        st.markdown(f"#### BU NS Comparison — {_sm_latest}")
        sc_rows = []
        for _sm in selected_months:
            _d = loaded_data.get(_sm)
            if not _d:
                continue
            _abbr = _sm.split()[0][:3]
            for _bu in ["ENG", "MC", "T&SI", "NUC"]:
                _bd = _d["bu_ns"][_bu]
                _mi = _cat_index(_bd["cats"], _abbr)
                for _tier in ["Order", "Offer", "Opp"]:
                    sc_rows.append({
                        "Period": _sm_label(_sm),
                        "BU":     _bu,
                        "Tier":   _tier,
                        "Value":  monthly_val(_bd[_tier], _bd["cats"], _mi),
                    })
        if sc_rows:
            _df_sc = pd.DataFrame(sc_rows)
            # Summary: total NS per BU (all tiers)
            _df_bu_sum = (
                _df_sc.groupby(["Period", "BU"])["Value"].sum().reset_index()
            )
            _df_bu_sum["Label"] = _df_bu_sum["Value"].apply(human_k)
            _fig_busum = px.bar(
                _df_bu_sum, x="BU", y="Value", color="BU",
                facet_col="Period" if len(selected_months) > 1 else None,
                text="Label",
                title="Total NS by BU (Order + Offer + Opp)",
                color_discrete_map=BU_COLORS,
            )
            _fig_busum.update_layout(
                plot_bgcolor="white",
                font=dict(family="Segoe UI", color=PRIMARY),
                showlegend=False,
                margin=dict(t=60, b=20),
            )
            _fig_busum.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(_fig_busum, use_container_width=True, key="sc_busum")

            # Tier breakdown per BU — color each (BU, Tier) using bu_tier_colors
            _df_tier = _df_sc.groupby(["BU", "Tier"])["Value"].sum().reset_index()
            _df_tier["BU_Tier"] = _df_tier["BU"] + " " + _df_tier["Tier"]
            _bu_tier_cmap = {}
            for _buc in ["ENG", "MC", "T&SI", "NUC"]:
                _tc = bu_tier_colors(_buc)
                _bu_tier_cmap[f"{_buc} Order"] = _tc[0]
                _bu_tier_cmap[f"{_buc} Offer"] = _tc[1]
                _bu_tier_cmap[f"{_buc} Opp"]   = _tc[2]
            _fig_tier = px.bar(
                _df_tier, x="BU", y="Value", color="BU_Tier",
                barmode="stack",
                title="NS Tier Breakdown by BU",
                color_discrete_map=_bu_tier_cmap,
                category_orders={"BU_Tier": [
                    f"{b} {t}" for b in ["ENG", "MC", "T&SI", "NUC"]
                    for t in ["Order", "Offer", "Opp"]
                ]},
            )
            _fig_tier.update_layout(
                plot_bgcolor="white",
                font=dict(family="Segoe UI", color=PRIMARY),
                showlegend=False,
                margin=dict(t=60, b=20),
            )
            _fig_tier.update_traces(texttemplate="%{y:,.0f}", textposition="inside", insidetextanchor="middle")
            st.plotly_chart(_fig_tier, use_container_width=True, key="sc_butiertotal")

            # EBIT by BU
            st.markdown("#### BU EBIT Comparison")
            _ebit_rows = []
            for _sm in selected_months:
                _d = loaded_data.get(_sm)
                if not _d: continue
                _abbr = _sm.split()[0][:3]
                for _bu in ["ENG", "MC", "T&SI", "NUC"]:
                    _bd = _d["bu_ebit"][_bu]
                    _mi = _cat_index(_bd["cats"], _abbr)
                    _ebit_rows.append({
                        "Period": _sm_label(_sm), "BU": _bu,
                        "EBIT":   monthly_val(_bd["Total"], _bd["cats"], _mi),
                    })
            if _ebit_rows:
                _df_ebit = pd.DataFrame(_ebit_rows)
                _df_ebit["Label"] = _df_ebit["EBIT"].apply(human_k)
                _fig_ebit = px.bar(
                    _df_ebit, x="BU", y="EBIT", color="BU",
                    facet_col="Period" if len(selected_months) > 1 else None,
                    text="Label",
                    title="EBIT by BU",
                    color_discrete_map=BU_COLORS,
                )
                _fig_ebit.update_layout(
                    plot_bgcolor="white",
                    font=dict(family="Segoe UI", color=PRIMARY),
                    showlegend=False, margin=dict(t=60, b=20),
                )
                _fig_ebit.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(_fig_ebit, use_container_width=True, key="sc_buebit")

    # ── S4: YTD vs Target ─────────────────────────────────────────────────────
    with sc_tabs[1]:
        if _d_latest:
            _ns    = _d_latest["ns"]
            _ebit  = _d_latest["ebit"]
            _t_idx = next((i for i, c in enumerate(_ns["cats"]) if "target" in str(c).lower()), None)
            _ytd_abbr = _sm_latest.split()[0][:3]
            _y_idx    = _cat_index(_ns["cats"], _ytd_abbr)

            if _t_idx is not None and _y_idx != -1:
                st.markdown(f"#### YTD vs 2026 Target — {_sm_latest}")
                st.caption("Progress bars show YTD actuals (cumulative) against the 2026 annual target.")

                # NS tiers
                st.markdown("**Net Sales**")
                for _tier, _label in [("Contract", "Contract"), ("WP", "Work Pipeline"), ("WO", "Full Outlook")]:
                    _target = _ns[_tier][_t_idx]
                    _ytd    = _ns[_tier][_y_idx]
                    _pct    = _ytd / _target if _target > 0 else 0
                    _tc1, _tc2, _tc3 = st.columns([3, 1, 1])
                    _tc1.markdown(f"**{_label}**")
                    _tc1.progress(min(1.0, max(0.0, _pct)))
                    _tc2.metric("YTD (kTL)", human_k(_ytd))
                    _tc3.metric("Target (kTL)", human_k(_target), delta=f"{_pct:.0%}")

                st.markdown("---")

                # EBIT tiers
                _et_idx = next((i for i, c in enumerate(_ebit["cats"]) if "target" in str(c).lower()), None)
                _ey_idx = _cat_index(_ebit["cats"], _ytd_abbr)
                if _et_idx is not None and _ey_idx != -1:
                    st.markdown("**EBIT**")
                    for _tier, _label in [
                        ("Contract", "Contract"),
                        ("Contract+WP", "Contract + WP"),
                        ("Contract+WP+WO", "Full Outlook"),
                    ]:
                        _target = _ebit[_tier][_et_idx]
                        _ytd    = _ebit[_tier][_ey_idx]
                        _pct    = _ytd / _target if _target > 0 else 0
                        _tc1, _tc2, _tc3 = st.columns([3, 1, 1])
                        _tc1.markdown(f"**{_label}**")
                        _tc1.progress(min(1.0, max(0.0, _pct)))
                        _tc2.metric("YTD (kTL)", human_k(_ytd))
                        _tc3.metric("Target (kTL)", human_k(_target), delta=f"{_pct:.0%}")
            else:
                st.info("Could not locate the 2026 Target column in the NS data.")
        else:
            st.warning("No data available for the selected month.")

    # ── S7: Pipeline Health ───────────────────────────────────────────────────
    with sc_tabs[2]:
        if _d_latest:
            st.markdown(f"#### NS Pipeline by BU — {_sm_latest}")
            st.caption(
                "Order = contracted revenue. Offer = additional NS if pending offers are won. "
                "Opp = additional NS from identified opportunities."
            )
            _abbr_l = _sm_latest.split()[0][:3]
            _pipe_rows = []
            for _bu in ["ENG", "MC", "T&SI", "NUC"]:
                _bd  = _d_latest["bu_ns"][_bu]
                _mi  = _cat_index(_bd["cats"], _abbr_l)
                if _mi == -1:
                    continue
                _ord = _bd["Order"][_mi] if _mi < len(_bd["Order"]) else 0
                _off = _bd["Offer"][_mi] if _mi < len(_bd["Offer"]) else 0
                _opp = _bd["Opp"][_mi]   if _mi < len(_bd["Opp"])   else 0
                _tot = _ord + _off + _opp
                _pipe_rows.append({
                    "BU":           _bu,
                    "Order":        _ord,
                    "Offer":        _off,
                    "Opp":          _opp,
                    "Total":        _tot,
                    "Order%":       round(_ord / _tot * 100, 1) if _tot > 0 else 0,
                    "Offer Rate%":  round((_ord + _off) / _tot * 100, 1) if _tot > 0 else 0,
                })

            if _pipe_rows:
                _df_pipe  = pd.DataFrame(_pipe_rows)
                _df_melt  = _df_pipe.melt(id_vars=["BU"], value_vars=["Order", "Offer", "Opp"],
                                           var_name="Stage", value_name="Value (kTL)")
                _tier_colors = {"Order": "#0EA5E9", "Offer": "#93C5FD", "Opp": "#CBD5E1"}
                _fig_pipe = px.bar(
                    _df_melt, x="BU", y="Value (kTL)", color="Stage",
                    barmode="stack",
                    title=f"Pipeline Stages by BU — {_sm_latest}",
                    color_discrete_map=_tier_colors,
                    category_orders={"Stage": ["Order", "Offer", "Opp"]},
                )
                _fig_pipe.update_layout(
                    plot_bgcolor="white",
                    font=dict(family="Segoe UI", color=PRIMARY),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                    legend_title_text="",
                    margin=dict(t=80, b=20),
                )
                _fig_pipe.update_traces(texttemplate="%{y:,.0f}", textposition="inside", insidetextanchor="middle")
                st.plotly_chart(_fig_pipe, use_container_width=True, key="sc_pipeline")

                st.markdown("**Conversion Rates**")
                _cr_cols = st.columns(len(_pipe_rows))
                for _ci, _row in enumerate(_pipe_rows):
                    _bu_col = _cr_cols[_ci]
                    _bu_col.markdown(f"**{_row['BU']}**")
                    _bu_col.metric("Total Pipeline", human_k(_row["Total"]))
                    _bu_col.metric("Order (converted)", f"{_row['Order%']}%")
                    _bu_col.metric("Offer + Order", f"{_row['Offer Rate%']}%")

                _df_pipe_disp = _df_pipe.copy()
                _df_pipe_disp["Total"] = _df_pipe_disp["Total"].apply(human_k)
                _df_pipe_disp["Order"] = _df_pipe_disp["Order"].apply(human_k)
                _df_pipe_disp["Offer"] = _df_pipe_disp["Offer"].apply(human_k)
                _df_pipe_disp["Opp"]   = _df_pipe_disp["Opp"].apply(human_k)
                st.dataframe(
                    _df_pipe_disp[["BU", "Order", "Offer", "Opp", "Total", "Order%", "Offer Rate%"]],
                    hide_index=True, use_container_width=True,
                )
        else:
            st.warning("No data available for the selected month.")

    # ── S9: Top 10 Clients ────────────────────────────────────────────────────
    with sc_tabs[3]:
        st.markdown("#### Top 10 Clients")
        _cv = st.radio(
            "Data source", ["Order Intake", "WIP"],
            horizontal=True, key="sc_client_src",
        )
        _formatter = human_k if _cv == "Order Intake" else human_tl
        _unit      = "kTL" if _cv == "Order Intake" else "TL"

        _cli_totals: dict[str, float] = {}
        if _cv == "Order Intake":
            for _sm in selected_months:
                _d = loaded_data.get(_sm)
                if not _d:
                    continue
                for _p in _d["oi_projects"]:
                    _c = _p.get("client") or "Unknown"
                    _cli_totals[_c] = _cli_totals.get(_c, 0) + _p["value"]
        else:
            _d = _d_latest
            if _d:
                for _p in _d["wip_projects"]:
                    _c = _p.get("client") or "Unknown"
                    _cli_totals[_c] = _cli_totals.get(_c, 0) + _p["wip_tl"]

        if _cli_totals:
            _top10 = sorted(_cli_totals.items(), key=lambda x: -x[1])[:10]
            _df_top = pd.DataFrame(_top10, columns=["Client", "Value"])
            _df_top["Label"] = _df_top["Value"].apply(_formatter)

            _t1, _t2 = st.columns(2)
            with _t1:
                _fig_tbar = px.bar(
                    _df_top, x="Value", y="Client", orientation="h",
                    text="Label",
                    title=f"Top 10 Clients by {_cv} ({_unit})",
                    color_discrete_sequence=[PRIMARY],
                    category_orders={"Client": _df_top["Client"].tolist()},
                )
                _fig_tbar.update_layout(
                    plot_bgcolor="white",
                    font=dict(family="Segoe UI", color=PRIMARY),
                    yaxis_title="", xaxis_title=_unit,
                    margin=dict(t=60, r=20),
                    height=max(350, len(_df_top) * 32),
                )
                _fig_tbar.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(_fig_tbar, use_container_width=True, key="sc_topbar")
            with _t2:
                _fig_tpie = px.pie(
                    _df_top, values="Value", names="Client",
                    title=f"Share by Client — {_cv}",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                _fig_tpie.update_traces(textinfo="label+percent", textposition="inside")
                _fig_tpie.update_layout(margin=dict(t=60, b=0))
                st.plotly_chart(_fig_tpie, use_container_width=True, key="sc_toppie")

            st.dataframe(
                _df_top[["Client", "Label"]].rename(columns={"Label": _unit}),
                hide_index=True, use_container_width=True,
            )
            st.download_button(
                "⬇ Export Top 10 to Excel",
                df_to_excel_bytes(_df_top[["Client", "Value", "Label"]].rename(columns={"Label": _unit})),
                file_name=f"Top10_Clients_{_cv.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sc_top10_export",
            )
        else:
            st.info("No client data available for the selected period.")

    # ── S5: BU Mix ────────────────────────────────────────────────────────────
    with sc_tabs[4]:
        st.markdown("#### BU Mix — Include / Exclude Business Units")
        st.caption(
            "Toggle BUs on/off to see the combined NS and EBIT for your selection, "
            "the impact of excluded BUs, and each selection's share of the company total. "
            "Covers all selected months."
        )

        _mix_bu_cols = st.columns(4)
        _bu_active = {
            bu: _mix_bu_cols[i].checkbox(bu, value=True, key=f"sc_mix_{bu}")
            for i, bu in enumerate(["ENG", "MC", "T&SI", "NUC"])
        }
        _sel_bus = [b for b, on in _bu_active.items() if on]
        _all_bus = ["ENG", "MC", "T&SI", "NUC"]

        if not _sel_bus:
            st.warning("Select at least one Business Unit.")
        else:
            _mix_summary  = []
            _mix_chart_rows = []

            for _sm in selected_months:
                _d = loaded_data.get(_sm)
                if not _d:
                    continue
                _abbr = _sm.split()[0][:3]

                _ns_bu, _ebit_bu = {}, {}
                for _bu in _all_bus:
                    _bd = _d["bu_ns"][_bu]
                    _mi = _cat_index(_bd["cats"], _abbr)
                    _ns_bu[_bu] = sum(
                        monthly_val(_bd[t], _bd["cats"], _mi) for t in ["Order", "Offer", "Opp"]
                    ) if _mi != -1 else 0

                    _be = _d["bu_ebit"][_bu]
                    _ei = _cat_index(_be["cats"], _abbr)
                    _ebit_bu[_bu] = monthly_val(_be["Total"], _be["cats"], _ei) if _ei != -1 else 0

                    _mix_chart_rows.append({
                        "Period":  _sm_label(_sm),
                        "BU":      _bu,
                        "NS":      _ns_bu[_bu],
                        "EBIT":    _ebit_bu[_bu],
                    })

                _ns_s  = sum(_ns_bu[b]   for b in _sel_bus)
                _ns_a  = sum(_ns_bu[b]   for b in _all_bus)
                _eb_s  = sum(_ebit_bu[b] for b in _sel_bus)
                _eb_a  = sum(_ebit_bu[b] for b in _all_bus)

                _mix_summary.append({
                    "Period":     _sm_label(_sm),
                    "NS Sel":     _ns_s,
                    "NS Total":   _ns_a,
                    "NS Excl":    _ns_a - _ns_s,
                    "NS Pct":     _ns_s / _ns_a if _ns_a else 0,
                    "EBIT Sel":   _eb_s,
                    "EBIT Total": _eb_a,
                    "EBIT Excl":  _eb_a - _eb_s,
                    "EBIT Pct":   _eb_s / _eb_a if _eb_a else 0,
                })

            # ── Latest-month metric cards ─────────────────────────────────────
            if _mix_summary:
                _lr = _mix_summary[-1]
                _sel_label = ", ".join(_sel_bus) if _sel_bus else "none"
                st.markdown(f"**{_lr['Period']} — Selected: {_sel_label}**")
                _mc1, _mc2, _mc3, _mc4 = st.columns(4)
                _mc1.metric(
                    "NS — Selected BUs",
                    human_k(_lr["NS Sel"]),
                    delta=f"−{human_k(_lr['NS Excl'])} excluded  |  {_lr['NS Pct']:.0%} of total",
                    delta_color="off",
                )
                _mc2.metric(
                    "NS — Excluded Impact",
                    human_k(_lr["NS Excl"]),
                    delta=f"{1 - _lr['NS Pct']:.0%} of total",
                    delta_color="off",
                )
                _mc3.metric(
                    "EBIT — Selected BUs",
                    human_k(_lr["EBIT Sel"]),
                    delta=f"−{human_k(_lr['EBIT Excl'])} excluded  |  {_lr['EBIT Pct']:.0%} of total",
                    delta_color="off",
                )
                _mc4.metric(
                    "EBIT — Excluded Impact",
                    human_k(_lr["EBIT Excl"]),
                    delta=f"{1 - _lr['EBIT Pct']:.0%} of total",
                    delta_color="off",
                )

            # ── Charts ────────────────────────────────────────────────────────
            if _mix_chart_rows:
                _df_mix = pd.DataFrame(_mix_chart_rows)
                _df_mix["NS_label"]   = _df_mix["NS"].apply(human_k)
                _df_mix["EBIT_label"] = _df_mix["EBIT"].apply(human_k)

                _color_map = {
                    bu: (BU_COLORS.get(bu, "#888888") if bu in _sel_bus else "#D1D5DB")
                    for bu in _all_bus
                }

                _fig_mix_ns = px.bar(
                    _df_mix, x="BU", y="NS", color="BU",
                    facet_col="Period" if len(selected_months) > 1 else None,
                    text="NS_label",
                    title="Net Sales by BU — active (coloured) vs excluded (grey)",
                    color_discrete_map=_color_map,
                )
                _fig_mix_ns.update_layout(
                    plot_bgcolor="white",
                    font=dict(family="Segoe UI", color=PRIMARY),
                    showlegend=False,
                    margin=dict(t=60, b=20),
                )
                _fig_mix_ns.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(_fig_mix_ns, use_container_width=True, key="sc_mix_ns")

                _fig_mix_eb = px.bar(
                    _df_mix, x="BU", y="EBIT", color="BU",
                    facet_col="Period" if len(selected_months) > 1 else None,
                    text="EBIT_label",
                    title="EBIT by BU — active (coloured) vs excluded (grey)",
                    color_discrete_map=_color_map,
                )
                _fig_mix_eb.update_layout(
                    plot_bgcolor="white",
                    font=dict(family="Segoe UI", color=PRIMARY),
                    showlegend=False,
                    margin=dict(t=60, b=20),
                )
                _fig_mix_eb.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(_fig_mix_eb, use_container_width=True, key="sc_mix_ebit")

            # ── Summary table (all months) ────────────────────────────────────
            if _mix_summary:
                st.markdown("**Summary across all selected months**")
                _df_mix_tbl = pd.DataFrame([{
                    "Period":            r["Period"],
                    "NS Selected":       human_k(r["NS Sel"]),
                    "NS Total":          human_k(r["NS Total"]),
                    "NS Excl. Impact":   human_k(r["NS Excl"]),
                    "NS Coverage":       f"{r['NS Pct']:.0%}",
                    "EBIT Selected":     human_k(r["EBIT Sel"]),
                    "EBIT Total":        human_k(r["EBIT Total"]),
                    "EBIT Excl. Impact": human_k(r["EBIT Excl"]),
                    "EBIT Coverage":     f"{r['EBIT Pct']:.0%}",
                } for r in _mix_summary])
                st.dataframe(_df_mix_tbl, hide_index=True, use_container_width=True)
                st.download_button(
                    "⬇ Export BU Mix to Excel",
                    df_to_excel_bytes(_df_mix_tbl),
                    file_name="BU_Mix_Scenario.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sc_mix_export",
                )


# ── Tab 5: Project History ────────────────────────────────────────────────────
with tab5:
    _render_project_history(loaded_data, selected_months, global_bu_view, _multi_year)

# ── Tab 6: Scenarios ──────────────────────────────────────────────────────────
with tab6:
    _render_scenarios(loaded_data, selected_months, global_bu_view, _multi_year)
