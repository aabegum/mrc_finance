import colorsys
import io
import os
import re
import smtplib
import subprocess
import sys
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
    _AGGRID_AVAILABLE = True
except ImportError:
    _AGGRID_AVAILABLE = False

from config_loader import ConfigLoader

st.set_page_config(page_title="MRC Business Review", layout="wide", page_icon="📊")

SCRIPT_DIR = Path(__file__).parent


# ── Config ─────────────────────────────────────────────────────────────────────
_cfg      = ConfigLoader(str(SCRIPT_DIR / "config"))
_r_cfg    = _cfg.get("Excel_Mapping.Rows")    or {}
_c_cfg    = _cfg.get("Excel_Mapping.Columns") or {}
_o_cfg    = _cfg.get("Excel_Mapping.Offsets") or {}
_dash_cfg = _cfg.get("Dashboard")             or {}

# ── Dashboard display constants (all sourced from Dashboard section in config) ──
_KTLD               = int(_dash_cfg.get("KTL_Decimal_Places",    2))
_WIP_NEG_THR        = int(_dash_cfg.get("WIP_Neg_Threshold",     -1000))
_CHART_TOP_N        = int(_dash_cfg.get("Chart_Top_N_Projects",  30))
_CLIENT_TOP_N       = int(_dash_cfg.get("Client_Top_N",          10))
_CHART_LABEL_PCT    = float(_dash_cfg.get("Chart_Label_Min_Pct", 0.03))
_EXT_PROD_ROW_LIMIT = int(_dash_cfg.get("ExtProd_Row_Limit",     5000))
_BU_TARGET_IDX      = int(_dash_cfg.get("BU_Target_Col_Index",   1))
_COST_TYPE_ORDER    = _dash_cfg.get("CostType_Order") or [
    "GROSS FEES", "REIMBURSABLES", "SUBCON 1", "SUBCON 2",
    "ASSOCIATE 1", "ENG. COST", "PROJECT EXP", "PROJECT EXP ACCR",
]

# ── Colors (sourced from config.yaml) ─────────────────────────────────────────
PRIMARY  = _cfg.get("THEMES.premium.Colors.Dark.Hex") or "#1E3A8A"
_FONT    = _cfg.get("THEMES.premium.Typography.Font_Family") or "Segoe UI"
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
    "GM_STATUS":              _c_cfg.get("GM_STATUS", 6),
    "GM_ITEM":                _c_cfg.get("GM_ITEM",   7),
    "GM_PROJ":                _c_cfg.get("GM_PROJ",   4),
    "GM_CLIENT":              _c_cfg.get("GM_CLIENT", 2),
    "GM_TOTAL":               _c_cfg.get("GM_TOTAL",  25),
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


def _tr_num(s: str) -> str:
    """Swap US separators to Turkish: 1,234.56 → 1.234,56"""
    return s.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def fmt_ktl(v, d=None):
    """Format a kTL value with d decimal places and thousand separators (Turkish format)."""
    _d = _KTLD if d is None else d
    try:
        return _tr_num(f"{float(v):,.{_d}f}")
    except Exception:
        return f"0,{'0' * _KTLD}"


def fmt_tl_as_ktl(v, d=None):
    """Convert raw TL to kTL and format with d decimal places (Turkish format)."""
    _d = _KTLD if d is None else d
    try:
        return _tr_num(f"{float(v) / 1000:,.{_d}f}")
    except Exception:
        return f"0,{'0' * _KTLD}"


def _mfmt(v: float) -> str:
    """Abbreviated format for metric-card values — avoids truncation in narrow columns."""
    return human_k(v)


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


def _find_eoy_idx(cats):
    """Return index of the end-of-year column (Dec or current-year label). Returns -1 if not found."""
    for i, c in enumerate(cats):
        if str(c).strip().lower()[:3] == "dec":
            return i
    yr = str(datetime.now().year)
    for i, c in enumerate(cats):
        if str(c).strip() == yr:
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

    # ── Row Offset Detection ── scan from sheet bottom upward ──────────────────
    # Probing a fixed narrow range around the expected row fails for files with
    # fewer projects (e.g. 2025 files) whose summary section sits higher up.
    # Scanning upward from the sheet bottom is reliable for any file version.
    _oi_offset       = 0
    _oi_offset_found = False
    expected_cat_row = R["Categories_OI_kTL"]
    _oi_col_start    = _c_cfg.get("OI_kTL_Start", 4)
    _scan_start      = min(len(df_oi) - 1, expected_cat_row + 50)  # start at most 50 past expected
    for scan_row in range(_scan_start, max(0, _scan_start - 300), -1):
        _c4 = str(df_oi.iloc[scan_row, _oi_col_start]).strip()
        if _oi_col_start + 2 < len(df_oi.columns):
            _c6      = str(df_oi.iloc[scan_row, _oi_col_start + 2]).strip()
            _c4_base = _c4.split(".")[0].strip()
            if _c4_base.isdigit() and len(_c4_base) == 4:
                if "jan" in _c6.lower() or "feb" in _c6.lower():
                    _oi_offset       = scan_row - expected_cat_row
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

    # WIP projects — Nuclear excluded by default but collected separately for optional display
    wip_projects      = []
    wip_projects_nuc  = []
    wip_gross_total     = 0.0  # Grand total: ALL non-NUC WIP (pos+neg) before 1M filter
    wip_gross_total_nuc = 0.0  # Same for NUC
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
        wip_tl = safe_float(row.iloc[C["WIP_Col_WIP_TL"]])
        # Accumulate grand total BEFORE threshold filter (non-zero entries only)
        if wip_tl != 0:
            if bu == "NUC":
                wip_gross_total_nuc += wip_tl
            else:
                wip_gross_total += wip_tl
        # Mirror reporting code: skip zeros and near-zero negatives.
        # Include: significant positives (>= WIP_MIN_TL) and significant negatives (< WIP_NEG_THR).
        if wip_tl > _WIP_NEG_THR and wip_tl < WIP_MIN_TL:
            continue
        _entry = {
            "name":          str(row.iloc[C["WIP_Col_Name"]]).strip()          if pd.notna(row.iloc[C["WIP_Col_Name"]])          else "",
            "client":        client,
            "bu":            bu,
            "orig_currency": str(row.iloc[C["WIP_Col_OrigCurrency"]]).strip()  if pd.notna(row.iloc[C["WIP_Col_OrigCurrency"]])  else "",
            "inv_oc":        safe_float(row.iloc[C["WIP_Col_TotalInvoiceOC"]]),
            "prod_oc":       safe_float(row.iloc[C["WIP_Col_TotalProdOC"]]),
            "wip_tl":        wip_tl,
        }
        if bu == "NUC":
            wip_projects_nuc.append(_entry)
        else:
            wip_projects.append(_entry)

    def _sort_wip(lst):
        _pos = sorted([p for p in lst if p["wip_tl"] >= 0], key=lambda x: x["wip_tl"], reverse=True)
        _neg = sorted([p for p in lst if p["wip_tl"] < 0],  key=lambda x: x["wip_tl"])
        return _pos + _neg

    wip_projects     = _sort_wip(wip_projects)
    wip_projects_nuc = _sort_wip(wip_projects_nuc)

    # ── ABNS from Gross Margin sheet ──────────────────────────────────────────
    abns_projects: list = []
    abns_total = 0.0
    if "Gross Margin" in xl.sheet_names:
        try:
            df_gm = pd.read_excel(xl, sheet_name="Gross Margin", header=None)
            for ri in range(4, df_gm.shape[0]):
                _row    = df_gm.iloc[ri]
                _status = str(_row.iloc[C["GM_STATUS"]]).strip().upper() if pd.notna(_row.iloc[C["GM_STATUS"]]) else ""
                _item   = str(_row.iloc[C["GM_ITEM"]]).strip().upper()   if pd.notna(_row.iloc[C["GM_ITEM"]])   else ""
                if _status != "ABNS" or _item != "GR":
                    continue
                _val = safe_float(_row.iloc[C["GM_TOTAL"]])
                if _val <= 0:
                    continue
                _proj   = str(_row.iloc[C["GM_PROJ"]]).strip()   if pd.notna(_row.iloc[C["GM_PROJ"]])   else ""
                _client = str(_row.iloc[C["GM_CLIENT"]]).strip() if pd.notna(_row.iloc[C["GM_CLIENT"]]) else ""
                if not _proj or _proj.lower() in ("nan", "0", ""):
                    continue
                abns_projects.append({"project": _proj, "client": _client, "value": _val})
            abns_projects.sort(key=lambda x: x["value"], reverse=True)
            abns_total = sum(p["value"] for p in abns_projects)
        except Exception:
            pass

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
        p_row_key = f"BU_EBIT_P_{b_key.replace('&', '')}"
        p_row_def = {"ENG": 108, "MC": 107, "T&SI": 109, "NUC": 110}[b_key]
        bu_ebit[b_key] = {
            "cats":  cats_bu_ebit,
            "Total": row_vals(df, _r_cfg.get(b_row, 100) - 1, c_ebit_s, c_ebit_e),
            "Pct":   row_vals(df, _r_cfg.get(p_row_key, p_row_def), c_ebit_s, c_ebit_e),
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
        "Pct":            row_vals(df, _r_cfg.get("EBIT_P_kTL", 42), 1, 16),
    }

    # ── EBIT % / Net Fees % rows — read from "EBIT Calc." sheet rows 67/68/69 ──
    # The EBIT Calc. sheet has date headers in row 2 (pandas row 1), months start
    # at column B (pandas col 1). Rows 67/68/69 (Excel) = 66/67/68 (0-indexed).
    _ec_sheet = "EBIT Calc."
    ebit_pct_out: dict = {"cats": [], "EBIT_Pct": [], "EBIT_Pct_Budget": [], "NetFees_Pct": []}
    if _ec_sheet in xl.sheet_names:
        try:
            df_ec    = pd.read_excel(xl, sheet_name=_ec_sheet, header=None)
            _ec_cats = []
            for _ec_c in range(1, 14):  # columns B-N (12 months + possible total)
                _cv = df_ec.iloc[1, _ec_c] if _ec_c < df_ec.shape[1] else None
                if _cv is None or (isinstance(_cv, float) and pd.isna(_cv)):
                    _ec_cats.append("")
                elif hasattr(_cv, "strftime"):
                    _ec_cats.append(_cv.strftime("%b"))
                else:
                    _s = str(_cv).strip()
                    _matched = next((m.capitalize() for m in _MONTH_ABBREVS if m in _s.lower()), None)
                    _ec_cats.append(_matched if _matched else clean_label(_cv))
            _ec_r_act = _r_cfg.get("EBIT_Pct_Actual",  66)
            _ec_r_bud = _r_cfg.get("EBIT_Pct_Budget",  67)
            _ec_r_nf  = _r_cfg.get("Net_Fees_Pct",     68)
            ebit_pct_out = {
                "cats":            _ec_cats,
                "EBIT_Pct":        row_vals(df_ec, _ec_r_act, 1, 14),
                "EBIT_Pct_Budget": row_vals(df_ec, _ec_r_bud, 1, 14),
                "NetFees_Pct":     row_vals(df_ec, _ec_r_nf,  1, 14),
            }
        except Exception:
            pass

    return {
        "ns":       ns_out,
        "ebit":     ebit_out,
        "ebit_pct": ebit_pct_out,
        "oi": {
            "cats": cats_oi,
            "ENG":  [safe_float(df_oi.iloc[R["Slide6_ENG"] + _oi_offset, c]) for c in range(4, 18)],
            "MC":   [safe_float(df_oi.iloc[R["Slide6_MC"]  + _oi_offset, c]) for c in range(4, 18)],
            "T&SI": [safe_float(df_oi.iloc[R["Slide6_TSI"] + _oi_offset, c]) for c in range(4, 18)],
            "NUC":  [safe_float(df_oi.iloc[R["Slide6_NUC"] + _oi_offset, c]) for c in range(4, 18)],
        },
        "oi_projects":      oi_projects,
        "wip_projects":     wip_projects,
        "wip_projects_nuc":   wip_projects_nuc,
        "wip_gross_total":    wip_gross_total,
        "wip_gross_total_nuc": wip_gross_total_nuc,
        "abns_projects":    abns_projects,
        "abns_total":     abns_total,
        "bu_ns":          bu_ns,
        "bu_ebit":        bu_ebit,
    }


@st.cache_data(ttl=3600)
def fetch_all_data(months_tuple):
    res = {}
    for sm in months_tuple:
        fp = EXCEL_FILES.get(sm)
        if fp and fp.exists():
            res[sm] = load_month(str(fp))
    return res


@st.cache_resource(ttl=3600)
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

        _EXT_ROW_LIMIT = _EXT_PROD_ROW_LIMIT
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


@st.cache_resource(ttl=3600)
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
    cats = data["cats"][:n]

    # Per-bar positive totals (used for labelling and totals-above-bar)
    pos_totals = [sum(max(data[k][i], 0) for k in keys) for i in range(n)]

    colors = color_palette or (
        [BU_COLORS.get(k, "#0EA5E9") for k in keys]
        if any(k in BU_COLORS for k in keys)
        else ["#0EA5E9", "#10B981", "#F59E0B", "#E11D48"]
    )

    fig = go.Figure()
    for ki, key in enumerate(keys):
        vals  = [data[key][i] for i in range(n)]
        texts = []
        for i, v in enumerate(vals):
            t = pos_totals[i]
            # Label shown only when segment physically has room — constraintext="inside"
            # handles the final pixel-level check; here we only drop zero/negative and
            # single-tier dominant bars (which are covered by the total label above).
            if t <= 0 or v <= 0 or v / t >= 0.95:
                texts.append("")
            else:
                texts.append(human_k(v))
        fig.add_trace(go.Bar(
            x=cats, y=vals, name=key,
            marker_color=colors[ki % len(colors)],
            text=texts,
            textposition="inside",
            insidetextanchor="middle",
            constraintext="inside",   # Plotly hides if text doesn't physically fit
            textfont=dict(size=13, color="white"),  # larger font = stricter physical fit check
            hovertemplate=f"<b>{key}</b>: %{{y:,.0f}}<extra></extra>",
        ))

    # Total label above each bar via Scatter text trace (avoids add_annotation
    # which breaks categorical axis when category names look numeric like "2024")
    fig.add_trace(go.Scatter(
        x=cats,
        y=pos_totals,
        mode="text",
        text=[human_k(t) if t > 0 else "" for t in pos_totals],
        textposition="top center",
        textfont=dict(size=10, color=PRIMARY, family=_FONT),
        showlegend=False,
        hoverinfo="skip",
    ))

    fig.update_layout(
        barmode="stack",
        plot_bgcolor="white",
        font=dict(family=_FONT, color=PRIMARY),
        title=title or "",
        xaxis=dict(type="category", categoryorder="array", categoryarray=cats, tickangle=-45),
        yaxis=dict(rangemode="tozero"),
        margin=dict(t=80, b=20, r=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        legend_title_text="",
    )
    return fig


def chart_grouped(data, title, keys, color_palette=None, percents=None):
    n    = min(len(data["cats"]), min(len(data[k]) for k in keys))
    cats = data["cats"][:n]
    colors = color_palette or ["#0EA5E9", "#10B981", "#F59E0B"]

    fig = go.Figure()
    for ki, key in enumerate(keys):
        vals  = [data[key][i] for i in range(n)]
        # Labels above bars — outside position avoids the rotated-text problem in grouped charts
        texts = [human_k(v) if v != 0 else "" for v in vals]
        fig.add_trace(go.Bar(
            x=cats, y=vals, name=key,
            marker_color=colors[ki % len(colors)],
            text=texts,
            textposition="outside",
            cliponaxis=False,
            textfont=dict(size=10, color=PRIMARY),
            hovertemplate=f"<b>{key}</b>: %{{y:,.0f}}<extra></extra>",
        ))

    # Compute y-axis max with 18% headroom so "outside" bar labels never clip
    all_vals = [v for k in keys for v in [data[k][i] for i in range(n)]]
    _ymax = max(all_vals) * 1.18 if all_vals and max(all_vals) > 0 else None

    fig.update_layout(
        barmode="group",
        plot_bgcolor="white",
        font=dict(family=_FONT, color=PRIMARY),
        title=title or "",
        xaxis=dict(type="category", categoryorder="array", categoryarray=cats, tickangle=-45),
        yaxis=dict(rangemode="tozero", range=[0, _ymax] if _ymax else None),
        margin=dict(t=80, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        legend_title_text="",
    )
    return fig


def chart_projects(projects, title):
    if not projects:
        return None
    df          = pd.DataFrame(projects[:_CHART_TOP_N])
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
        font=dict(family=_FONT, color=PRIMARY),
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
    df          = pd.DataFrame(wip_list[:_CHART_TOP_N])
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
        font=dict(family=_FONT, color=PRIMARY),
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

# ── Context-aware metric formatter ────────────────────────────────────────────
# In comparison mode (2+ months side-by-side) use abbreviated human_k to avoid
# truncation in narrow columns. In single-month mode use fmt_ktl (full precision
# with the kTL Decimal Places slider applied).
if len(selected_months) > 1:
    def _mfmt(v: float) -> str:
        return human_k(v)
else:
    def _mfmt(v: float) -> str:
        return fmt_ktl(v)

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

    st.markdown("---")
    with st.expander("⚙️ Display Settings"):
        st.caption("Overrides config defaults for this session.")
        st.slider("kTL Decimal Places", 0, 4, _KTLD,         key="ss_ktld")
        st.slider("Chart Top-N Projects", 5, 50, _CHART_TOP_N, key="ss_chart_top_n")
        st.slider("Top-N Clients", 5, 20, _CLIENT_TOP_N,     key="ss_client_top_n")
        st.slider(
            "Chart Label Min % (of bar)", 1, 10,
            max(1, int(_CHART_LABEL_PCT * 100)), key="ss_label_pct_int",
        )

    st.markdown("---")
    with st.expander("📧 Email Settings"):
        st.caption("SMTP credentials — session only, not stored on disk.")
        st.text_input("SMTP Host", value="smtp.gmail.com",  key="em_host")
        st.number_input("Port",    value=587, step=1, format="%d", key="em_port")
        st.text_input("Username (sender)", value="", key="em_user")
        st.text_input("Password / App-key", value="", type="password", key="em_pass")
        st.text_input("Recipients (comma-separated)", value="", key="em_to")
        st.caption("💡 For Gmail, use an App Password (not your login password).")

    st.markdown("---")
    st.markdown("### Generate + Email Report")
    st.caption("Generate PPTX then email it:")
    email_btn = st.button("📧 Generate + Email PPTX", use_container_width=True)

# ── Email helper ───────────────────────────────────────────────────────────────
def _send_pptx_email(pptx_path: str, month_label: str) -> str:
    """Send generated PPTX as email attachment. Returns '' on success, error string on failure."""
    host   = st.session_state.get("em_host", "smtp.gmail.com")
    port   = int(st.session_state.get("em_port", 587))
    user   = st.session_state.get("em_user", "")
    pwd    = st.session_state.get("em_pass", "")
    to_raw = st.session_state.get("em_to", "")
    if not all([host, user, pwd, to_raw]):
        return "Email settings incomplete. Fill in SMTP credentials in the sidebar."
    recipients = [r.strip() for r in to_raw.split(",") if r.strip()]
    if not recipients:
        return "No valid recipient addresses found."
    msg            = MIMEMultipart()
    msg["From"]    = user
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = f"MRC Business Review — {month_label}"
    msg.attach(MIMEText(
        f"Please find the MRC Business Review report for {month_label} attached.", "plain"
    ))
    fname = Path(pptx_path).name
    with open(pptx_path, "rb") as fh:
        part = MIMEApplication(fh.read(), Name=fname)
    part["Content-Disposition"] = f'attachment; filename="{fname}"'
    msg.attach(part)
    try:
        with smtplib.SMTP(host, port) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.login(user, pwd)
            smtp.sendmail(user, recipients, msg.as_string())
        return ""
    except Exception as exc:
        return str(exc)


# ── Run report ─────────────────────────────────────────────────────────────────
def _run_generator(sel_file) -> tuple[int, str, str, str | None]:
    """Run PPTX generator subprocess; return (returncode, stdout, stderr, pptx_path|None)."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT_DIR / "generate_full_mbr_stacked.py"), str(sel_file)],
        capture_output=True, text=True, cwd=str(SCRIPT_DIR),
    )
    pptx_path = None
    if result.returncode == 0:
        done_line = next(
            (l.strip() for l in result.stdout.splitlines() if l.strip().startswith("[DONE]")),
            None,
        )
        if done_line:
            # [DONE] path/to/file.pptx  OR  [DONE] Saved to path/to/file.pptx
            _tail = done_line.removeprefix("[DONE]").strip()
            for _tok in _tail.split():
                if _tok.endswith(".pptx") and Path(_tok).exists():
                    pptx_path = _tok
                    break
            if not pptx_path:
                # try scanning SCRIPT_DIR for recently modified .pptx
                _candidates = sorted(
                    SCRIPT_DIR.glob("*.pptx"),
                    key=lambda p: p.stat().st_mtime, reverse=True,
                )
                if _candidates:
                    pptx_path = str(_candidates[0])
    return result.returncode, result.stdout, result.stderr, pptx_path


if run_btn and sel_file:
    st.markdown("---")
    with st.spinner(f"Generating PPTX for {gen_month}…"):
        _rc, _stdout, _stderr, _pptx_path = _run_generator(sel_file)
    if _rc == 0:
        st.success(f"Report generated successfully for {gen_month}!")
        done_line = next(
            (l.strip() for l in _stdout.splitlines() if l.strip().startswith("[DONE]")),
            None,
        )
        if done_line:
            st.info(done_line.removeprefix("[DONE]").strip())
        if _stdout:
            with st.expander("Show generation log", expanded=False):
                st.code(_stdout, language="")
    else:
        st.error("Generator encountered an error — see details below.")
        if _stdout:
            st.code(_stdout, language="")
        if _stderr:
            with st.expander("Error details", expanded=True):
                st.code(_stderr, language="")
    st.markdown("---")

if email_btn and sel_file:
    st.markdown("---")
    with st.spinner(f"Generating PPTX for {gen_month}…"):
        _rc, _stdout, _stderr, _pptx_path = _run_generator(sel_file)
    if _rc != 0:
        st.error("PPTX generation failed — cannot send email.")
        if _stderr:
            with st.expander("Error details", expanded=True):
                st.code(_stderr, language="")
    elif not _pptx_path:
        st.warning("PPTX generated but output file path could not be determined. Check the log.")
        if _stdout:
            with st.expander("Generation log", expanded=True):
                st.code(_stdout, language="")
    else:
        st.success(f"PPTX generated: {Path(_pptx_path).name}")
        with st.spinner("Sending email…"):
            _err = _send_pptx_email(_pptx_path, gen_month)
        if _err:
            st.error(f"Email failed: {_err}")
        else:
            _to_display = st.session_state.get("em_to", "")
            st.success(f"Email sent to {_to_display}")
    st.markdown("---")


# ── Apply sidebar display-setting overrides ────────────────────────────────────
_KTLD            = int(st.session_state.get("ss_ktld",         _KTLD))
_CHART_TOP_N     = int(st.session_state.get("ss_chart_top_n",  _CHART_TOP_N))
_CLIENT_TOP_N    = int(st.session_state.get("ss_client_top_n", _CLIENT_TOP_N))
_CHART_LABEL_PCT = st.session_state.get("ss_label_pct_int",    int(_CHART_LABEL_PCT * 100)) / 100


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
    eoy_idx    = _find_eoy_idx(ns["cats"])

    # Metric cards show per-month delta; charts use the full cumulative array.
    # Round each value to _KTLD before summing so Total matches the displayed sum exactly.
    totals = {k: round(monthly_val(ns[k], ns["cats"], ns_idx), _KTLD) for k in k_list}
    totals["Total"] = round(sum(totals.values()), _KTLD)

    prev_totals = {}
    if p_ns and prev_sm:
        p_abbr = prev_sm.split()[0][:3]
        p_idx  = _cat_index(p_ns["cats"], p_abbr)
        if p_idx != -1:
            prev_totals = {k: round(monthly_val(p_ns[k], p_ns["cats"], p_idx), _KTLD) for k in k_list}
            prev_totals["Total"] = round(sum(prev_totals.values()), _KTLD)

    all_metric_keys = k_list + ["Total"]

    st.caption(f"Monthly — {sm} (kTL)")
    mcols = st.columns(len(all_metric_keys))
    for idx_m, k in enumerate(all_metric_keys):
        val   = totals[k]
        delta = val - prev_totals.get(k, 0) if prev_totals else None
        mcols[idx_m].metric(k, _mfmt(val), delta=_mfmt(delta) if delta is not None else None)

    if ns_idx != -1:
        st.caption("YTD Cumulative (kTL)")
        ytd_cols  = st.columns(len(all_metric_keys))
        ytd_total = 0.0
        for idx_m, k in enumerate(k_list):
            ytd_val = ns[k][ns_idx] if ns_idx < len(ns[k]) else 0.0
            ytd_total += ytd_val
            ytd_cols[idx_m].metric(f"YTD {k}", _mfmt(ytd_val))
        ytd_cols[len(k_list)].metric("YTD Total", _mfmt(ytd_total))

    if eoy_idx != -1:
        st.caption("End of Year (kTL)")
        eoy_cols  = st.columns(len(all_metric_keys))
        eoy_total = 0.0
        for idx_m, k in enumerate(k_list):
            eoy_val = ns[k][eoy_idx] if eoy_idx < len(ns[k]) else 0.0
            eoy_total += eoy_val
            eoy_cols[idx_m].metric(f"EOY {k}", _mfmt(eoy_val))
        eoy_cols[len(k_list)].metric("EOY Total", _mfmt(eoy_total))

    # ABNS (Awarded But, Not Signed) — shown in Net Sales tab per business request
    if is_total:
        _abns_total = data.get("abns_total", 0.0)
        _abns_projs = data.get("abns_projects", [])
        if _abns_total > 0:
            st.caption("ABNS — Awarded But, Not Signed (kTL)")
            _ac1, _ac2 = st.columns([1, 3])
            _ac1.metric("Total ABNS", fmt_ktl(_abns_total))
            if _abns_projs:
                with _ac2.expander(f"ABNS Projects ({len(_abns_projs)})", expanded=False):
                    _df_abns = pd.DataFrame(_abns_projs)
                    _df_abns["Value (kTL)"] = _df_abns["value"].apply(fmt_ktl)
                    st.dataframe(
                        _df_abns[["project", "client", "Value (kTL)"]],
                        hide_index=True, use_container_width=True,
                    )

    palette = None if is_total else bu_tier_colors(bu_name)
    st.plotly_chart(
        chart_stacked(ns, "", k_list, palette),
        use_container_width=True, key=f"ns_{bu_name}_{sm}",
    )

    # ── Chart data table ───────────────────────────────────────────────────────
    with st.expander("📊 Chart Data (YTD Cumulative, kTL)", expanded=False):
        _n_ns  = min(len(ns["cats"]), min(len(ns[k]) for k in k_list))
        _deltas = {k: to_monthly_deltas(ns[k], ns["cats"]) for k in k_list}
        _tbl_rows = []
        for _i in range(_n_ns):
            _row = {"Period": ns["cats"][_i]}
            for k in k_list:
                _row[f"{k} (Monthly)"] = fmt_ktl(_deltas[k][_i])
                _row[f"{k} (YTD)"]     = fmt_ktl(ns[k][_i])
            _tbl_rows.append(_row)
        st.dataframe(pd.DataFrame(_tbl_rows), hide_index=True, use_container_width=True)
        st.download_button(
            "⬇ Export NS Data",
            df_to_excel_bytes(pd.DataFrame(_tbl_rows)),
            file_name=f"NS_Data_{sm.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"ns_data_export_{bu_name}_{sm}",
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
    eoy_idx       = _find_eoy_idx(ebit["cats"])
    prev_ebit_idx = -1
    if p_ebit and prev_sm:
        prev_abbr     = prev_sm.split()[0][:3]
        prev_ebit_idx = _cat_index(p_ebit["cats"], prev_abbr)

    if not is_total:
        # BU view — compact 4-column row: Monthly | YTD | EOY | EBIT %
        _mon_val = round(monthly_val(ebit["Total"], ebit["cats"], ebit_idx), _KTLD)
        _p_val   = round(monthly_val(p_ebit["Total"], p_ebit["cats"], prev_ebit_idx), _KTLD) if p_ebit and prev_ebit_idx != -1 else None
        _delta   = round(_mon_val - _p_val, _KTLD) if _p_val is not None else None
        _ytd_val = ebit["Total"][ebit_idx] if ebit_idx != -1 and ebit_idx < len(ebit["Total"]) else 0.0
        _eoy_val = ebit["Total"][eoy_idx]  if eoy_idx  != -1 and eoy_idx  < len(ebit["Total"]) else 0.0
        _pct_raw = safe_float(ebit["Pct"][ebit_idx]) if ebit_idx != -1 and ebit_idx < len(ebit.get("Pct", [])) else None

        st.caption(f"Monthly / YTD / EOY — {sm} (kTL)")
        _bc1, _bc2, _bc3, _bc4 = st.columns(4)
        _bc1.metric("Monthly",  _mfmt(_mon_val), delta=_mfmt(_delta) if _delta is not None else None)
        _bc2.metric("YTD",      _mfmt(_ytd_val))
        _bc3.metric("EOY",      _mfmt(_eoy_val))
        if _pct_raw is not None:
            _bc4.metric("EBIT %", f"{_pct_raw * 100:.1f}%")
    else:
        # Company Total — three-row layout
        st.caption(f"Monthly — {sm} (kTL)")
        mcols = st.columns(len(k_list))
        for idx_m, k in enumerate(k_list):
            val   = round(monthly_val(ebit[k], ebit["cats"], ebit_idx), _KTLD)
            p_val = round(monthly_val(p_ebit[k], p_ebit["cats"], prev_ebit_idx), _KTLD) if p_ebit and prev_ebit_idx != -1 else None
            delta = round(val - p_val, _KTLD) if p_val is not None else None
            mcols[idx_m].metric(k, _mfmt(val), delta=_mfmt(delta) if delta is not None else None)

        if ebit_idx != -1:
            st.caption("YTD Cumulative (kTL)")
            ytd_cols = st.columns(len(k_list))
            for idx_m, k in enumerate(k_list):
                ytd_val = ebit[k][ebit_idx] if ebit_idx < len(ebit[k]) else 0.0
                ytd_cols[idx_m].metric(f"YTD {k}", _mfmt(ytd_val))

        if eoy_idx != -1:
            st.caption("End of Year (kTL)")
            eoy_cols = st.columns(len(k_list))
            for idx_m, k in enumerate(k_list):
                eoy_val = ebit[k][eoy_idx] if eoy_idx < len(ebit[k]) else 0.0
                eoy_cols[idx_m].metric(f"EOY {k}", _mfmt(eoy_val))

    ebit_palette = None if is_total else [BU_COLORS.get(bu_name, "#0EA5E9")]
    st.plotly_chart(
        chart_grouped(ebit, "", k_list, ebit_palette, percents=ebit.get("Pct")),
        use_container_width=True, key=f"ebit_{bu_name}_{sm}",
    )

    # BU view — EBIT % trend chart (mirrors Company Total margin rates display)
    if not is_total:
        _pct_data = ebit.get("Pct", [])
        _pct_month_idxs = [
            i for i, c in enumerate(ebit["cats"][:len(ebit["Total"])])
            if str(c).strip().lower()[:3] in _MONTH_ABBREVS
        ]
        if _pct_month_idxs and len(_pct_data) > 0:
            _pct_cats = [ebit["cats"][i] for i in _pct_month_idxs if i < len(_pct_data)]
            _pct_vals = [safe_float(_pct_data[i]) * 100 for i in _pct_month_idxs if i < len(_pct_data)]
            if _pct_cats and any(v != 0 for v in _pct_vals):
                fig_bu_pct = go.Figure()
                fig_bu_pct.add_trace(go.Scatter(
                    x=_pct_cats, y=_pct_vals,
                    name="EBIT % (YTD)",
                    mode="lines+markers+text",
                    text=[f"{v:.1f}%" for v in _pct_vals],
                    textposition="top center",
                    line=dict(color=BU_COLORS.get(bu_name, PRIMARY), width=2),
                    marker=dict(size=7),
                ))
                fig_bu_pct.update_layout(
                    plot_bgcolor="white",
                    font=dict(family=_FONT, color=PRIMARY),
                    yaxis=dict(title="%", ticksuffix="%"),
                    xaxis=dict(type="category"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02),
                    margin=dict(t=40, b=20),
                    height=260,
                )
                st.caption(f"EBIT % (YTD) — {bu_name}")
                st.plotly_chart(fig_bu_pct, use_container_width=True, key=f"ebit_bu_pct_{bu_name}_{sm}")

    # ── Chart data table ───────────────────────────────────────────────────────
    with st.expander("📊 Chart Data (YTD Cumulative, kTL)", expanded=False):
        _n_eb  = min(len(ebit["cats"]), min(len(ebit[k]) for k in k_list))
        _ed    = {k: to_monthly_deltas(ebit[k], ebit["cats"]) for k in k_list}
        _etbl  = []
        for _i in range(_n_eb):
            _row = {"Period": ebit["cats"][_i]}
            for k in k_list:
                _row[f"{k} (Monthly)"] = fmt_ktl(_ed[k][_i])
                _row[f"{k} (YTD)"]     = fmt_ktl(ebit[k][_i])
            _etbl.append(_row)
        st.dataframe(pd.DataFrame(_etbl), hide_index=True, use_container_width=True)
        st.download_button(
            "⬇ Export EBIT Data",
            df_to_excel_bytes(pd.DataFrame(_etbl)),
            file_name=f"EBIT_Data_{sm.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"ebit_data_export_{bu_name}_{sm}",
        )

    # ── EBIT % margin rates — computed from loaded EBIT / NS data ─────────────
    # Rows 66-68 in the Excel can be empty; computing from data we already load
    # is reliable and avoids depending on potentially-zeroed % rows.
    if is_total:
        _ns   = data["ns"]
        _eb   = ebit  # already resolved above (ebit = data["ebit"] for is_total)

        _epct    = data.get("ebit_pct")
        _ep_i    = _cat_index(_epct["cats"], month_abbr) if _epct else -1

        # Current-month YTD cumulative margin (from Excel row 66)
        _act_raw = safe_float(_epct["EBIT_Pct"][_ep_i]) if (_epct and _ep_i != -1 and _ep_i < len(_epct["EBIT_Pct"])) else 0
        _act_pct = _act_raw * 100

        # Budget margin: from Excel row 67
        _bud_raw = safe_float(_epct["EBIT_Pct_Budget"][_ep_i]) if (_epct and _ep_i != -1 and _ep_i < len(_epct["EBIT_Pct_Budget"])) else 0
        _bud_pct = _bud_raw * 100

        # Net Fees %: try the stored row first (it may be non-zero); fall back to N/A
        _nf_raw  = safe_float(_epct["NetFees_Pct"][_ep_i]) if (_epct and _ep_i != -1 and _ep_i < len(_epct["NetFees_Pct"])) else 0
        _nf_pct  = _nf_raw * 100 if _nf_raw != 0 else None  # None = no data

        st.caption(f"EBIT Margin Rates — {sm} (YTD Cumulative)")
        _pc1, _pc2, _pc3 = st.columns(3)
        _pc1.metric(
            "EBIT % (Actual, YTD)",
            f"{_act_pct:.1f}%",
            delta=f"{_act_pct - _bud_pct:+.1f} pp vs Budget",
        )
        _pc2.metric("EBIT % (Budget target)", f"{_bud_pct:.1f}%")
        _pc3.metric("Net Fees %", f"{_nf_pct:.1f}%" if _nf_pct is not None else "N/A")

        # Monthly trend: YTD cumulative EBIT % for each month column
        _ep_month_idxs = [
            i for i, c in enumerate(_eb["cats"][:len(_eb["Contract"])])
            if str(c).strip().lower()[:3] in _MONTH_ABBREVS
        ]
        if _ep_month_idxs:
            _ep_cats   = [_eb["cats"][i] for i in _ep_month_idxs]
            _ep_act    = []
            _ep_bud_ln = []
            for _mi in _ep_month_idxs:
                _ep_cat_str = _eb["cats"][_mi]
                _epct_i = _cat_index(_epct["cats"], _ep_cat_str) if _epct else -1
                
                _act_m_raw = safe_float(_epct["EBIT_Pct"][_epct_i]) if (_epct and _epct_i != -1 and _epct_i < len(_epct["EBIT_Pct"])) else 0
                _ep_act.append(_act_m_raw * 100)
                
                _bud_m_raw = safe_float(_epct["EBIT_Pct_Budget"][_epct_i]) if (_epct and _epct_i != -1 and _epct_i < len(_epct["EBIT_Pct_Budget"])) else 0
                _ep_bud_ln.append(_bud_m_raw * 100)

            fig_pct = go.Figure()
            fig_pct.add_trace(go.Scatter(
                x=_ep_cats, y=_ep_act, name="EBIT % (YTD Actual)",
                mode="lines+markers+text",
                text=[f"{v:.1f}%" for v in _ep_act],
                textposition="top center",
                line=dict(color=PRIMARY, width=2),
                marker=dict(size=7),
            ))
            fig_pct.add_trace(go.Scatter(
                x=_ep_cats, y=_ep_bud_ln, name="EBIT % Budget Target",
                mode="lines",
                line=dict(color="#F59E0B", width=2, dash="dash"),
            ))
            if _nf_pct is not None:
                _ep_nf = [safe_float(_epct["NetFees_Pct"][_cat_index(_epct["cats"], _eb["cats"][i])]) * 100
                          for i in _ep_month_idxs]
                fig_pct.add_trace(go.Scatter(
                    x=_ep_cats, y=_ep_nf, name="Net Fees %",
                    mode="lines+markers",
                    line=dict(color="#10B981", width=2, dash="dot"),
                    marker=dict(size=6),
                ))
            fig_pct.update_layout(
                plot_bgcolor="white",
                font=dict(family=_FONT, color=PRIMARY),
                yaxis=dict(title="%", ticksuffix="%"),
                xaxis=dict(type="category"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                margin=dict(t=60, b=20),
                height=300,
            )
            st.plotly_chart(fig_pct, use_container_width=True, key=f"ebit_pct_{bu_name}_{sm}")


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

        # 4-column metric grid (one per BU) — monthly
        st.caption(f"Monthly — {sm} (kTL)")
        mcols = st.columns(4)
        for idx_m, k in enumerate(keys_oi):
            val   = oi[k][oi_idx] if oi_idx != -1 else 0
            delta = val - prev_oi[k][prev_oi_idx] if prev_oi and prev_oi_idx != -1 else None
            mcols[idx_m].metric(k, _mfmt(val), delta=_mfmt(delta) if delta is not None else None)

        # YTD OI (sum of monthly columns from Jan to current month)
        if oi_idx != -1:
            _oi_base = O.get("MON_OI_Base", 2)
            st.caption("YTD (kTL)")
            ytd_cols     = st.columns(5)
            ytd_oi_total = 0.0
            for idx_m, k in enumerate(keys_oi):
                ytd_val = sum(oi[k][i] for i in range(_oi_base, oi_idx + 1) if i < len(oi[k]))
                ytd_oi_total += ytd_val
                ytd_cols[idx_m].metric(f"YTD {k}", _mfmt(ytd_val))
            ytd_cols[4].metric("YTD Total", _mfmt(ytd_oi_total))

        # Stacked OI chart (all BUs combined over time)
        st.plotly_chart(
            chart_stacked(oi, "", keys_oi, [BU_COLORS.get(k, "#888") for k in keys_oi]),
            use_container_width=True, key=f"oi_chart_total_{sm}",
        )

        # ── Chart data table ─────────────────────────────────────────────────
        with st.expander("📊 Chart Data (kTL)", expanded=False):
            _oi_n = min(len(oi["cats"]), min(len(oi[k]) for k in keys_oi))
            _oi_tbl = []
            for _i in range(_oi_n):
                _row = {"Period": oi["cats"][_i]}
                _row_total = 0.0
                for _k in keys_oi:
                    _v = oi[_k][_i]
                    _row[_k] = fmt_ktl(_v)
                    _row_total += _v
                _row["Total"] = fmt_ktl(_row_total)
                _oi_tbl.append(_row)
            _df_oi_tbl = pd.DataFrame(_oi_tbl)
            st.dataframe(_df_oi_tbl, hide_index=True, use_container_width=True)
            st.download_button(
                "⬇ Export OI Data",
                df_to_excel_bytes(_df_oi_tbl),
                file_name=f"OI_Data_{sm.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"oi_chart_export_{sm}",
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
            f"Total Order Intake ({len(bg_oi)} projs) (kTL)",
            _mfmt(total_oi),
            delta=_mfmt(delta_oi) if delta_oi is not None else None,
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
    _base_wip = data["wip_projects"]
    if is_total:
        wip       = filter_wip(_base_wip, proj_sel, proj_text, cli_sel, cli_text)
        total_wip = sum(p["wip_tl"] for p in wip)
        delta_wip = None
        if prev_data:
            _prev_base = prev_data["wip_projects"]
            prev_total = sum(p["wip_tl"] for p in filter_wip(_prev_base, proj_sel, proj_text, cli_sel, cli_text))
            delta_wip  = fmt_tl_as_ktl(total_wip - prev_total)
    else:
        wip       = [p for p in filter_wip(_base_wip, proj_sel, proj_text, cli_sel, cli_text) if p["bu"] == bu_name]
        total_wip = sum(p["wip_tl"] for p in wip)
        delta_wip = None
        if prev_data:
            _prev_base = prev_data["wip_projects"]
            prev_wip  = [p for p in filter_wip(_prev_base, proj_sel, proj_text, cli_sel, cli_text) if p["bu"] == bu_name]
            delta_wip = fmt_tl_as_ktl(total_wip - sum(p["wip_tl"] for p in prev_wip))

    _wip_neg_count = sum(1 for p in wip if p["wip_tl"] < 0)
    _wip_lbl = f"Total WIP ≥1M ({len(wip)} projs"
    if _wip_neg_count:
        _wip_lbl += f", incl. {_wip_neg_count} negative"
    _wip_lbl += ") (kTL)"

    _grand = data.get("wip_gross_total", 0.0)

    _mc1, _mc2 = st.columns(2)
    _mc1.metric(_wip_lbl, fmt_tl_as_ktl(total_wip), delta=delta_wip)
    _mc2.metric("Grand Total WIP (all projects, kTL)", fmt_tl_as_ktl(_grand))

    fig = chart_wip([p for p in wip if p["wip_tl"] > 0], "")
    if fig:
        st.plotly_chart(fig, use_container_width=True, key=f"wip_fig_{bu_name}_{sm}")
    if wip:
        df_disp           = pd.DataFrame(wip)
        df_disp["WIP TL"] = df_disp["wip_tl"].apply(human_tl)
        cols_show = (
            ["name", "client", "bu", "orig_currency", "WIP TL"]
            if is_total
            else ["name", "client", "orig_currency", "WIP TL"]
        )
        # Total row
        _total_entry = {c: "" for c in cols_show}
        _total_entry["name"]    = "TOTAL"
        _total_entry["WIP TL"]  = human_tl(total_wip)
        df_disp_with_total = pd.concat(
            [df_disp[cols_show], pd.DataFrame([_total_entry])],
            ignore_index=True,
        )
        st.dataframe(
            df_disp_with_total, hide_index=True,
            use_container_width=True, key=f"wip_df_{bu_name}_{sm}",
        )
        st.download_button(
            "⬇ Export to Excel",
            df_to_excel_bytes(df_disp[cols_show]),
            file_name=f"WIP_{sm.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"wip_export_{bu_name}_{sm}",
        )

        # ── AG-Grid advanced view — hidden for now, re-enable by setting True ──────
        if False and _AGGRID_AVAILABLE:
            with st.expander("🔬 Advanced Table (AG-Grid)", expanded=False):
                st.caption(
                    "Sortable, filterable, groupable grid — independent of the table above. "
                    "Use column headers to sort; right-click for filter options."
                )
                _ag_df = df_disp[cols_show].copy()
                _gb    = GridOptionsBuilder.from_dataframe(_ag_df)
                _gb.configure_default_column(
                    resizable=True, sortable=True, filter=True,
                    wrapText=False, autoHeight=False,
                )
                if "bu" in cols_show:
                    _gb.configure_column("bu",  rowGroup=False, enableRowGroup=True)
                if "client" in cols_show:
                    _gb.configure_column("client", rowGroup=False, enableRowGroup=True)
                _gb.configure_column(
                    "WIP TL",
                    type=["numericColumn"],
                    cellStyle=JsCode(
                        "function(params){ "
                        "  if(params.value && params.value.startsWith('-'))"
                        "    return {'color':'#EF4444','fontWeight':'600'};"
                        "  return {'color':'#059669','fontWeight':'600'};"
                        "}"
                    ),
                )
                _gb.configure_selection("multiple", use_checkbox=False)
                _gb.configure_grid_options(domLayout="autoHeight")
                AgGrid(
                    _ag_df,
                    gridOptions=_gb.build(),
                    allow_unsafe_jscode=True,
                    use_container_width=True,
                    height=min(600, max(300, len(_ag_df) * 30 + 60)),
                    key=f"wip_aggrid_{bu_name}_{sm}",
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

tab1, tab2, tab3, tab4, tab5, tab6, tab_exec = st.tabs([
    "Net Sales", "EBIT", "Order Intake", "WIP", "Project History", "Scenarios", "Executive Summary"
])

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
                                   delta=fmt_ktl(_v2 - _v1) if _v1 else None)

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
                                   delta=fmt_ktl(_v2 - _v1) if _v1 else None)

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


def _pipeline_health_section(d_latest, sm_latest, key_pfx):
    """Render NS Pipeline Health by BU. key_pfx keeps widget keys unique."""
    _abbr_ph = sm_latest.split()[0][:3]
    _pipe_rows_ph = []
    for _bu in ["ENG", "MC", "T&SI", "NUC"]:
        _bd = d_latest["bu_ns"][_bu]
        _mi = _cat_index(_bd["cats"], _abbr_ph)
        if _mi == -1:
            continue
        _ord = _bd["Order"][_mi] if _mi < len(_bd["Order"]) else 0
        _off = _bd["Offer"][_mi] if _mi < len(_bd["Offer"]) else 0
        _opp = _bd["Opp"][_mi]   if _mi < len(_bd["Opp"])   else 0
        _tot = _ord + _off + _opp
        _tgt = sum(
            _bd[t][_BU_TARGET_IDX] for t in ["Order", "Offer", "Opp"]
            if _BU_TARGET_IDX < len(_bd[t])
        )
        _pipe_rows_ph.append({
            "BU":               _bu,
            "Order":            _ord,
            "Offer":            _off,
            "Opp":              _opp,
            "Total":            _tot,
            "Year-End Target":  _tgt,
            "YTD vs Tgt%":      round(_tot / _tgt * 100, 1) if _tgt > 0 else 0,
            "Secured%":         round(_ord / _tot * 100, 1) if _tot > 0 else 0,
            "Secured+Offer%":   round((_ord + _off) / _tot * 100, 1) if _tot > 0 else 0,
        })

    if not _pipe_rows_ph:
        st.warning("No pipeline data available.")
        return

    _df_ph = pd.DataFrame(_pipe_rows_ph)

    # Stacked bar chart
    _df_melt_ph = _df_ph.melt(
        id_vars=["BU"], value_vars=["Order", "Offer", "Opp"],
        var_name="Stage", value_name="Value (kTL)",
    )
    _bu_stage_cmap_ph = {}
    for _buc in ["ENG", "MC", "T&SI", "NUC"]:
        _tc = bu_tier_colors(_buc)
        _bu_stage_cmap_ph[f"{_buc} Order"] = _tc[0]
        _bu_stage_cmap_ph[f"{_buc} Offer"] = _tc[1]
        _bu_stage_cmap_ph[f"{_buc} Opp"]   = _tc[2]
    _df_melt_ph["BU_Stage"] = _df_melt_ph["BU"] + " " + _df_melt_ph["Stage"]
    _ph_seg_max = _df_melt_ph["Value (kTL)"].abs().max() or 1
    _df_melt_ph["Label"] = _df_melt_ph["Value (kTL)"].apply(
        lambda v: human_k(v) if abs(v) / _ph_seg_max >= _CHART_LABEL_PCT else ""
    )
    _fig_ph = px.bar(
        _df_melt_ph, x="BU", y="Value (kTL)", color="BU_Stage",
        barmode="stack",
        color_discrete_map=_bu_stage_cmap_ph,
        text="Label",
        category_orders={"BU_Stage": [
            f"{b} {s}" for b in ["ENG", "MC", "T&SI", "NUC"]
            for s in ["Order", "Offer", "Opp"]
        ]},
    )
    _fig_ph.update_layout(
        plot_bgcolor="white",
        font=dict(family=_FONT, color=PRIMARY),
        legend=dict(
            orientation="h", yanchor="top", y=-0.18,
            xanchor="center", x=0.5, font=dict(size=11),
        ),
        legend_title_text="",
        margin=dict(t=20, b=160),
    )
    _fig_ph.update_traces(
        textposition="inside", insidetextanchor="middle",
        textfont=dict(size=11, color="white"), cliponaxis=False,
    )
    st.plotly_chart(_fig_ph, use_container_width=True, key=f"{key_pfx}_pipe_chart")

    # Conversion rates vs year-end target
    st.markdown("**Conversion Rates vs Year-End Target**")
    _ph_cr_cols = st.columns(len(_pipe_rows_ph))
    for _ci, _row in enumerate(_pipe_rows_ph):
        _buc = _ph_cr_cols[_ci]
        _buc.markdown(f"**{_row['BU']}**")
        _ytd_frac = min(1.0, max(0.0, _row["YTD vs Tgt%"] / 100))
        _buc.progress(_ytd_frac, text=f"{_row['YTD vs Tgt%']}% of target")
        _buc.metric("YTD Pipeline (kTL)",    fmt_ktl(_row["Total"]))
        _buc.metric("Year-End Target (kTL)", fmt_ktl(_row["Year-End Target"]))
        _buc.metric("Secured (Order%)",      f"{_row['Secured%']}%")
        _buc.metric("Secured+Offer%",        f"{_row['Secured+Offer%']}%")

    # Summary table with company total row
    _comp_tot = _df_ph["Total"].sum()
    _comp_tgt = _df_ph["Year-End Target"].sum()
    _df_ph_disp = _df_ph.copy()
    for _fc in ["Total", "Year-End Target", "Order", "Offer", "Opp"]:
        _df_ph_disp[_fc] = _df_ph_disp[_fc].apply(fmt_ktl)
    # Convert rate columns to strings now so the TOTAL row "—" doesn't mix types
    _df_ph_disp["YTD vs Tgt%"]    = _df_ph_disp["YTD vs Tgt%"].apply(lambda v: f"{v}%")
    _df_ph_disp["Secured%"]        = _df_ph_disp["Secured%"].apply(lambda v: f"{v}%")
    _df_ph_disp["Secured+Offer%"]  = _df_ph_disp["Secured+Offer%"].apply(lambda v: f"{v}%")
    _total_ph_row = {
        "BU":               "TOTAL",
        "Order":            fmt_ktl(_df_ph["Order"].sum()),
        "Offer":            fmt_ktl(_df_ph["Offer"].sum()),
        "Opp":              fmt_ktl(_df_ph["Opp"].sum()),
        "Total":            fmt_ktl(_comp_tot),
        "Year-End Target":  fmt_ktl(_comp_tgt),
        "YTD vs Tgt%":      f"{round(_comp_tot / _comp_tgt * 100, 1)}%" if _comp_tgt > 0 else "—",
        "Secured%":         "—",
        "Secured+Offer%":   "—",
    }
    _df_ph_display = pd.concat(
        [_df_ph_disp, pd.DataFrame([_total_ph_row])], ignore_index=True,
    )
    st.dataframe(
        _df_ph_display[["BU", "Order", "Offer", "Opp", "Total", "Year-End Target",
                         "YTD vs Tgt%", "Secured%", "Secured+Offer%"]],
        hide_index=True, use_container_width=True,
    )


@st.fragment
def _render_exec_summary(loaded_data, selected_months, multi_year):
    """Executive Summary — top-line KPIs, YTD vs target, BU snapshot, MoM delta."""
    _sm_l    = lambda sm: sm if multi_year else sm.split()[0]
    _sm_cur  = selected_months[-1]
    _d       = loaded_data.get(_sm_cur)
    if not _d:
        st.warning("No data for the selected month.")
        return

    _abbr    = _sm_cur.split()[0][:3]
    _ns      = _d["ns"]
    _ebit    = _d["ebit"]
    _oi      = _d["oi"]

    _ns_i    = _cat_index(_ns["cats"],   _abbr)
    _eb_i    = _cat_index(_ebit["cats"], _abbr)
    _oi_i    = _cat_index(_oi["cats"],   _abbr)
    _eoy_i   = _find_eoy_idx(_ns["cats"])
    _eb_eoy  = _find_eoy_idx(_ebit["cats"])
    _t_idx   = next((i for i, c in enumerate(_ns["cats"])   if "target" in str(c).lower()), -1)
    _et_idx  = next((i for i, c in enumerate(_ebit["cats"]) if "target" in str(c).lower()), -1)

    # ── Per-month values ──────────────────────────────────────────────────────
    _ns_m    = monthly_val(_ns["Contract"],   _ns["cats"],   _ns_i)
    _ebit_m  = monthly_val(_ebit["Contract"], _ebit["cats"], _eb_i)
    _ns_ytd  = _ns["Contract"][_ns_i]           if _ns_i  != -1 and _ns_i  < len(_ns["Contract"])   else 0
    _eb_ytd  = _ebit["Contract"][_eb_i]         if _eb_i  != -1 and _eb_i  < len(_ebit["Contract"])  else 0
    _ns_eoy  = _ns["Contract"][_eoy_i]          if _eoy_i != -1 and _eoy_i < len(_ns["Contract"])    else 0
    _eb_eoy  = _ebit["Contract"][_eb_eoy]       if _eb_eoy != -1 and _eb_eoy < len(_ebit["Contract"]) else 0
    _ns_tgt  = _ns["Contract"][_t_idx]          if _t_idx != -1 and _t_idx  < len(_ns["Contract"])   else 0
    _eb_tgt  = _ebit["Contract"][_et_idx]       if _et_idx != -1 and _et_idx < len(_ebit["Contract"]) else 0
    _ebit_pct = _eb_ytd / _ns_ytd * 100         if _ns_ytd != 0 else 0
    _bud_pct  = _eb_tgt / _ns_tgt * 100         if _ns_tgt != 0 else 0

    _oi_base = O.get("MON_OI_Base", 2)
    _oi_m    = _oi["ENG"][_oi_i] + _oi["MC"][_oi_i] + _oi["T&SI"][_oi_i] + _oi["NUC"][_oi_i] \
               if _oi_i != -1 else 0
    _oi_ytd  = sum(
        sum(_oi[k][i] for i in range(_oi_base, (_oi_i + 1) if _oi_i != -1 else _oi_base)
            if i < len(_oi[k]))
        for k in ["ENG", "MC", "T&SI", "NUC"]
    )

    # MoM comparison
    _prev_sm   = selected_months[-2] if len(selected_months) >= 2 else None
    _prev_d    = loaded_data.get(_prev_sm) if _prev_sm else None
    _prev_abbr = _prev_sm.split()[0][:3] if _prev_sm else None

    def _prev_m(series_dict, series_key):
        if not _prev_d or not _prev_abbr:
            return None
        _s = _prev_d[series_dict][series_key]
        _c = _prev_d[series_dict]["cats"]
        _i = _cat_index(_c, _prev_abbr)
        return monthly_val(_s, _c, _i) if _i != -1 else None

    _p_ns_m  = _prev_m("ns",   "Contract")
    _p_eb_m  = _prev_m("ebit", "Contract")

    st.markdown(f"### Executive Summary — {_sm_l(_sm_cur)}")
    st.caption("Top-line figures for the current period. YTD values are cumulative; EOY is the December/year-end forecast.")

    # ── Shared computed values ────────────────────────────────────────────────
    _epct_exec    = _d.get("ebit_pct", {})
    _ep_exec_idx  = _cat_index(_epct_exec.get("cats", []), _abbr) if _epct_exec else -1
    _v_act_exec   = safe_float(_epct_exec["EBIT_Pct"][_ep_exec_idx])        if _ep_exec_idx != -1 and _epct_exec else _ebit_pct / 100
    _v_bud_exec   = safe_float(_epct_exec["EBIT_Pct_Budget"][_ep_exec_idx]) if _ep_exec_idx != -1 and _epct_exec else _bud_pct / 100
    _v_nf_exec    = safe_float(_epct_exec["NetFees_Pct"][_ep_exec_idx])     if _ep_exec_idx != -1 and _epct_exec else 0.0

    def _nv(k):
        return _ns[k][_ns_i]  if _ns_i  != -1 and _ns_i  < len(_ns[k])  else 0
    def _nt(k):
        return _ns[k][_t_idx] if _t_idx != -1 and _t_idx < len(_ns[k])  else 0
    def _ev(k):
        return _ebit[k][_eb_i]   if _eb_i   != -1 and _eb_i   < len(_ebit[k])  else 0
    def _et(k):
        return _ebit[k][_et_idx] if _et_idx != -1 and _et_idx < len(_ebit[k])  else 0

    def _pbar4(lbl, ytd_v, tgt_v):
        """Progress-bar row with abbreviated numbers to avoid truncation."""
        _p = ytd_v / tgt_v if tgt_v > 0 else 0
        _c1, _c2, _c3, _c4 = st.columns([2.5, 1.5, 1.5, 1])
        _c1.markdown(f"**{lbl}**")
        _c1.progress(min(1.0, max(0.0, _p)))
        _c2.metric("YTD (kTL)",    human_k(ytd_v), help=fmt_ktl(ytd_v))
        _c3.metric("Target (kTL)", human_k(tgt_v), help=fmt_ktl(tgt_v))
        _c4.metric("Coverage", f"{_p:.0%}")

    # ── Section selector (radio avoids nested-tabs-in-fragment rendering glitch) ─
    _exec_section = st.radio(
        "Section", ["Overview", "YTD vs Target", "Rolling 12M", "Pipeline Health"],
        horizontal=True, key="exec_section_sel",
        label_visibility="collapsed",
    )
    st.markdown("---")

    # ── Overview ─────────────────────────────────────────────────────────────
    if _exec_section == "Overview":
        st.markdown("#### Net Sales (kTL)")
        _nc1, _nc2, _nc3, _nc4 = st.columns(4)
        _nc1.metric("NS — Monthly",  fmt_ktl(_ns_m),
                    delta=fmt_ktl(_ns_m - _p_ns_m) if _p_ns_m is not None else None)
        _nc2.metric("NS — YTD",      fmt_ktl(_ns_ytd))
        _nc3.metric("NS — EOY Fcst", fmt_ktl(_ns_eoy))
        _nc4.metric("NS — Target",   fmt_ktl(_ns_tgt),
                    delta=f"{_ns_ytd / _ns_tgt:.0%} achieved" if _ns_tgt else None,
                    delta_color="off")

        st.markdown("#### EBIT (kTL)")
        _ec1, _ec2, _ec3, _ec4 = st.columns(4)
        _ec1.metric("EBIT — Monthly",  fmt_ktl(_ebit_m),
                    delta=fmt_ktl(_ebit_m - _p_eb_m) if _p_eb_m is not None else None)
        _ec2.metric("EBIT — YTD",      fmt_ktl(_eb_ytd))
        _ec3.metric("EBIT — EOY Fcst", fmt_ktl(_eb_eoy))
        _ec4.metric("EBIT — Target",   fmt_ktl(_eb_tgt),
                    delta=f"{_eb_ytd / _eb_tgt:.0%} achieved" if _eb_tgt else None,
                    delta_color="off")

        st.markdown("#### Margins & Order Intake")
        _mc1, _mc2, _mc3, _mc4, _mc5 = st.columns(5)
        _mc1.metric("EBIT % (Actual)",   f"{_v_act_exec * 100:.1f}%",
                    delta=f"{(_v_act_exec - _v_bud_exec) * 100:+.1f} pp vs Budget")
        _mc2.metric("EBIT % (Budget)",   f"{_v_bud_exec * 100:.1f}%")
        _mc3.metric("Net Fees %",        f"{_v_nf_exec * 100:.1f}%")
        _mc4.metric("OI — Monthly (kTL)", fmt_ktl(_oi_m))
        _mc5.metric("OI — YTD (kTL)",    fmt_ktl(_oi_ytd))

        st.markdown("---")
        st.markdown("#### Business Unit Snapshot")
        _bu_rows = []
        for _bu in ["ENG", "MC", "T&SI", "NUC"]:
            _bd_ns   = _d["bu_ns"][_bu]
            _bd_eb   = _d["bu_ebit"][_bu]
            _bns_i   = _cat_index(_bd_ns["cats"], _abbr)
            _beb_i   = _cat_index(_bd_eb["cats"], _abbr)
            _bns_tgt_i = _BU_TARGET_IDX
            _bns_m   = sum(monthly_val(_bd_ns[t], _bd_ns["cats"], _bns_i) for t in ["Order", "Offer", "Opp"]) if _bns_i != -1 else 0
            _bns_ytd = sum(_bd_ns[t][_bns_i] for t in ["Order", "Offer", "Opp"] if _bns_i != -1 and _bns_i < len(_bd_ns[t])) if _bns_i != -1 else 0
            _bns_tgt = sum(_bd_ns[t][_bns_tgt_i] for t in ["Order", "Offer", "Opp"] if _bns_tgt_i < len(_bd_ns[t]))
            _beb_m   = monthly_val(_bd_eb["Total"], _bd_eb["cats"], _beb_i) if _beb_i != -1 else 0
            _beb_ytd = _bd_eb["Total"][_beb_i] if _beb_i != -1 and _beb_i < len(_bd_eb["Total"]) else 0
            _bns_pct = _bns_ytd / _bns_tgt * 100 if _bns_tgt != 0 else 0
            _bu_rows.append({
                "BU":                 _bu,
                "NS Monthly (kTL)":   fmt_ktl(_bns_m),
                "NS YTD (kTL)":       fmt_ktl(_bns_ytd),
                "NS Target (kTL)":    fmt_ktl(_bns_tgt),
                "vs Target":          f"{_bns_pct:.0f}%",
                "EBIT Monthly (kTL)": fmt_ktl(_beb_m),
                "EBIT YTD (kTL)":     fmt_ktl(_beb_ytd),
            })
        st.dataframe(pd.DataFrame(_bu_rows), hide_index=True, use_container_width=True)
        st.download_button(
            "⬇ Export Exec Summary",
            df_to_excel_bytes(pd.DataFrame(_bu_rows)),
            file_name=f"ExecSummary_{_sm_cur.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="exec_export",
        )

        if _prev_sm and _prev_d:
            st.markdown("---")
            _is_yoy  = _sm_cur.split()[1] != _prev_sm.split()[1]
            _cmp_tag = "YoY" if _is_yoy else "MoM"
            st.markdown(f"#### {_cmp_tag} Δ — {_sm_l(_prev_sm)} → {_sm_l(_sm_cur)}")
            _p_ns_ytd  = _prev_d["ns"]["Contract"][_cat_index(_prev_d["ns"]["cats"], _prev_abbr)] \
                         if _cat_index(_prev_d["ns"]["cats"], _prev_abbr) != -1 else 0
            _p_eb_ytd  = _prev_d["ebit"]["Contract"][_cat_index(_prev_d["ebit"]["cats"], _prev_abbr)] \
                         if _cat_index(_prev_d["ebit"]["cats"], _prev_abbr) != -1 else 0
            _dd1, _dd2, _dd3, _dd4 = st.columns(4)
            _dd1.metric("NS Monthly Δ",   fmt_ktl(_ns_m   - (_p_ns_m or 0)),
                        delta=f"{(_ns_m - (_p_ns_m or 0)) / abs(_p_ns_m) * 100:+.1f}%" if _p_ns_m else None)
            _dd2.metric("NS YTD Δ",       fmt_ktl(_ns_ytd - _p_ns_ytd),
                        delta=f"{(_ns_ytd - _p_ns_ytd) / abs(_p_ns_ytd) * 100:+.1f}%" if _p_ns_ytd else None)
            _dd3.metric("EBIT Monthly Δ", fmt_ktl(_ebit_m - (_p_eb_m or 0)),
                        delta=f"{(_ebit_m - (_p_eb_m or 0)) / abs(_p_eb_m) * 100:+.1f}%" if _p_eb_m else None)
            _dd4.metric("EBIT YTD Δ",     fmt_ktl(_eb_ytd - _p_eb_ytd),
                        delta=f"{(_eb_ytd - _p_eb_ytd) / abs(_p_eb_ytd) * 100:+.1f}%" if _p_eb_ytd else None)

    # ── YTD vs Target ────────────────────────────────────────────────────────
    elif _exec_section == "YTD vs Target":
        st.markdown(f"#### YTD vs Target — {_sm_l(_sm_cur)}")
        st.caption("Cumulative YTD actuals vs annual target. Hover the number cards for full precision. NS tiers are independent; EBIT tiers are cumulative.")

        if _t_idx != -1 and _ns_i != -1:
            st.markdown("**Net Sales**")
            _ns_contract_ytd = _nv("Contract")
            _ns_wp_ytd       = _nv("WP")
            _ns_wo_ytd       = _nv("WO")
            _pbar4("Contract",                         _ns_contract_ytd, _nt("Contract"))
            _pbar4("Weighted Proposals (WP)",          _ns_wp_ytd,       _nt("WP"))
            _ytd_fo  = _ns_contract_ytd + _ns_wp_ytd + _ns_wo_ytd
            _tgt_fo  = sum(_nt(k) for k in ["Contract", "WP", "WO"])
            _pbar4("Full Outlook (Contract + WP + WO)", _ytd_fo, _tgt_fo)
        else:
            st.info("Target column not found in NS data.")

        if _et_idx != -1 and _eb_i != -1:
            st.markdown("---")
            st.markdown("**EBIT**")
            _pbar4("Contract",                        _ev("Contract"),       _et("Contract"))
            _pbar4("Contract + Weighted Proposals",   _ev("Contract+WP"),    _et("Contract+WP"))
            _pbar4("Full Outlook (Contract+WP+WO)",   _ev("Contract+WP+WO"), _et("Contract+WP+WO"))

        if _ep_exec_idx != -1 and _epct_exec:
            st.markdown("---")
            st.markdown("**Margin Rates — current period**")
            _xmr1, _xmr2, _xmr3 = st.columns(3)
            _xmr1.metric("EBIT % (Actual)",  f"{_v_act_exec * 100:.1f}%",
                         delta=f"{(_v_act_exec - _v_bud_exec) * 100:+.1f} pp vs Budget")
            _xmr2.metric("EBIT % (Budget)",  f"{_v_bud_exec * 100:.1f}%")
            _xmr3.metric("Net Fees %",        f"{_v_nf_exec * 100:.1f}%")

    # ── Rolling 12M ──────────────────────────────────────────────────────────
    elif _exec_section == "Rolling 12M":
        st.markdown("#### Rolling 12-Month Net Sales")
        st.caption(
            "Monthly NS over the trailing 12 months. "
            "Current-year months use their own file; previous-year months use the December file (fully closed)."
        )

        _r12_view = st.radio(
            "Breakdown", ["Company Total", "By Business Unit", "By Project", "By Client"],
            horizontal=True, key="exec_r12_view",
        )

        _cur_month_name = _sm_cur.split()[0]
        _cur_year       = int(_sm_cur.split()[1])
        _cur_month_num  = MONTH_ORDER.index(_cur_month_name) + 1

        _trailing_slots = []
        for _offset in range(11, -1, -1):
            _mo, _yr = _cur_month_num - _offset, _cur_year
            while _mo <= 0:
                _mo += 12; _yr -= 1
            _trailing_slots.append((_yr, _mo, MONTH_ORDER[_mo - 1]))

        _avail_sorted = sorted(
            EXCEL_FILES.keys(),
            key=lambda x: (int(x.split()[1]), MONTH_ORDER.index(x.split()[0])),
        )
        _r12_labels = [f"{mn[:3]} {str(yr)[2:]}" for yr, _, mn in _trailing_slots]

        def _file_for_slot(yr, mn_name):
            if yr == _cur_year:
                # Current year: use exact month file, fall back to latest available
                key = f"{mn_name} {yr}"
                if key in EXCEL_FILES:
                    return key
                yf = [k for k in _avail_sorted if k.split()[1] == str(yr)]
                return yf[-1] if yf else None
            else:
                # Previous years: always prefer December (fully closed annual data)
                dec = f"December {yr}"
                if dec in EXCEL_FILES:
                    return dec
                yf = [k for k in _avail_sorted if k.split()[1] == str(yr)]
                return yf[-1] if yf else None

        def _load_slot(yr, mn_name):
            fk = _file_for_slot(yr, mn_name)
            if not fk:
                return None, None
            fp = EXCEL_FILES[fk]
            return load_month(str(fp)) if fp and fp.exists() else None, fk

        _r12_cur_sfx = str(_cur_year)[2:]

        def _r12_bar_chart(fig, key):
            fig.update_layout(
                plot_bgcolor="white", font=dict(family=_FONT, color=PRIMARY),
                yaxis_title="kTL", xaxis=dict(type="category"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                margin=dict(t=60, b=20), height=340,
            )
            st.plotly_chart(fig, use_container_width=True, key=key)
            st.caption(f"Current year ({_cur_year}) bars are highlighted; prior year in grey.")

        if _r12_view == "Company Total":
            _r12_vals = []
            for _yr_s, _, _mn_s in _trailing_slots:
                _d_slot, _ = _load_slot(_yr_s, _mn_s)
                if _d_slot:
                    _i = _cat_index(_d_slot["ns"]["cats"], _mn_s[:3])
                    _r12_vals.append(monthly_val(_d_slot["ns"]["Contract"], _d_slot["ns"]["cats"], _i) if _i != -1 else 0)
                else:
                    _r12_vals.append(0)
            if any(v != 0 for v in _r12_vals):
                _lf = [human_k(v) for v in _r12_vals]
                _fig_r12 = go.Figure(go.Bar(
                    x=_r12_labels, y=_r12_vals, text=_lf, textposition="outside",
                    cliponaxis=False,
                    marker_color=[PRIMARY if lbl.endswith(_r12_cur_sfx) else "#94A3B8" for lbl in _r12_labels],
                    hovertemplate="%{x}: %{customdata} kTL<extra></extra>", customdata=_lf,
                ))
                _r12_bar_chart(_fig_r12, "exec_r12_total")
                _mom = [None] + [_r12_vals[i] - _r12_vals[i-1] for i in range(1, len(_r12_vals))]
                _df_r12 = pd.DataFrame({
                    "Period": _r12_labels,
                    "NS Contract (kTL)": [fmt_ktl(v) for v in _r12_vals],
                    "MoM Δ": [f"+{human_k(d)}" if d is not None and d >= 0 else human_k(d) if d is not None else "—" for d in _mom],
                })
                st.dataframe(_df_r12, hide_index=True, use_container_width=True)
            else:
                st.info("No NS data found for the trailing 12 months.")

        elif _r12_view == "By Business Unit":
            _bu_series: dict[str, list] = {bu: [] for bu in ["ENG", "MC", "T&SI", "NUC"]}
            for _yr_s, _, _mn_s in _trailing_slots:
                _d_slot, _ = _load_slot(_yr_s, _mn_s)
                for _bu in ["ENG", "MC", "T&SI", "NUC"]:
                    if _d_slot:
                        _bd = _d_slot["bu_ns"][_bu]
                        _bi = _cat_index(_bd["cats"], _mn_s[:3])
                        _v  = sum(monthly_val(_bd[t], _bd["cats"], _bi) for t in ["Order","Offer","Opp"]) if _bi != -1 else 0
                    else:
                        _v = 0
                    _bu_series[_bu].append(_v)

            if any(any(v != 0 for v in vals) for vals in _bu_series.values()):
                # Compute per-slot totals to determine label visibility threshold
                _r12_bu_totals = [
                    sum(_bu_series[_bu][i] for _bu in ["ENG", "MC", "T&SI", "NUC"])
                    for i in range(len(_r12_labels))
                ]
                _r12_seg_max = max(_r12_bu_totals) if _r12_bu_totals else 1

                _fig_bu = go.Figure()
                for _bu in ["ENG", "MC", "T&SI", "NUC"]:
                    _vals = _bu_series[_bu]
                    # Show label only when the segment is tall enough to hold it
                    _texts = [
                        human_k(v) if v > 0 and v / _r12_seg_max >= _CHART_LABEL_PCT else ""
                        for v in _vals
                    ]
                    _fig_bu.add_trace(go.Bar(
                        x=_r12_labels, y=_vals, name=_bu,
                        marker_color=BU_COLORS.get(_bu, "#888"),
                        text=_texts,
                        textposition="inside",
                        insidetextanchor="middle",
                        constraintext="inside",
                        textfont=dict(size=11, color="white"),
                        hovertemplate=f"<b>{_bu}</b> %{{x}}: %{{y:,.0f}} kTL<extra></extra>",
                    ))

                # Total label above each bar via Scatter text trace
                _fig_bu.add_trace(go.Scatter(
                    x=_r12_labels,
                    y=_r12_bu_totals,
                    mode="text",
                    text=[human_k(t) if t > 0 else "" for t in _r12_bu_totals],
                    textposition="top center",
                    textfont=dict(size=10, color=PRIMARY, family=_FONT),
                    showlegend=False,
                    hoverinfo="skip",
                ))

                _fig_bu.update_layout(barmode="stack")
                _r12_bar_chart(_fig_bu, "exec_r12_bu")

                _tbl_rows = []
                for i, lbl in enumerate(_r12_labels):
                    _row = {"Period": lbl}
                    _tot = 0
                    for _bu in ["ENG", "MC", "T&SI", "NUC"]:
                        _row[f"{_bu} (kTL)"] = fmt_ktl(_bu_series[_bu][i])
                        _tot += _bu_series[_bu][i]
                    _row["Total (kTL)"] = fmt_ktl(_tot)
                    _tbl_rows.append(_row)
                st.dataframe(pd.DataFrame(_tbl_rows), hide_index=True, use_container_width=True)
            else:
                st.info("No BU NS data found for the trailing 12 months.")

        else:
            _is_proj = (_r12_view == "By Project")
            _ref_fk  = _file_for_slot(_cur_year, _cur_month_name)
            _ref_fp  = EXCEL_FILES.get(_ref_fk) if _ref_fk else None
            _ref_md  = load_margin_data(str(_ref_fp)) if _ref_fp and _ref_fp.exists() else None

            if not _ref_md:
                st.info("Ext. Prod. data unavailable — cannot build project/client breakdown.")
            else:
                if _is_proj:
                    _opts = sorted({info["name"] for info in _ref_md["projects"].values()})
                    _sel  = st.multiselect("Select Projects (max 5)", _opts, max_selections=5, key="exec_r12_projs")
                else:
                    _opts = sorted({info["client"] for info in _ref_md["projects"].values() if info.get("client")})
                    _sel  = st.multiselect("Select Clients (max 5)",  _opts, max_selections=5, key="exec_r12_clis")

                if not _sel:
                    st.info("Select at least one option to display.")
                else:
                    _series: dict[str, list] = {s: [] for s in _sel}
                    for _yr_s, _, _mn_s in _trailing_slots:
                        _fk_s = _file_for_slot(_yr_s, _mn_s)
                        _fp_s = EXCEL_FILES.get(_fk_s) if _fk_s else None
                        _md_s = load_margin_data(str(_fp_s)) if _fp_s and _fp_s.exists() else None
                        _month_lbl = f"{_mn_s} {_yr_s}"
                        for _sel_item in _sel:
                            _v = 0
                            if _md_s:
                                for _pi in _md_s["projects"].values():
                                    _match = (_pi["name"] == _sel_item) if _is_proj else (_pi.get("client") == _sel_item)
                                    if _match:
                                        _gf = _pi.get("type_data", {}).get("GROSS FEES", {})
                                        _v += _gf.get(_month_lbl, 0)
                            _series[_sel_item].append(_v)

                    _lcolors = [PRIMARY, "#10B981", "#F59E0B", "#8B5CF6", "#EC4899"]
                    _fig_sel = go.Figure()
                    for _ci, _sel_item in enumerate(_sel):
                        _vals = _series[_sel_item]
                        _lbf  = [human_tl(v) for v in _vals]
                        _fig_sel.add_trace(go.Scatter(
                            x=_r12_labels, y=_vals, name=_sel_item[:30],
                            mode="lines+markers+text",
                            text=_lbf, textposition="top center",
                            textfont=dict(size=9),
                            line=dict(color=_lcolors[_ci % len(_lcolors)], width=2),
                            marker=dict(size=7),
                            hovertemplate=f"<b>{_sel_item}</b> %{{x}}: %{{customdata}}<extra></extra>",
                            customdata=_lbf,
                        ))
                    _fig_sel.update_layout(
                        plot_bgcolor="white", font=dict(family=_FONT, color=PRIMARY),
                        yaxis_title="TL",
                        xaxis=dict(type="category", tickangle=-45),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02),
                        margin=dict(t=60, b=20), height=360,
                    )
                    st.plotly_chart(_fig_sel, use_container_width=True, key="exec_r12_sel")

                    _tbl_rows = []
                    for i, lbl in enumerate(_r12_labels):
                        _row = {"Period": lbl}
                        for _sel_item in _sel:
                            _row[_sel_item[:25]] = human_tl(_series[_sel_item][i])
                        _tbl_rows.append(_row)
                    _df_sel = pd.DataFrame(_tbl_rows)
                    st.dataframe(_df_sel, hide_index=True, use_container_width=True)
                    st.download_button(
                        "⬇ Export to Excel",
                        df_to_excel_bytes(_df_sel),
                        file_name=f"NS_12M_{'Project' if _is_proj else 'Client'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="exec_r12_export",
                    )

    # ── Pipeline Health ───────────────────────────────────────────────────────
    else:  # Pipeline Health
        if _d:
            st.markdown("#### NS Pipeline Health by BU")
            st.caption(
                "YTD cumulative pipeline (Order/Offer/Opp) per BU vs year-end target. "
                "Order = contracted. Offer = pending. Opp = identified opportunities. All values in kTL."
            )
            _pipeline_health_section(_d, _sm_cur, "exec")


@st.fragment
def _render_project_history(loaded_data, selected_months, global_bu_view, multi_year):
    st.markdown("### Project History")
    st.caption(
        "Track project-based monthly Production (TL) and project WIP TL history. "
        "Production values are loaded from 'Ext. Prod.' and WIP values from the monthly WIP data."
    )

    with st.spinner("Loading margin data…"):
        margin_data = merge_margin_data(tuple(selected_months))

    # Per-file data for breakdown comparison (current vs previous month file)
    _all_avail_sm  = sorted(EXCEL_FILES.keys(),
                            key=lambda x: (x.split()[1], MONTH_ORDER.index(x.split()[0])))
    _cur_sm_ph     = selected_months[-1]
    _cur_pos_ph    = next((i for i, x in enumerate(_all_avail_sm) if x == _cur_sm_ph), -1)
    _prev_sm_ph    = _all_avail_sm[_cur_pos_ph - 1] if _cur_pos_ph > 0 else None
    _cur_fp_ph     = EXCEL_FILES.get(_cur_sm_ph)
    _prev_fp_ph    = EXCEL_FILES.get(_prev_sm_ph) if _prev_sm_ph else None
    _cur_file_md   = load_margin_data(str(_cur_fp_ph))  if _cur_fp_ph  and _cur_fp_ph.exists()  else None
    _prev_file_md  = load_margin_data(str(_prev_fp_ph)) if _prev_fp_ph and _prev_fp_ph.exists() else None

    if margin_data and margin_data.get("projects"):
        projects    = margin_data["projects"]
        month_cols  = margin_data["month_cols"]
        bl_year     = margin_data.get("bl_year")
        bo_year     = margin_data.get("bo_year")

        # ── Filters ───────────────────────────────────────────────────────────
        _all_ph_clients = sorted(
            {info["client"] for info in projects.values()
             if info["client"] and info["client"].lower() not in ("", "nan")}
        )
        _fc1, _fc2 = st.columns(2)
        with _fc1:
            if global_bu_view == "Company Total":
                ph_bu_filter = st.multiselect(
                    "Business Unit",
                    ["ENG", "MC", "T&SI", "NUC"],
                    default=["ENG", "MC", "T&SI", "NUC"],
                    key="ph_bu_filter",
                )
            else:
                # Lock to the global BU selection to prevent empty WIP/chart results
                ph_bu_filter = [global_bu_view]
                st.multiselect(
                    "Business Unit",
                    ["ENG", "MC", "T&SI", "NUC"],
                    default=[global_bu_view],
                    key="ph_bu_filter",
                    disabled=True,
                    help=f"Locked to {global_bu_view} by the top Business Unit selector. Switch to 'Company Total' to search across all BUs.",
                )
        with _fc2:
            ph_cli_filter = st.multiselect(
                "Client (type to search)",
                _all_ph_clients,
                key="ph_cli_filter",
                help="Leave empty to show all clients. Select one or more to filter.",
            )

        filtered_projects = {
            key: info for key, info in projects.items()
            if info["bu"] in ph_bu_filter
            and (not ph_cli_filter or info["client"] in ph_cli_filter)
        }

        # Build display labels: "Project Name" or "Project Name (CURRENCY)" for duplicate names
        _name_counts: dict[str, int] = {}
        for info in filtered_projects.values():
            _name_counts[info["name"]] = _name_counts.get(info["name"], 0) + 1

        def _display_label(info: dict) -> str:
            name = info["name"]
            return f"{name} ({info['currency']})" if _name_counts.get(name, 0) > 1 else name

        label_to_key  = {_display_label(v): k for k, v in filtered_projects.items()}
        all_labels    = sorted(label_to_key.keys())

        # ── Project selector + compare controls ───────────────────────────────
        sel_label = None
        _compare_proj_keys: list[str] = []
        _compare_cli_keys:  list[str] = []

        if not all_labels:
            st.warning("No projects found for the selected filters.")
        else:
            _ps1, _ps2 = st.columns([3, 1])
            with _ps1:
                ph_proj_text = st.text_input(
                    "Search Project", placeholder="Type to narrow the list…", key="ph_proj_text",
                )
                visible_labels = (
                    [l for l in all_labels if ph_proj_text.lower() in l.lower()]
                    if ph_proj_text else all_labels
                )
                sel_label = st.selectbox(
                    "Select Project", visible_labels if visible_labels else all_labels,
                    key="margin_proj",
                )
            with _ps2:
                ph_compare_on = st.checkbox("Compare projects", key="ph_compare_on")
                ph_compare_cli = st.checkbox("Compare clients", key="ph_compare_cli")

            if ph_compare_on:
                _other_labels = [l for l in all_labels if l != sel_label]
                _compare_raw = st.multiselect(
                    "Add up to 2 more projects (chart only)",
                    _other_labels, max_selections=2, key="ph_compare_ms",
                    help="Lines added to the production chart. Detail table stays for the primary project.",
                )
                _compare_proj_keys = [label_to_key[l] for l in _compare_raw if l in label_to_key]

            if ph_compare_cli:
                _other_clients = [c for c in _all_ph_clients
                                  if sel_label and c != filtered_projects.get(label_to_key.get(sel_label, ""), {}).get("client")]
                _compare_cli_raw = st.multiselect(
                    "Compare with clients (aggregated GROSS FEES, chart only)",
                    _other_clients, max_selections=2, key="ph_compare_cli_ms",
                    help="Each selected client's total GROSS FEES production is plotted as a line.",
                )
                _compare_cli_keys = _compare_cli_raw

        sel_proj           = label_to_key.get(sel_label) if sel_label else None
        _compare_proj_keys = _compare_proj_keys if st.session_state.get("ph_compare_on") else []
        _compare_cli_keys  = _compare_cli_keys  if st.session_state.get("ph_compare_cli") else []

        if sel_proj and sel_proj in filtered_projects:
            proj_info = filtered_projects[sel_proj]
            st.markdown(
                f"**Code:** {proj_info['code']} | **Client:** {proj_info['client']} "
                f"| **BU:** {proj_info['bu']} | **Currency:** {proj_info['currency']}"
            )

            _type_order = _COST_TYPE_ORDER
            _all_type_data  = proj_info.get("type_data", {})
            _available_types = list(_all_type_data.keys())
            _available_types_sorted = (
                [t for t in _type_order if t in _available_types]
                + [t for t in _available_types if t not in _type_order]
            )
            if not _available_types_sorted:
                _available_types_sorted = ["GROSS FEES"]
            # "Margin" aggregates every cost type into one net margin series
            _ALLTYPE_LABEL = "Margin (All Types)"
            _dropdown_options = [_ALLTYPE_LABEL] + _available_types_sorted

            ph_view_col, ph_type_col, ph_flip_col = st.columns([2, 2, 1])
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
                    _dropdown_options,
                    index=0,
                    key="ph_cost_type",
                    help=(
                        "'Margin (All Types)' sums every cost type to show net project margin. "
                        "GROSS FEES is the main project revenue. "
                        "Other types are cost components (typically negative)."
                    ),
                ) if ph_view == "Production (Ext. Prod.)" else "GROSS FEES"
            with ph_flip_col:
                ph_flip_neg = st.checkbox(
                    "Show costs upward",
                    key="ph_flip_neg",
                    help=(
                        "Multiplies negative values by −1 so cost-type bars point upward. "
                        "Only affects the chart; raw data below is unchanged."
                    ),
                ) if ph_view == "Production (Ext. Prod.)" else False

            # Build the active data: aggregate all types or pick the selected one
            if ph_cost_type == _ALLTYPE_LABEL:
                _active_margin_data: dict = {}
                for _td in _all_type_data.values():
                    for _k, _v in _td.items():
                        _active_margin_data[_k] = _active_margin_data.get(_k, 0) + _v
            else:
                _active_margin_data = _all_type_data.get(ph_cost_type, proj_info["margin_data"])

            # Cost type breakdown summary (always visible when multiple types exist)
            if len(_available_types_sorted) > 1 and ph_view == "Production (Ext. Prod.)":
                with st.expander("📊 All Cost Types — Breakdown Summary", expanded=False):
                    # Resolve current month label (e.g. "February 2026") from the current file
                    _cur_month_name_bk = _cur_sm_ph.split()[0]
                    _cur_month_lbl_bk  = None
                    if _cur_file_md:
                        _cur_month_lbl_bk = next(
                            (m["label"] for m in (_cur_file_md.get("month_cols") or [])
                             if m["label"].lower().startswith(_cur_month_name_bk.lower())),
                            None,
                        )

                    _cur_bo_yr  = (_cur_file_md.get("bo_year")  if _cur_file_md  else None) or "Year"
                    _prev_bo_yr = (_prev_file_md.get("bo_year") if _prev_file_md else None) or "Year"
                    _has_prev   = _prev_file_md is not None and _prev_sm_ph is not None

                    # Column headers
                    _col_cur_m  = f"{_cur_month_name_bk} Actual"
                    _col_prev_m = f"{_cur_month_name_bk} (Prev File — {_prev_sm_ph.split()[0] if _prev_sm_ph else '—'})"
                    _col_cur_bo = f"End {_cur_bo_yr} Forecast (Current)"
                    _col_prv_bo = f"End {_prev_bo_yr} Forecast (Prev)"

                    _breakdown_rows = []
                    for _ct in _available_types_sorted:
                        # Type-level data from current file
                        _td_cur = {}
                        if _cur_file_md and sel_proj in _cur_file_md.get("projects", {}):
                            _td_cur = _cur_file_md["projects"][sel_proj].get("type_data", {}).get(_ct, {})

                        # Type-level data from previous file
                        _td_prv = {}
                        if _has_prev and sel_proj in _prev_file_md.get("projects", {}):
                            _td_prv = _prev_file_md["projects"][sel_proj].get("type_data", {}).get(_ct, {})

                        # Current month value from current file (actual)
                        _cur_m_val  = _td_cur.get(_cur_month_lbl_bk, 0) if _cur_month_lbl_bk else 0
                        # Same month's value as recorded in the previous file (projected)
                        _prv_m_val  = _td_prv.get(_cur_month_lbl_bk, 0) if (_cur_month_lbl_bk and _td_prv) else None

                        # End-of-year forecast from each file
                        _cur_bo_key = next((k for k in _td_cur if k.startswith("BO (")), None)
                        _prv_bo_key = next((k for k in _td_prv if k.startswith("BO (")), None)
                        _cur_bo_val = _td_cur.get(_cur_bo_key, 0) if _cur_bo_key else 0
                        _prv_bo_val = _td_prv.get(_prv_bo_key, 0) if _prv_bo_key else None

                        # Month delta (Actual − Projected): positive = beat forecast
                        _m_delta = (_cur_m_val - _prv_m_val) if (_prv_m_val is not None) else None
                        # EoY delta (current forecast − previous forecast)
                        _bo_delta = (_cur_bo_val - _prv_bo_val) \
                                    if (_prv_bo_val is not None and _cur_bo_key) else None

                        def _delta_str(d):
                            if d is None:
                                return "—"
                            sign = "+" if d >= 0 else ""
                            return f"{sign}{human_tl(d)}"

                        _row = {
                            "Cost Type":  _ct,
                            _col_cur_m:   human_tl(_cur_m_val),
                            _col_cur_bo:  human_tl(_cur_bo_val) if _cur_bo_key else "—",
                        }
                        if _has_prev:
                            _row[_col_prev_m]       = human_tl(_prv_m_val) if _prv_m_val is not None else "—"
                            _row["Δ Month"]          = _delta_str(_m_delta)
                            _row[_col_prv_bo]        = human_tl(_prv_bo_val) if (_prv_bo_key and _prv_bo_val is not None) else "—"
                            _row["Δ EoY Forecast"]   = _delta_str(_bo_delta)
                        _breakdown_rows.append(_row)

                    # Reorder columns: Cost Type | Actual | Projected | Δ Month | EoY Cur | EoY Prev | Δ EoY
                    _cols_order = ["Cost Type", _col_cur_m, _col_cur_bo]
                    if _has_prev:
                        _cols_order = [
                            "Cost Type",
                            _col_cur_m, _col_prev_m, "Δ Month",
                            _col_cur_bo, _col_prv_bo, "Δ EoY Forecast",
                        ]
                    _df_bkdn = pd.DataFrame(_breakdown_rows)[_cols_order]

                    if _has_prev:
                        st.caption(
                            f"Actual = value from **{_cur_sm_ph}** file. "
                            f"Prev File = same month as projected in **{_prev_sm_ph}** file. "
                            f"EoY Forecast = end-of-year outlook from each respective file."
                        )
                    else:
                        st.caption(f"Values from **{_cur_sm_ph}** file. Load an additional month to see prior-file comparison.")
                    st.dataframe(_df_bkdn, hide_index=True, use_container_width=True)
                    st.download_button(
                        "⬇ Export Breakdown",
                        df_to_excel_bytes(_df_bkdn),
                        file_name=f"Breakdown_{proj_info['code']}_{_cur_sm_ph.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"bkdn_export_{sel_proj}",
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

                # Column and axis label: "Margin (TL)" when All Types combined, else "Production (TL)"
                _val_col   = "Margin (TL)" if ph_cost_type == _ALLTYPE_LABEL else "Production (TL)"
                _axis_lbl  = "Margin (TL)" if ph_cost_type == _ALLTYPE_LABEL else "Production (TL)"

                def _apply_flip(vals):
                    """Flip negatives to positive when 'Show costs upward' is on."""
                    if not ph_flip_neg:
                        return list(vals)
                    return [-v if v < 0 else v for v in vals]

                if margin_rows:
                    df_margin = pd.DataFrame([
                        {"Period": r["Period"], _val_col: r["Production (TL)"]}
                        for r in margin_rows
                    ])
                    df_margin["Label"] = df_margin[_val_col].apply(human_tl)

                    _flip_suffix = " ↑ costs shown as positive" if ph_flip_neg else ""
                    _type_suffix = (
                        " [All Types — Margin]" if ph_cost_type == _ALLTYPE_LABEL
                        else (f" [{ph_cost_type}]" if ph_cost_type != "GROSS FEES" else "")
                    )
                    _period_list    = [m["label"] for m in month_cols_sorted]
                    _txt_size       = 9 if len(_period_list) > 12 else 10
                    _line_colors    = [PRIMARY, "#10B981", "#F59E0B"]

                    if not _compare_proj_keys:
                        # Single-project line chart
                        _chart_y    = _apply_flip(df_margin[_val_col])
                        _chart_lbl  = [human_tl(v) for v in _chart_y]
                        fig_margin  = go.Figure()
                        fig_margin.add_trace(go.Scatter(
                            x=df_margin["Period"], y=_chart_y,
                            name=sel_label,
                            mode="lines+markers+text",
                            text=_chart_lbl,
                            textposition="top center",
                            textfont=dict(size=_txt_size, color=PRIMARY),
                            line=dict(color=PRIMARY, width=2),
                            marker=dict(size=8),
                            hovertemplate="%{x}: %{customdata}<extra></extra>",
                            customdata=_chart_lbl,
                        ))
                    else:
                        # Multi-project line chart
                        fig_margin  = go.Figure()
                        _all_cmp_keys = [sel_proj] + _compare_proj_keys
                        _all_cmp_lbls = [sel_label] + [
                            _display_label(filtered_projects[k]) for k in _compare_proj_keys
                            if k in filtered_projects
                        ]
                        for _pi, (_pk, _plbl) in enumerate(zip(_all_cmp_keys, _all_cmp_lbls)):
                            if _pk not in filtered_projects:
                                continue
                            _pinfo = filtered_projects[_pk]
                            if ph_cost_type == _ALLTYPE_LABEL:
                                _p_data: dict = {}
                                for _ptd in _pinfo.get("type_data", {}).values():
                                    for _pk2, _pv in _ptd.items():
                                        _p_data[_pk2] = _p_data.get(_pk2, 0) + _pv
                            else:
                                _p_data = _pinfo.get("type_data", {}).get(
                                    ph_cost_type, _pinfo["margin_data"]
                                )
                            _yvals = _apply_flip([_p_data.get(p, 0) for p in _period_list])
                            _ylbls = [human_tl(v) for v in _yvals]
                            fig_margin.add_trace(go.Scatter(
                                x=_period_list, y=_yvals,
                                name=_plbl,
                                mode="lines+markers+text",
                                text=_ylbls,
                                textposition="top center",
                                textfont=dict(size=_txt_size, color=_line_colors[_pi % len(_line_colors)]),
                                line=dict(color=_line_colors[_pi % len(_line_colors)], width=2),
                                marker=dict(size=7),
                                hovertemplate=f"<b>{_plbl}</b> — %{{x}}: %{{customdata}}<extra></extra>",
                                customdata=_ylbls,
                            ))

                    # ── Client comparison traces ──────────────────────────────
                    if _compare_cli_keys:
                        _cli_colors_extra = ["#8B5CF6", "#EC4899"]
                        for _ci, _cli in enumerate(_compare_cli_keys):
                            # Aggregate GROSS FEES for all projects of this client
                            _cli_data: dict = {}
                            for _, _pi_c in filtered_projects.items():
                                if _pi_c.get("client") != _cli:
                                    continue
                                _gf_c = _pi_c.get("type_data", {}).get("GROSS FEES", _pi_c.get("margin_data", {}))
                                for _kc, _vc in _gf_c.items():
                                    _cli_data[_kc] = _cli_data.get(_kc, 0) + _vc
                            _yvals_cli = _apply_flip([_cli_data.get(p, 0) for p in _period_list])
                            _ylbls_cli = [human_tl(v) for v in _yvals_cli]
                            _cc        = _cli_colors_extra[_ci % len(_cli_colors_extra)]
                            fig_margin.add_trace(go.Scatter(
                                x=_period_list, y=_yvals_cli,
                                name=f"{_cli} (client total)",
                                mode="lines+markers+text",
                                text=_ylbls_cli,
                                textposition="top center",
                                textfont=dict(size=_txt_size, color=_cc),
                                line=dict(color=_cc, width=2, dash="dot"),
                                marker=dict(size=7, symbol="diamond"),
                                hovertemplate=f"<b>{_cli}</b> — %{{x}}: %{{customdata}}<extra></extra>",
                                customdata=_ylbls_cli,
                            ))

                    _has_cmp = bool(_compare_proj_keys or _compare_cli_keys)
                    fig_margin.update_layout(
                        plot_bgcolor="white",
                        font=dict(family=_FONT, color=PRIMARY),
                        title=(
                            f"Comparison — {_axis_lbl}{_type_suffix}{_flip_suffix}"
                            if _has_cmp
                            else f"{sel_label}{_type_suffix} — {_axis_lbl}{_flip_suffix}"
                        ),
                        yaxis_title=_axis_lbl,
                        xaxis=dict(type="category", tickangle=-45),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02),
                        margin=dict(t=70, b=20),
                    )
                    st.plotly_chart(fig_margin, use_container_width=True,
                                    key=f"margin_chart_{sel_proj}_{ph_cost_type}")

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
                        {"Period": row["Period"], _val_col: human_tl(row[_val_col])}
                        for row in df_margin.to_dict("records")
                    ])
                    if bl_year:
                        df_display = pd.concat([df_display, pd.DataFrame([
                            {"Period": f"{bl_year} Production (TL)", _val_col: human_tl(bl_val)}
                        ])], ignore_index=True)
                    if bo_year:
                        df_display = pd.concat([df_display, pd.DataFrame([
                            {"Period": f"End {bo_year} Forecast (TL)", _val_col: human_tl(bo_val)}
                        ])], ignore_index=True)
                    st.dataframe(df_display, hide_index=True, use_container_width=True)
                    _export_label = "Margin" if ph_cost_type == _ALLTYPE_LABEL else "Production"
                    st.download_button(
                        f"⬇ Export {_export_label} to Excel",
                        df_to_excel_bytes(df_display),
                        file_name=f"{_export_label}_{proj_info['code']}.xlsx",
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
                        font=dict(family=_FONT, color=PRIMARY),
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
    sc_tabs = st.tabs(["BU Comparison", "Top 10 Clients", "Portfolio Mix", "What-If"])

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
                font=dict(family=_FONT, color=PRIMARY),
                showlegend=False,
                yaxis_title="kTL",
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
            _df_tier["Label"] = _df_tier["Value"].apply(human_k)
            _fig_tier.update_layout(
                plot_bgcolor="white",
                font=dict(family=_FONT, color=PRIMARY),
                showlegend=False,
                yaxis_title="kTL",
                margin=dict(t=60, b=20),
            )
            _fig_tier.update_traces(
                text=_df_tier["Label"],
                textposition="inside", insidetextanchor="middle",
                constraintext="inside",
            )
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
                    font=dict(family=_FONT, color=PRIMARY),
                    showlegend=False,
                    yaxis_title="kTL",
                    margin=dict(t=60, b=20),
                )
                _fig_ebit.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(_fig_ebit, use_container_width=True, key="sc_buebit")

    with sc_tabs[1]:
        st.markdown("#### Top Clients")
        _cv = st.radio(
            "Data source", ["Net Sales (Production)", "WIP"],
            horizontal=True, key="sc_client_src",
        )
        _formatter = human_tl
        _unit      = "TL"

        _cli_totals: dict[str, float] = {}
        if _cv == "Net Sales (Production)":
            # Aggregate GROSS FEES production YTD from Ext. Prod. for each client
            _ns_fp = EXCEL_FILES.get(_sm_latest)
            if _ns_fp and _ns_fp.exists():
                _ns_md = load_margin_data(str(_ns_fp))
                if _ns_md:
                    for _proj_info in _ns_md["projects"].values():
                        _c = _proj_info.get("client") or "Unknown"
                        if not _c or _c.lower() in ("nan", "none", ""):
                            continue
                        # Sum all monthly GROSS FEES production (YTD)
                        _gf = _proj_info.get("type_data", {}).get(
                            "GROSS FEES", _proj_info.get("margin_data", {})
                        )
                        _ytd = sum(
                            v for k, v in _gf.items()
                            if not any(k.startswith(p) for p in ("BL (", "BO (", "Next Year", "TD"))
                        )
                        if _ytd > 0:
                            _cli_totals[_c] = _cli_totals.get(_c, 0) + _ytd
            if not _cli_totals:
                st.info("NS production data unavailable — 'Ext. Prod.' sheet may be missing. Load more months or check the file.")
        else:
            _d = _d_latest
            if _d:
                for _p in _d["wip_projects"]:
                    _c = _p.get("client") or "Unknown"
                    if _p["wip_tl"] > 0:
                        _cli_totals[_c] = _cli_totals.get(_c, 0) + _p["wip_tl"]

        if _cli_totals:
            _top10 = sorted(_cli_totals.items(), key=lambda x: -x[1])[:_CLIENT_TOP_N]
            _df_top = pd.DataFrame(_top10, columns=["Client", "Value"])
            _df_top["Label"] = _df_top["Value"].apply(_formatter)

            _t1, _t2 = st.columns(2)
            with _t1:
                _fig_tbar = px.bar(
                    _df_top, x="Value", y="Client", orientation="h",
                    text="Label",
                    title=f"Top {_CLIENT_TOP_N} Clients — {_cv} ({_unit})",
                    color_discrete_sequence=[PRIMARY],
                    category_orders={"Client": _df_top["Client"].tolist()},
                )
                _fig_tbar.update_layout(
                    plot_bgcolor="white",
                    font=dict(family=_FONT, color=PRIMARY),
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
                "⬇ Export to Excel",
                df_to_excel_bytes(_df_top[["Client", "Value", "Label"]].rename(columns={"Label": _unit})),
                file_name=f"Top_Clients_{_cv.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sc_top10_export",
            )

    # ── Portfolio Mix ─────────────────────────────────────────────────────────
    with sc_tabs[2]:
        st.markdown("#### Portfolio Mix — Include / Exclude Business Units")
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
                    "NS — Selected BUs (kTL)",
                    fmt_ktl(_lr["NS Sel"]),
                    delta=f"−{fmt_ktl(_lr['NS Excl'])} excl  |  {_lr['NS Pct']:.0%} of total",
                    delta_color="off",
                )
                _mc2.metric(
                    "NS — Excluded Impact (kTL)",
                    fmt_ktl(_lr["NS Excl"]),
                    delta=f"{1 - _lr['NS Pct']:.0%} of total",
                    delta_color="off",
                )
                _mc3.metric(
                    "EBIT — Selected BUs (kTL)",
                    fmt_ktl(_lr["EBIT Sel"]),
                    delta=f"−{fmt_ktl(_lr['EBIT Excl'])} excl  |  {_lr['EBIT Pct']:.0%} of total",
                    delta_color="off",
                )
                _mc4.metric(
                    "EBIT — Excluded Impact (kTL)",
                    fmt_ktl(_lr["EBIT Excl"]),
                    delta=f"{1 - _lr['EBIT Pct']:.0%} of total",
                    delta_color="off",
                )

            # ── Charts ────────────────────────────────────────────────────────
            if _mix_chart_rows:
                _df_mix = pd.DataFrame(_mix_chart_rows)
                _df_mix["NS_label"]   = _df_mix["NS"].apply(fmt_ktl)
                _df_mix["EBIT_label"] = _df_mix["EBIT"].apply(fmt_ktl)

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
                    font=dict(family=_FONT, color=PRIMARY),
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
                    font=dict(family=_FONT, color=PRIMARY),
                    showlegend=False,
                    margin=dict(t=60, b=20),
                )
                _fig_mix_eb.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(_fig_mix_eb, use_container_width=True, key="sc_mix_ebit")

            # ── Summary table (all months) ────────────────────────────────────
            if _mix_summary:
                st.markdown("**Summary across all selected months**")
                _df_mix_tbl = pd.DataFrame([{
                    "Period":                r["Period"],
                    "NS Selected (kTL)":     fmt_ktl(r["NS Sel"]),
                    "NS Total (kTL)":        fmt_ktl(r["NS Total"]),
                    "NS Excl. Impact (kTL)": fmt_ktl(r["NS Excl"]),
                    "NS Coverage":           f"{r['NS Pct']:.0%}",
                    "EBIT Selected (kTL)":   fmt_ktl(r["EBIT Sel"]),
                    "EBIT Total (kTL)":      fmt_ktl(r["EBIT Total"]),
                    "EBIT Excl. (kTL)":      fmt_ktl(r["EBIT Excl"]),
                    "EBIT Coverage":         f"{r['EBIT Pct']:.0%}",
                } for r in _mix_summary])
                st.dataframe(_df_mix_tbl, hide_index=True, use_container_width=True)
                st.download_button(
                    "⬇ Export BU Mix to Excel",
                    df_to_excel_bytes(_df_mix_tbl),
                    file_name="BU_Mix_Scenario.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="sc_mix_export",
                )

    # ── What-If ───────────────────────────────────────────────────────────────
    with sc_tabs[3]:
        st.markdown("#### What-If Analysis")
        st.caption(
            "Adjust assumptions to see projected NS and EBIT impact on the latest selected month. "
            "Sliders apply to Contract NS only; EBIT impact uses the current EBIT % margin."
        )
        if not _d_latest:
            st.warning("No data available for the selected month.")
        else:
            _wi_ns   = _d_latest["ns"]
            _wi_ebit = _d_latest["ebit"]
            _wi_epct = _d_latest.get("ebit_pct")

            _wi_abbr  = _sm_latest.split()[0][:3]
            _wi_ns_i  = _cat_index(_wi_ns["cats"],   _wi_abbr)
            _wi_eb_i  = _cat_index(_wi_ebit["cats"], _wi_abbr)
            _wi_ep_i  = _cat_index(_wi_epct["cats"], _wi_abbr) if _wi_epct else -1

            _wi_ns_base  = monthly_val(_wi_ns["Contract"],       _wi_ns["cats"],   _wi_ns_i)
            _wi_eb_base  = monthly_val(_wi_ebit["Contract"],     _wi_ebit["cats"], _wi_eb_i)
            _wi_ebit_pct = safe_float(_wi_epct["EBIT_Pct"][_wi_ep_i]) if _wi_epct and _wi_ep_i != -1 else 0

            _wi_c1, _wi_c2 = st.columns(2)
            with _wi_c1:
                st.markdown("**Revenue Assumptions**")
                _wi_ns_growth = st.slider(
                    "Contract NS growth vs actual (%)",
                    min_value=-50, max_value=100, value=0, step=1, key="wi_ns_growth",
                )
                _wi_new_proj = st.number_input(
                    "Additional pipeline wins (kTL)",
                    min_value=0, value=0, step=1000, key="wi_new_proj",
                    help="Extra one-off order intake not in current actuals.",
                )
            with _wi_c2:
                st.markdown("**Cost/Margin Assumptions**")
                _wi_margin_adj = st.slider(
                    "EBIT margin adjustment (percentage points)",
                    min_value=-20, max_value=20, value=0, step=1, key="wi_margin_adj",
                    help="Shifts the EBIT % up or down from the current actual.",
                )

            # ── Projected values ──────────────────────────────────────────────
            _wi_ns_adj  = _wi_ns_base * (1 + _wi_ns_growth / 100) + _wi_new_proj
            _wi_pct_adj = _wi_ebit_pct + _wi_margin_adj / 100
            _wi_eb_adj  = _wi_ns_adj * _wi_pct_adj

            st.markdown("---")
            st.markdown(f"**Projected results for {_sm_latest}**")
            _r1, _r2, _r3, _r4 = st.columns(4)
            _r1.metric(
                "Actual NS (kTL)",      fmt_ktl(_wi_ns_base),
            )
            _r2.metric(
                "Projected NS (kTL)",   fmt_ktl(_wi_ns_adj),
                delta=fmt_ktl(_wi_ns_adj - _wi_ns_base),
            )
            _r3.metric(
                "Actual EBIT (kTL)",    fmt_ktl(_wi_eb_base),
            )
            _r4.metric(
                "Projected EBIT (kTL)", fmt_ktl(_wi_eb_adj),
                delta=fmt_ktl(_wi_eb_adj - _wi_eb_base),
            )
            _m1, _m2 = st.columns(2)
            _m1.metric(
                "Actual EBIT %",    f"{_wi_ebit_pct * 100:.1f}%",
            )
            _m2.metric(
                "Projected EBIT %", f"{_wi_pct_adj * 100:.1f}%",
                delta=f"{_wi_margin_adj:+d} pp",
            )

            # Waterfall chart: Actual NS → Growth → New wins → Projected NS
            # Build arrays dynamically — skip intermediate bars that are exactly zero
            _wf_growth_val = _wi_ns_base * _wi_ns_growth / 100
            _wf_new_val    = float(_wi_new_proj)
            _wf_x       = ["Actual NS"]
            _wf_y_plot  = [_wi_ns_base]
            _wf_measure = ["absolute"]
            _wf_text    = [fmt_ktl(_wi_ns_base)]
            if abs(_wf_growth_val) > 0.01:
                _wf_x.append(f"Growth {_wi_ns_growth:+d}%")
                _wf_y_plot.append(_wf_growth_val)
                _wf_measure.append("relative")
                _wf_text.append(fmt_ktl(_wf_growth_val))
            if abs(_wf_new_val) > 0.01:
                _wf_x.append("New Wins")
                _wf_y_plot.append(_wf_new_val)
                _wf_measure.append("relative")
                _wf_text.append(fmt_ktl(_wf_new_val))
            _wf_x.append("Projected NS")
            _wf_y_plot.append(0)  # go.Waterfall "total" measure ignores y — running sum is used
            _wf_measure.append("total")
            _wf_text.append(fmt_ktl(_wi_ns_adj))
            # Compute y-axis max with 20% headroom so outside labels never clip
            _wf_bar_max = max((_wi_ns_base, _wi_ns_adj), default=0)
            _wf_ymax    = _wf_bar_max * 1.20 if _wf_bar_max > 0 else None

            _fig_wf = go.Figure(go.Waterfall(
                x=_wf_x,
                y=_wf_y_plot,
                measure=_wf_measure,
                text=_wf_text,
                textposition="outside",
                cliponaxis=False,
                textfont=dict(size=12),
                connector=dict(line=dict(color="#9CA3AF")),
                increasing=dict(marker_color="#10B981"),
                decreasing=dict(marker_color="#EF4444"),
                totals=dict(marker_color=PRIMARY),
            ))
            _fig_wf.update_layout(
                plot_bgcolor="white",
                font=dict(family=_FONT, color=PRIMARY),
                yaxis=dict(title="kTL", range=[0, _wf_ymax] if _wf_ymax else None),
                margin=dict(t=60, b=20),
                height=360,
            )
            st.plotly_chart(_fig_wf, use_container_width=True, key="wi_waterfall")
            st.caption(
                "Waterfall shows the NS movement from actual to projected. "
                "EBIT projection = Projected NS × Adjusted EBIT %."
            )


# ── Tab Exec: Executive Summary ────────────────────────────────────────────────
with tab_exec:
    _render_exec_summary(loaded_data, selected_months, _multi_year)

# ── Tab 5: Project History ────────────────────────────────────────────────────
with tab5:
    _render_project_history(loaded_data, selected_months, global_bu_view, _multi_year)

# ── Tab 6: Scenarios ──────────────────────────────────────────────────────────
with tab6:
    _render_scenarios(loaded_data, selected_months, global_bu_view, _multi_year)
