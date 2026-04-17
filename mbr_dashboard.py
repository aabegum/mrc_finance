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


# ── Colors ─────────────────────────────────────────────────────────────────────
PRIMARY = "#1E3A8A"
BU_COLORS = {"ENG": "#0EA5E9", "MC": "#10B981", "T&SI": "#FAB611", "NUC": "#93C572"}
# Raw BU names from Excel → abbreviation used everywhere internally
BU_ABBREV = {
    "Engineering": "ENG",
    "MC": "MC",
    "T&SI": "T&SI",
    "TSI": "T&SI",
    "Nuclear": "NUC",
}

# ── Excel row indices and column indices loaded from config ────────────────────
_cfg = ConfigLoader(str(SCRIPT_DIR / "config"))
_r_cfg = _cfg.get("Excel_Mapping.Rows") or {}
_c_cfg = _cfg.get("Excel_Mapping.Columns") or {}
_o_cfg = _cfg.get("Excel_Mapping.Offsets") or {}

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
    "MON_NS_Base":             _o_cfg.get("MON_NS_Base", 3),
    "MON_EBIT_Base":           _o_cfg.get("MON_EBIT_Base", 3),
    "MON_BU_Base":             _o_cfg.get("MON_BU_Base", 2),
    "MON_OI_Base":             _o_cfg.get("MON_OI_Base", 2),
    "OI_Col_Base":             _o_cfg.get("OI_Col_Base", 6),
}

# Column indices kept separate from row indices
C = {
    "WIP_Col_Name":            _c_cfg.get("WIP_Col_Name", 2),
    "WIP_Col_Type":            _c_cfg.get("WIP_Col_Type", 3),
    "WIP_Col_Client":          _c_cfg.get("WIP_Col_Client", 4),
    "WIP_Col_BU":              _c_cfg.get("WIP_Col_BU", 10),
    "WIP_Col_OrigCurrency":    _c_cfg.get("WIP_Col_OrigCurrency", 11),
    "WIP_Col_TotalInvoiceOC":  _c_cfg.get("WIP_Col_TotalInvoiceOC", 93),
    "WIP_Col_TotalProdOC":     _c_cfg.get("WIP_Col_TotalProdOC", 94),
    "WIP_Col_WIP_TL":          _c_cfg.get("WIP_Col_WIP_TL", 98),
}

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


def row_vals(df, row, col_start, col_end):
    return [safe_float(df.iloc[row, c]) for c in range(col_start, col_end)]


def _cat_index(cats, label):
    """Case-insensitive index lookup into a category list. Returns -1 if not found."""
    label_l = str(label).lower()
    for i, c in enumerate(cats):
        if str(c).lower() == label_l:
            return i
    return -1


# ── File scanning ──────────────────────────────────────────────────────────────
def _parse_ver(v):
    """Parse version string like 'v1.10' into a comparable tuple (avoids lexicographic errors)."""
    return tuple(int(n) for n in re.findall(r"\d+", v))


def scan_excel_files():
    pattern = re.compile(
        r"(\d{6})_Project_Budget_Analysis_\w+_(v[\d.]+)\.xlsx", re.IGNORECASE
    )
    found = {}
    for f in SCRIPT_DIR.glob("*.xlsx"):
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
            # keep highest version when multiple files exist for same month
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

    # ── Row Offset Detection ──── Dynamic adjustment for file variations ────────
    # If the Excel file has extra or fewer rows compared to expected structure,
    # detect the actual categories row and calculate offset to apply globally.
    _oi_offset = 0
    _oi_offset_found = False
    expected_cat_row = R["Categories_OI_kTL"]
    for probe_offset in range(-5, 6):
        probe_row = expected_cat_row + probe_offset
        if 0 <= probe_row < len(df_oi):
            cell_val = str(df_oi.iloc[probe_row, 4]).strip().lower()
            if cell_val in ["order", "offer", "opportunity"] or "2025" in cell_val or "target" in cell_val:
                _oi_offset = probe_offset
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

    # OI projects — BU is normalized to abbreviation at load time so filters and
    # color maps work correctly everywhere without extra mapping steps downstream.
    # Extract month from filename to determine which column to use for monthly value
    month_idx = 0
    match = re.search(r"(\d{6})", str(filepath))
    if match:
        try:
            date_part = match.group(1)
            month_num = int(date_part[2:4])
            month_idx = month_num - 1  # Jan=0, Feb=1, etc.
        except Exception:
            month_idx = 0

    oi_col = O["OI_Col_Base"] + month_idx
    # Project rows run from row 2 up to (but not including) the first BU summary row.
    # Slide6_ENG is that first BU summary row, so this bound is intentionally exclusive.
    _oi_data_end = R["Slide6_ENG"] + _oi_offset
    oi_projects = []
    for ri in range(2, _oi_data_end):
        row = df_oi.iloc[ri]
        name = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        if not name or name.lower() == "nan":
            continue
        month_val = safe_float(row.iloc[oi_col])
        if month_val > 0:
            bu_raw = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
            oi_projects.append({
                "project": name,
                "client":  str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
                "bu":      BU_ABBREV.get(bu_raw, bu_raw),
                "value":   month_val,
            })
    oi_projects.sort(key=lambda x: x["value"], reverse=True)

    # WIP projects — BU also normalized; Nuclear excluded here per business rule.
    wip_projects = []
    for ri in range(R["WIP_DataStartRow"], len(df_wip)):
        row = df_wip.iloc[ri]
        proj_type = str(row.iloc[C["WIP_Col_Type"]]).strip() if pd.notna(row.iloc[C["WIP_Col_Type"]]) else ""
        if proj_type != "GROSS FEES":
            continue
        client = str(row.iloc[C["WIP_Col_Client"]]).strip() if pd.notna(row.iloc[C["WIP_Col_Client"]]) else ""
        if not client or "akkuyu" in client.lower():
            continue
        bu_raw = str(row.iloc[C["WIP_Col_BU"]]).strip() if pd.notna(row.iloc[C["WIP_Col_BU"]]) else ""
        bu = BU_ABBREV.get(bu_raw, bu_raw)
        if bu == "NUC":
            continue
        wip_tl = safe_float(row.iloc[C["WIP_Col_WIP_TL"]])
        if wip_tl < 1_000_000:
            continue
        wip_projects.append({
            "name":          str(row.iloc[C["WIP_Col_Name"]]).strip() if pd.notna(row.iloc[C["WIP_Col_Name"]]) else "",
            "client":        client,
            "bu":            bu,
            "orig_currency": str(row.iloc[C["WIP_Col_OrigCurrency"]]).strip() if pd.notna(row.iloc[C["WIP_Col_OrigCurrency"]]) else "",
            "inv_oc":        safe_float(row.iloc[C["WIP_Col_TotalInvoiceOC"]]),
            "prod_oc":       safe_float(row.iloc[C["WIP_Col_TotalProdOC"]]),
            "wip_tl":        wip_tl,
        })
    wip_projects.sort(key=lambda x: x["wip_tl"], reverse=True)

    # ── BU Specific Data Extraction ──────────────────────────────────────────
    c_ns_s   = _c_cfg.get("BU_NS_Start", 52)
    c_ns_e   = _c_cfg.get("BU_NS_End", 66)
    c_ebit_s = _c_cfg.get("BU_EBIT_Start", 51)
    c_ebit_e = _c_cfg.get("BU_EBIT_End", 66)

    cats_bu_ns   = [clean_label(df.iloc[_r_cfg.get("Categories_BU_NS", 14), c])   for c in range(c_ns_s, c_ns_e)]
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
            "Order": row_vals(df, _r_cfg.get(b_ord, 19 if b_key == "ENG" else (16 if b_key == "MC" else (22 if b_key == "T&SI" else 25))), c_ns_s, c_ns_e),
            "Offer": row_vals(df, _r_cfg.get(b_off, 20 if b_key == "ENG" else (17 if b_key == "MC" else (23 if b_key == "T&SI" else 26))), c_ns_s, c_ns_e),
            "Opp":   row_vals(df, _r_cfg.get(b_opp, 21 if b_key == "ENG" else (18 if b_key == "MC" else (24 if b_key == "T&SI" else 27))), c_ns_s, c_ns_e),
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
            "Total": row_vals(df, _r_cfg.get(b_row, 100), c_ebit_s, c_ebit_e),
        }

    return {
        "ns": {
            "cats":     cats_ns,
            "Contract": row_vals(df, R["Slide2_Contract"], 1, 17),
            "WP":       row_vals(df, R["Slide2_WP"],       1, 17),
            "WO":       row_vals(df, R["Slide2_WO"],       1, 17),
        },
        "ebit": {
            "cats":              cats_ebit,
            "Contract":          row_vals(df, R["Slide4_Contract"],       1, 16),
            "Contract+WP":       row_vals(df, R["Slide4_Contract_WP"],    1, 16),
            "Contract+WP+WO":    row_vals(df, R["Slide4_Contract_WP_WO"], 1, 16),
        },
        "oi": {
            "cats": cats_oi,
            "ENG":  [safe_float(df_oi.iloc[R["Slide6_ENG"] + _oi_offset, c]) for c in range(4, 18)],
            "MC":   [safe_float(df_oi.iloc[R["Slide6_MC"] + _oi_offset,  c]) for c in range(4, 18)],
            "T&SI": [safe_float(df_oi.iloc[R["Slide6_TSI"] + _oi_offset, c]) for c in range(4, 18)],
            "NUC":  [safe_float(df_oi.iloc[R["Slide6_NUC"] + _oi_offset, c]) for c in range(4, 18)],
        },
        "oi_projects":  oi_projects,
        "wip_projects": wip_projects,
        "bu_ns":        bu_ns,
        "bu_ebit":      bu_ebit,
    }


# ── Data fetcher (cache-safe: takes a tuple, never mutated) ────────────────────
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

        df_ext = pd.read_excel(xl, sheet_name="Ext. Prod.", header=None)

        # Parse month headers from first 2 rows
        # Row 0 has month names, Row 1 has dates (e.g., 31.01.2026)
        # Columns are: Project Code (A), Project Name (B), Type (C), Client (D), BU (E),
        # then month groups with PRODUCTION (TL), PRODUCTION €, PRODUCTION OC, AVG TRL/OC for each month

        month_cols = {}  # Month name -> {"col_idx": idx, "date": date_str}
        if len(df_ext) > 1:
            row1_data = df_ext.iloc[0]
            row2_data = df_ext.iloc[1]

            current_month = None
            for col in range(6, len(row1_data)):
                val = str(row1_data.iloc[col]).strip()
                if val and val.upper() not in ["PRODUCTION OC", "PRODUCTION", "PRODUCTION €", "AVG TRL/OC", "NAN"]:
                    if "JANUARY" in val.upper() or "FEBRUARY" in val.upper() or "MARCH" in val.upper() or \
                       "APRIL" in val.upper() or "MAY" in val.upper() or "JUNE" in val.upper() or \
                       "JULY" in val.upper() or "AUGUST" in val.upper() or "SEPTEMBER" in val.upper() or \
                       "OCTOBER" in val.upper() or "NOVEMBER" in val.upper() or "DECEMBER" in val.upper():
                        current_month = val
                elif current_month:
                    col_header = val.upper()
                    # Look for "PRODUCTION" (TL) column, not "PRODUCTION €"
                    if col_header == "PRODUCTION" and current_month not in month_cols:
                        date_str = str(row2_data.iloc[col]).strip() if col < len(row2_data) else ""
                        month_cols[current_month] = {"col_idx": col, "date": date_str}

        # Find BL and BO columns (2026 PRODUCTION and END OF 2026 TOTAL)
        bl_col = None
        bo_col = None
        for col in range(len(df_ext.iloc[0])):
            val = str(df_ext.iloc[0, col]).strip().upper()
            if "2026 PRODUCTION" in val and "END OF" not in val and bl_col is None:
                bl_col = col
            elif "END OF 2026 TOTAL PRODUCTION" in val and bo_col is None:
                bo_col = col

        # Extract projects and margin data
        _EXT_ROW_LIMIT = 500
        if len(df_ext) > _EXT_ROW_LIMIT:
            st.warning(
                f"'Ext. Prod.' sheet has {len(df_ext)} rows but only the first "
                f"{_EXT_ROW_LIMIT} are read. Projects beyond row {_EXT_ROW_LIMIT + 3} "
                f"will be missing from the margin history view."
            )
        projects = {}  # project_name -> {"client": str, "bu": str, "code": str, "margin_data": {month: value}}
        for ri in range(3, min(len(df_ext), _EXT_ROW_LIMIT)):
            row = df_ext.iloc[ri]
            proj_code = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            proj_name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""

            if not proj_code or not proj_name or proj_name.lower() == "nan":
                continue

            client = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
            bu_raw = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
            bu = BU_ABBREV.get(bu_raw, bu_raw)

            margin_vals = {}
            for month_name, month_info in month_cols.items():
                col_idx = month_info["col_idx"]
                val = safe_float(row.iloc[col_idx]) if col_idx < len(row) else 0
                margin_vals[month_name] = val

            # Add BL and BO
            bl_val = safe_float(row.iloc[bl_col]) if bl_col and bl_col < len(row) else 0
            bo_val = safe_float(row.iloc[bo_col]) if bo_col and bo_col < len(row) else 0
            margin_vals["BL (2026 Prod)"] = bl_val
            margin_vals["BO (End 2026 Total)"] = bo_val

            projects[proj_name] = {
                "code": proj_code,
                "client": client,
                "bu": bu,
                "margin_data": margin_vals,
            }

        return {"projects": projects, "month_cols": month_cols, "bl_col": bl_col, "bo_col": bo_col}
    except Exception as e:
        st.error(f"Error loading margin data: {e}")
        return None


# ── Chart builders ─────────────────────────────────────────────────────────────
def chart_stacked(data, title, keys, color_palette=None):
    n = min(len(data["cats"]), min(len(data[k]) for k in keys))
    df = pd.DataFrame({"Category": data["cats"][:n]})
    for k in keys:
        df[k] = data[k][:n]
    df_m = df.melt("Category", var_name="Tier", value_name="Value")

    if color_palette:
        colors = color_palette
    else:
        colors = (
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
        margin=dict(t=60, b=20),
        # Hide labels that are too small to fit inside their segment
        uniformtext_minsize=8,
        uniformtext_mode="hide",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    fig.update_traces(texttemplate="%{y:,.0f}", textposition="inside", insidetextanchor="middle")
    return fig


def chart_grouped(data, title, keys):
    n = min(len(data["cats"]), min(len(data[k]) for k in keys))
    df = pd.DataFrame({"Category": data["cats"][:n]})
    for k in keys:
        df[k] = data[k][:n]
    df_m = df.melt("Category", var_name="Tier", value_name="Value")
    fig = px.bar(
        df_m, x="Category", y="Value", color="Tier", title=title,
        color_discrete_sequence=["#0EA5E9", "#10B981", "#F59E0B"],
        barmode="group", hover_data={"Value": ":,.0f"},
    )
    fig.update_layout(
        plot_bgcolor="white",
        font=dict(family="Segoe UI", color=PRIMARY),
        xaxis_tickangle=-45,
        margin=dict(t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    fig.update_traces(texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False)
    return fig


def chart_projects(projects, title):
    if not projects:
        return None
    df = pd.DataFrame(projects[:30])
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
    df = pd.DataFrame(wip_list[:30])
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
    all_years = sorted(set(k.split()[1] for k in EXCEL_FILES))
    selected_year = st.selectbox("Year", all_years, index=len(all_years) - 1)

    months_for_year = sorted(
        [k for k in EXCEL_FILES if k.endswith(selected_year)],
        key=lambda x: MONTH_ORDER.index(x.split()[0]),
    )

    compare_mode = st.checkbox("Compare months", value=False)
    year_compare_mode = st.checkbox("Compare years", value=False)

    if compare_mode and year_compare_mode:
        st.warning(
            "Both comparison modes are active. "
            "**Month comparison takes priority** — year comparison is disabled."
        )
        year_compare_mode = False

    if year_compare_mode:
        compare_year_options = [y for y in all_years if y != selected_year]
        if compare_year_options:
            comparison_year = st.selectbox("Compare with:", compare_year_options)
            months_for_comparison = sorted(
                [k for k in EXCEL_FILES if k.endswith(comparison_year)],
                key=lambda x: MONTH_ORDER.index(x.split()[0]),
            )
        else:
            comparison_year = None
            months_for_comparison = []
    else:
        comparison_year = None
        months_for_comparison = []

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
    else:
        single_month = st.selectbox(
            "Month", months_for_year,
            index=len(months_for_year) - 1,
            format_func=lambda x: x.split()[0],
        )
        selected_months = [single_month] if single_month else []

        # If year comparison enabled, also add same month from comparison year
        if year_compare_mode and comparison_year and months_for_comparison:
            comp_month_name = selected_months[0].split()[0] if selected_months else None
            if comp_month_name:
                comp_months = [m for m in months_for_comparison if m.startswith(comp_month_name)]
                if comp_months:
                    selected_months.append(comp_months[0])

if not selected_months:
    st.warning("Please select at least one month from the sidebar.")
    st.stop()

with st.spinner("Loading data..."):
    loaded_data = fetch_all_data(tuple(selected_months))

# Build project/client sets for filter widgets (skip None month data)
all_projects = set()
all_clients = set()
for data in loaded_data.values():
    if not data:
        continue
    for p in data.get("oi_projects", []):
        if p.get("project"): all_projects.add(p["project"])
        if p.get("client"):  all_clients.add(p["client"])
    for p in data.get("wip_projects", []):
        if p.get("name"):   all_projects.add(p["name"])
        if p.get("client"): all_clients.add(p["client"])

# ── Sidebar 2: Filters ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("### Filters")

    bu_options = ["ENG", "MC", "T&SI", "NUC"]
    bu_filter      = st.multiselect("Business Unit", bu_options, default=bu_options)
    project_search = st.multiselect("Search Project", sorted(all_projects))
    client_search  = st.multiselect("Search Client",  sorted(all_clients))

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
    else:
        st.error("Generator encountered an error — see details below.")
    if result.stdout:
        st.code(result.stdout, language="")
    if result.stderr:
        with st.expander("Error details"):
            st.code(result.stderr, language="")
    st.markdown("---")


# ── Filter helpers ─────────────────────────────────────────────────────────────
# Both OI and WIP projects now store BU as normalized abbreviations, so we
# compare directly against bu_filter without any extra mapping.
def filter_oi(projects):
    out = projects
    if project_search:
        out = [p for p in out if p["project"] in project_search]
    if client_search:
        out = [p for p in out if p["client"] in client_search]
    if bu_filter:
        out = [p for p in out if p["bu"] in bu_filter]
    return out


def filter_wip(wip):
    out = wip
    if project_search:
        out = [p for p in out if p["name"] in project_search]
    if client_search:
        out = [p for p in out if p["client"] in client_search]
    if bu_filter:
        out = [p for p in out if p["bu"] in bu_filter]
    return out


# ── Render helpers ─────────────────────────────────────────────────────────────
def render_ns_view(sm, data, prev_sm, prev_data, is_total, bu_name):
    if is_total:
        ns    = data["ns"]
        k_list = ["Contract", "WP", "WO"]
        p_ns  = prev_data["ns"] if prev_data else None
    else:
        ns    = data["bu_ns"][bu_name]
        k_list = ["Order", "Offer", "Opp"]
        p_ns  = prev_data["bu_ns"][bu_name] if prev_data else None

    month_abbr = sm.split()[0][:3]
    ns_idx = _cat_index(ns["cats"], month_abbr)

    totals = {k: ns[k][ns_idx] for k in k_list} if ns_idx != -1 else {k: 0 for k in k_list}
    totals["Total"] = sum(totals.values())

    prev_totals = {}
    if p_ns and prev_sm:
        p_abbr = prev_sm.split()[0][:3]
        p_idx  = _cat_index(p_ns["cats"], p_abbr)
        if p_idx != -1:
            prev_totals = {k: p_ns[k][p_idx] for k in k_list}
            prev_totals["Total"] = sum(prev_totals.values())

    all_metric_keys = k_list + ["Total"]
    mcols = st.columns(len(all_metric_keys))
    for idx_m, k in enumerate(all_metric_keys):
        val   = totals[k]
        delta = val - prev_totals.get(k, 0) if prev_totals else None
        mcols[idx_m].metric(k, human_k(val), delta=human_k(delta) if delta is not None else None)

    palette = None if is_total else ["#0EA5E9", "#10B981", "#CBD5E1"]
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

    mcols = st.columns(min(3, len(k_list)))
    for idx_m, k in enumerate(k_list):
        val   = ebit[k][ebit_idx] if ebit_idx != -1 else 0
        delta = val - p_ebit[k][prev_ebit_idx] if p_ebit and prev_ebit_idx != -1 else None
        mcols[idx_m % len(mcols)].metric(k, human_k(val), delta=human_k(delta) if delta is not None else None)

    if is_total:
        st.plotly_chart(
            chart_grouped(ebit, "", k_list),
            use_container_width=True, key=f"ebit_{bu_name}_{sm}",
        )
    else:
        # Single-series BU view — chart_grouped is semantically correct here;
        # chart_stacked with one series is just a plain bar with no stacking meaning.
        st.plotly_chart(
            chart_grouped(ebit, "", k_list),
            use_container_width=True, key=f"ebit_{bu_name}_{sm}",
        )


def _month_cols_iter(months, data_map):
    """Yield (col, sm, data, prev_sm, prev_data) for each month column."""
    cols = st.columns(len(months))
    for i, sm in enumerate(months):
        data = data_map.get(sm)
        if not data:
            continue
        prev_sm   = months[i - 1] if i > 0 else None
        prev_data = data_map.get(prev_sm) if prev_sm else None
        yield cols[i], sm, data, prev_sm, prev_data


# ── Main UI ───────────────────────────────────────────────────────────────────
global_bu_view = st.radio(
    "Business Unit View", ["Company Total", "ENG", "MC", "T&SI", "NUC"], horizontal=True,
    help="Applies to all tabs. 'Company Total' shows aggregated figures; selecting a BU filters to that unit's data.",
)
st.markdown("---")
is_total = global_bu_view == "Company Total"

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Net Sales", "EBIT", "Order Intake", "WIP", "Project History"])

# ── Tab 1: Net Sales ──────────────────────────────────────────────────────────
with tab1:
    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        with col:
            st.subheader(f"{sm.split()[0]} Net Sales")
            render_ns_view(sm, data, prev_sm, prev_data, is_total, global_bu_view)

# ── Tab 2: EBIT ───────────────────────────────────────────────────────────────
with tab2:
    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        with col:
            st.subheader(f"{sm.split()[0]} EBIT")
            render_ebit_view(sm, data, prev_sm, prev_data, is_total, global_bu_view)

# ── Tab 3: Order Intake ───────────────────────────────────────────────────────
with tab3:
    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        with col:
            st.subheader(f"{sm.split()[0]} Order Intake")
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

                mcols = st.columns(2)
                for idx_m, k in enumerate(keys_oi):
                    val   = oi[k][oi_idx] if oi_idx != -1 else 0
                    delta = val - prev_oi[k][prev_oi_idx] if prev_oi and prev_oi_idx != -1 else None
                    mcols[idx_m % 2].metric(k, human_k(val), delta=human_k(delta) if delta is not None else None)

                bg_oi = filter_oi(data["oi_projects"])
            else:
                bg_oi     = [p for p in filter_oi(data["oi_projects"]) if p["bu"] == global_bu_view]
                total_oi  = sum(p["value"] for p in bg_oi)
                delta_oi  = None
                if prev_data:
                    p_bg_oi  = [p for p in filter_oi(prev_data["oi_projects"]) if p["bu"] == global_bu_view]
                    delta_oi = total_oi - sum(p["value"] for p in p_bg_oi)
                st.metric(
                    f"Total Order Intake ({len(bg_oi)} projs)",
                    human_k(total_oi),
                    delta=human_k(delta_oi) if delta_oi is not None else None,
                )
                fig_bu_oi = chart_projects(bg_oi, "")
                if fig_bu_oi:
                    st.plotly_chart(fig_bu_oi, use_container_width=True, key=f"oi_bu_fig_{global_bu_view}_{sm}")

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
                    use_container_width=True, key=f"oi_df_{global_bu_view}_{sm}",
                )
            else:
                st.info("No projects match the current filters.")

# ── Tab 4: WIP ────────────────────────────────────────────────────────────────
with tab4:
    for col, sm, data, prev_sm, prev_data in _month_cols_iter(selected_months, loaded_data):
        with col:
            st.subheader(f"{sm.split()[0]} WIP")
            if is_total:
                wip       = filter_wip(data["wip_projects"])
                total_wip = sum(p["wip_tl"] for p in wip)
                delta_wip = None
                if prev_data:
                    prev_total = sum(p["wip_tl"] for p in filter_wip(prev_data["wip_projects"]))
                    delta_wip  = human_tl(total_wip - prev_total)
            else:
                wip       = [p for p in filter_wip(data["wip_projects"]) if p["bu"] == global_bu_view]
                total_wip = sum(p["wip_tl"] for p in wip)
                delta_wip = None
                if prev_data:
                    prev_wip  = [p for p in filter_wip(prev_data["wip_projects"]) if p["bu"] == global_bu_view]
                    delta_wip = human_tl(total_wip - sum(p["wip_tl"] for p in prev_wip))

            st.metric(f"Total WIP ({len(wip)} projs)", human_tl(total_wip), delta=delta_wip)

            fig = chart_wip(wip, "")
            if fig:
                st.plotly_chart(fig, use_container_width=True, key=f"wip_fig_{global_bu_view}_{sm}")
            if wip:
                df_disp   = pd.DataFrame(wip)
                df_disp["WIP TL"] = df_disp["wip_tl"].apply(human_tl)
                cols_show = (
                    ["name", "client", "bu", "WIP TL"]
                    if is_total
                    else ["name", "client", "WIP TL"]
                )
                st.dataframe(
                    df_disp[cols_show], hide_index=True,
                    use_container_width=True, key=f"wip_df_{global_bu_view}_{sm}",
                )

# ── Tab 5: Margin History ────────────────────────────────────────────────────────
with tab5:
    st.markdown("### Margin History")
    st.caption(f"Track monthly production margins and projections from 'Ext. Prod.' sheet.")

    if not selected_months:
        st.warning("Please select at least one month to view margin data.")
    else:
        # Load margin data from the first selected month's file
        # (uses a distinct variable to avoid shadowing the PPTX sel_file in the sidebar)
        margin_file = EXCEL_FILES.get(selected_months[0])
        if margin_file and margin_file.exists():
            with st.spinner("Loading margin data…"):
                margin_data = load_margin_data(str(margin_file))

            if margin_data and margin_data.get("projects"):
                projects = margin_data["projects"]
                month_cols = margin_data["month_cols"]

                # Project selector with filtering
                all_projs = sorted(projects.keys())
                sel_proj = st.selectbox("Select Project", all_projs, key="margin_proj")

                def _margin_month_sort(k):
                    """Sort month keys like 'JANUARY 2026' chronologically via MONTH_ORDER."""
                    for i, m in enumerate(MONTH_ORDER):
                        if m.upper() in k.upper():
                            return i
                    return 99

                if sel_proj and sel_proj in projects:
                    proj_info = projects[sel_proj]
                    st.markdown(f"**Code:** {proj_info['code']} | **Client:** {proj_info['client']} | **BU:** {proj_info['bu']}")

                    # Display monthly margins as line chart
                    rows = []
                    for mo_key in sorted(month_cols.keys(), key=_margin_month_sort):
                        val = proj_info['margin_data'].get(mo_key, 0)
                        rows.append({"Month": mo_key[:3], "Production (TL)": val, "Full": mo_key})

                    if rows:
                        df_margin = pd.DataFrame(rows)
                        df_margin["Label"] = df_margin["Production (TL)"].apply(human_tl)

                        fig_margin = px.line(
                            df_margin, x="Full", y="Production (TL)", markers=True,
                            title=f"{sel_proj} — Monthly Production (TL)",
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
                        st.plotly_chart(fig_margin, use_container_width=True)

                    # Year-end projections
                    bl_val = proj_info['margin_data'].get("BL (2026 Prod)", 0)
                    bo_val = proj_info['margin_data'].get("BO (End 2026 Total)", 0)

                    mcol1, mcol2 = st.columns(2)
                    mcol1.metric("2026 Production (TL)", human_tl(bl_val))
                    mcol2.metric("End 2026 Total (TL)", human_tl(bo_val))

                    # Data table
                    df_display = pd.DataFrame([
                        {
                            "Period": mo_key,
                            "Production (TL)": human_tl(proj_info['margin_data'].get(mo_key, 0))
                        }
                        for mo_key in sorted(month_cols.keys(), key=_margin_month_sort)
                    ])
                    df_display = pd.concat([df_display, pd.DataFrame([
                        {"Period": "2026 Production (TL)", "Production (TL)": human_tl(bl_val)},
                        {"Period": "End 2026 Total (TL)", "Production (TL)": human_tl(bo_val)},
                    ])], ignore_index=True)

                    st.dataframe(df_display, hide_index=True, use_container_width=True)
                else:
                    st.info("Please select a project to view margin history.")
            else:
                st.error("Could not load margin data from 'Ext. Prod.' sheet. Verify sheet exists and has correct structure.")
        else:
            st.warning("No Excel file found for the selected month (Tab 5).")
