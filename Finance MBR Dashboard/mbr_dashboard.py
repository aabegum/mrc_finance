import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="MRC Business Review", layout="wide", page_icon="📊")

SCRIPT_DIR = Path(__file__).parent

# ── Colors ─────────────────────────────────────────────────────────────────────
PRIMARY = "#1E3A8A"
BU_COLORS = {"ENG": "#0EA5E9", "MC": "#10B981", "T&SI": "#FAB611", "NUC": "#93C572"}
# Raw BU names from Excel → abbreviation used for color lookup
BU_ABBREV = {
    "Engineering": "ENG",
    "MC": "MC",
    "T&SI": "T&SI",
    "TSI": "T&SI",
    "Nuclear": "NUC",
}

# ── Excel row indices (0-based, header=None reads) ─────────────────────────────
R = {
    "Categories_NS_kTL":       14,
    "Categories_EBIT_kTL":     36,
    "Categories_OI_kTL":      543,
    "Slide2_Contract":         15,
    "Slide2_WP":               16,
    "Slide2_WO":               17,
    "Slide4_Contract":         37,
    "Slide4_Contract_WP":      38,
    "Slide4_Contract_WP_WO":   39,
    "Slide6_ENG":             544,
    "Slide6_MC":              545,
    "Slide6_TSI":             546,
    "Slide6_NUC":             547,
    "WIP_DataStartRow":         3,
    "WIP_Col_Name":             2,
    "WIP_Col_Type":             3,
    "WIP_Col_Client":           4,
    "WIP_Col_BU":              10,
    "WIP_Col_OrigCurrency":    11,
    "WIP_Col_TotalInvoiceOC":  93,
    "WIP_Col_TotalProdOC":     94,
    "WIP_Col_WIP_TL":          98,
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


# ── File scanning ──────────────────────────────────────────────────────────────
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
            if label not in found or ver > found[label][1]:
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

    cats_ns   = [clean_label(df.iloc[R["Categories_NS_kTL"], c])    for c in range(1, 17)]
    cats_ebit = [clean_label(df.iloc[R["Categories_EBIT_kTL"], c])  for c in range(1, 17)]
    cats_oi   = [clean_label(df_oi.iloc[R["Categories_OI_kTL"], c]) for c in range(4, 18)]

    # OI projects — scan every project row and sum all monthly columns
    oi_projects = []
    for ri in range(2, R["Slide6_ENG"]):
        row = df_oi.iloc[ri]
        name = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        if not name or name.lower() == "nan":
            continue
        total_val = sum(safe_float(row.iloc[c]) for c in range(4, 18))
        if total_val > 0:
            oi_projects.append({
                "project": name,
                "client":  str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
                "bu":      str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else "",
                "value":   total_val,
            })
    oi_projects.sort(key=lambda x: x["value"], reverse=True)

    # WIP projects
    wip_projects = []
    for ri in range(R["WIP_DataStartRow"], len(df_wip)):
        row = df_wip.iloc[ri]
        proj_type = str(row.iloc[R["WIP_Col_Type"]]).strip() if pd.notna(row.iloc[R["WIP_Col_Type"]]) else ""
        if proj_type != "GROSS FEES":
            continue
        client = str(row.iloc[R["WIP_Col_Client"]]).strip() if pd.notna(row.iloc[R["WIP_Col_Client"]]) else ""
        if not client or "akkuyu" in client.lower():
            continue
        bu = str(row.iloc[R["WIP_Col_BU"]]).strip() if pd.notna(row.iloc[R["WIP_Col_BU"]]) else ""
        if bu == "Nuclear":
            continue
        wip_tl = safe_float(row.iloc[R["WIP_Col_WIP_TL"]])
        if wip_tl < 1_000_000:
            continue
        wip_projects.append({
            "name":          str(row.iloc[R["WIP_Col_Name"]]).strip() if pd.notna(row.iloc[R["WIP_Col_Name"]]) else "",
            "client":        client,
            "bu":            bu,
            "orig_currency": str(row.iloc[R["WIP_Col_OrigCurrency"]]).strip() if pd.notna(row.iloc[R["WIP_Col_OrigCurrency"]]) else "",
            "inv_oc":        safe_float(row.iloc[R["WIP_Col_TotalInvoiceOC"]]),
            "prod_oc":       safe_float(row.iloc[R["WIP_Col_TotalProdOC"]]),
            "wip_tl":        wip_tl,
        })
    wip_projects.sort(key=lambda x: x["wip_tl"], reverse=True)

    return {
        "ns": {
            "cats": cats_ns,
            "Contract": row_vals(df, R["Slide2_Contract"], 1, 17),
            "WP":       row_vals(df, R["Slide2_WP"],       1, 17),
            "WO":       row_vals(df, R["Slide2_WO"],       1, 17),
        },
        "ebit": {
            "cats": cats_ebit,
            "Contract":         row_vals(df, R["Slide4_Contract"],       1, 16),
            "Contract+WP":      row_vals(df, R["Slide4_Contract_WP"],    1, 16),
            "Contract+WP+WO":   row_vals(df, R["Slide4_Contract_WP_WO"], 1, 16),
        },
        "oi": {
            "cats": cats_oi,
            "ENG":  [safe_float(df_oi.iloc[R["Slide6_ENG"], c]) for c in range(4, 18)],
            "MC":   [safe_float(df_oi.iloc[R["Slide6_MC"],  c]) for c in range(4, 18)],
            "T&SI": [safe_float(df_oi.iloc[R["Slide6_TSI"], c]) for c in range(4, 18)],
            "NUC":  [safe_float(df_oi.iloc[R["Slide6_NUC"], c]) for c in range(4, 18)],
        },
        "oi_projects":  oi_projects,
        "wip_projects": wip_projects,
    }


# ── Chart builders ─────────────────────────────────────────────────────────────
def chart_stacked(data, title, keys):
    n = min(len(data["cats"]), min(len(data[k]) for k in keys))
    df = pd.DataFrame({"Category": data["cats"][:n]})
    for k in keys:
        df[k] = data[k][:n]
    df_m = df.melt("Category", var_name="Tier", value_name="Value")
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
        xaxis_tickangle=30,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    fig.update_traces(texttemplate="%{y:,.0f}", textposition="inside")
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
        xaxis_tickangle=30,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    fig.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
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
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        height=max(400, len(df) * 22),
    )
    fig.update_traces(textposition="outside")
    return fig


def chart_wip(wip_list, title):
    if not wip_list:
        return None
    df = pd.DataFrame(wip_list[:30])
    df["label"] = df["wip_tl"].apply(human_tl)
    df["bu_key"] = df["bu"].apply(lambda b: BU_ABBREV.get(b, b))
    fig = px.bar(
        df, x="wip_tl", y="name", color="bu_key", orientation="h",
        title=title, text="label",
        color_discrete_map=BU_COLORS,
        category_orders={"name": df["name"].tolist()},
        hover_data={"client": True, "orig_currency": True, "wip_tl": ":,.0f"},
    )
    fig.update_layout(
        plot_bgcolor="white",
        font=dict(family="Segoe UI", color=PRIMARY),
        yaxis_title="", xaxis_title="TL",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        height=max(400, len(df) * 22),
    )
    fig.update_traces(textposition="outside")
    return fig


# ── App ────────────────────────────────────────────────────────────────────────
st.title("MRC Business Review Dashboard")

if not EXCEL_FILES:
    st.error(
        "No Excel files found. Place `*_Project_Budget_Analysis_*_v*.xlsx` files "
        "in the same folder as this script."
    )
    st.stop()

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Period")
    all_years = sorted(set(k.split()[1] for k in EXCEL_FILES))
    selected_year = st.selectbox("Year", all_years, index=len(all_years) - 1)

    months_for_year = sorted(
        [k for k in EXCEL_FILES if k.endswith(selected_year)],
        key=lambda x: MONTH_ORDER.index(x.split()[0]),
    )
    selected_month = st.selectbox(
        "Month", months_for_year, index=len(months_for_year) - 1
    )

    st.markdown("---")
    st.markdown("### Filters")

    bu_options = ["ENG", "MC", "T&SI", "NUC"]
    bu_filter = st.multiselect("Business Unit", bu_options, default=bu_options)
    project_search = st.text_input("Search Project", "")
    client_search  = st.text_input("Search Client",  "")

    st.markdown("---")
    st.markdown("### Generate Report")
    sel_file = EXCEL_FILES.get(selected_month)
    st.caption(f"File: {sel_file.name if sel_file else '—'}")
    run_btn = st.button("▶ Generate PPTX", type="primary", use_container_width=True)

# ── Load data ──────────────────────────────────────────────────────────────────
fpath = EXCEL_FILES.get(selected_month)
if not fpath or not fpath.exists():
    st.error(f"File not found for {selected_month}")
    st.stop()

with st.spinner("Loading data..."):
    data = load_month(str(fpath))

if data is None:
    st.stop()

# ── Run report ─────────────────────────────────────────────────────────────────
if run_btn:
    st.markdown("---")
    with st.spinner(f"Generating PPTX for {selected_month}…"):
        result = subprocess.run(
            [sys.executable, str(SCRIPT_DIR / "generate_full_mbr_stacked.py"), str(fpath)],
            capture_output=True, text=True, cwd=str(SCRIPT_DIR),
        )
    if result.returncode == 0:
        st.success(f"Report generated successfully for {selected_month}!")
    else:
        st.error("Generator encountered an error — see details below.")
    if result.stdout:
        st.code(result.stdout, language="")
    if result.stderr:
        with st.expander("Error details"):
            st.code(result.stderr, language="")
    st.markdown("---")

# ── Filter helpers ─────────────────────────────────────────────────────────────
def filter_oi(projects):
    out = projects
    if project_search:
        out = [p for p in out if project_search.lower() in p["project"].lower()]
    if client_search:
        out = [p for p in out if client_search.lower() in p["client"].lower()]
    if bu_filter:
        out = [p for p in out if p["bu"] in bu_filter]
    return out


def filter_wip(wip):
    out = wip
    if project_search:
        out = [p for p in out if project_search.lower() in p["name"].lower()]
    if client_search:
        out = [p for p in out if client_search.lower() in p["client"].lower()]
    if bu_filter:
        out = [p for p in out if BU_ABBREV.get(p["bu"], p["bu"]) in bu_filter]
    return out


# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["Net Sales", "EBIT", "Order Intake", "WIP"])

with tab1:
    st.subheader(f"Net Sales — {selected_month}")
    ns = data["ns"]
    totals = {k: sum(ns[k]) for k in ["Contract", "WP", "WO"]}
    totals["Total"] = sum(totals.values())
    cols = st.columns(4)
    for col, (label, val) in zip(cols, totals.items()):
        col.metric(label, human_k(val))
    st.plotly_chart(
        chart_stacked(ns, f"Net Sales (kTL) — {selected_month}", ["Contract", "WP", "WO"]),
        use_container_width=True,
    )

with tab2:
    st.subheader(f"EBIT — {selected_month}")
    ebit = data["ebit"]
    keys_e = ["Contract", "Contract+WP", "Contract+WP+WO"]
    cols = st.columns(3)
    for col, k in zip(cols, keys_e):
        col.metric(k, human_k(sum(ebit[k])))
    st.plotly_chart(
        chart_grouped(ebit, f"EBIT (kTL) — {selected_month}", keys_e),
        use_container_width=True,
    )

with tab3:
    st.subheader(f"Order Intake — {selected_month}")
    oi = data["oi"]
    keys_oi = ["ENG", "MC", "T&SI", "NUC"]
    cols = st.columns(4)
    for col, k in zip(cols, keys_oi):
        col.metric(k, human_k(sum(oi[k])))

    view = st.radio("View", ["Monthly Chart", "By Project"], horizontal=True)
    if view == "Monthly Chart":
        st.plotly_chart(
            chart_stacked(oi, f"Order Intake (kTL) — {selected_month}", keys_oi),
            use_container_width=True,
        )
    else:
        proj = filter_oi(data["oi_projects"])
        fig = chart_projects(proj, f"Projects — {selected_month}")
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        if proj:
            df_proj = pd.DataFrame(proj)
            df_proj["Value (kTL)"] = df_proj["value"].apply(human_k)
            st.dataframe(
                df_proj[["project", "client", "bu", "Value (kTL)"]],
                hide_index=True, use_container_width=True,
            )
        else:
            st.info("No projects match the current filters.")

with tab4:
    st.subheader(f"WIP — {selected_month}")
    wip = filter_wip(data["wip_projects"])
    total_wip = sum(p["wip_tl"] for p in wip)
    st.metric(f"Total WIP ({len(wip)} projects)", human_tl(total_wip))
    fig = chart_wip(wip, f"WIP — {selected_month}")
    if fig:
        st.plotly_chart(fig, use_container_width=True)
    if wip:
        df_wip_display = pd.DataFrame(wip)
        df_wip_display["WIP TL"]  = df_wip_display["wip_tl"].apply(human_tl)
        df_wip_display["Inv OC"]  = df_wip_display["inv_oc"].apply(human_tl)
        df_wip_display["Prod OC"] = df_wip_display["prod_oc"].apply(human_tl)
        st.dataframe(
            df_wip_display[["name", "client", "bu", "orig_currency", "Inv OC", "Prod OC", "WIP TL"]],
            hide_index=True, use_container_width=True,
        )
    else:
        st.info("No WIP projects match the current filters.")

st.caption("MRC Business Review Dashboard")
