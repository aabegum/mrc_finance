# MRC Finance - Business Review Generator

Financial reporting automation system that transforms Excel budget data into interactive dashboards and PowerPoint presentations.

## Overview

| Component | File | Purpose |
|-----------|------|---------|
| **Dashboard** | `mbr_dashboard.py` | Streamlit interactive web UI |
| **PPTX Generator** | `generate_full_mbr_stacked.py` | Automated slide generation |
| **Config Loader** | `config_loader.py` | YAML configuration manager |
| **Configuration** | `config/config.yaml` | Centralized settings (327 lines) |

## Data Flow

```
Excel Files (monthly folders)
        │
        ▼
┌───────────────────┐
│  load_month()     │  ← Parses Summary TL, Order Intake, WIP & BL sheets
│  (mbr_dashboard)  │  ← Applies row offset detection
└───────────────────┘
        │
        ▼
┌──────────────────────────────┐
│  Data Dictionary              │
│  • ns  → Net Sales            │
│  • ebit → Earnings             │
│  • oi  → Order Intake          │
│  • wip → Work in Progress     │
│  • bu_ns/bu_ebit → By BU      │
└──────────────────────────────┘
        │
        ▼
    Outputs
    ┌─────────────┬──────────────┐
    │  Streamlit  │   PPTX       │
    │  Dashboard  │   Slides     │
    └─────────────┴──────────────┘
```

## Excel Source Data

Files named: `260331_Project_Budget_Analysis_March_v1.0.xlsx`

| Sheet | Purpose |
|-------|---------|
| `Summary TL` | Net Sales, EBIT, BU-level data |
| `Order Intake` | New orders by project/client/BU |
| `WIP & BL` | Work in Progress & Billing |
| `Ext. Prod.` | Margin history (optional) |

### Expected Excel Structure

- **Net Sales**: Categories row at ~14, data rows 15-17 (Contract/WP/WO)
- **EBIT**: Categories row at ~36, data rows 37-39
- **Order Intake**: Categories row at ~543, BU rows 544-547
- **WIP**: Data starts row 3, key columns: Name(2), Type(3), Client(4), BU(10), Currency(11), WIP_TL(98)

## Business Units (BU)

| Code | Full Name |
|------|-----------|
| ENG | Engineering |
| MC | MC (Metal Construction) |
| T&SI | TSI (Technical Services & Innovation) |
| NUC | Nuclear |

## Running the Dashboard

```bash
# Activate virtual environment
.venv\Scripts\activate  # Windows

# Run Streamlit
streamlit run mbr_dashboard.py
```

Dashboard tabs: **Net Sales | EBIT | Order Intake | WIP | Project History**

## Configuration

All settings in `config/config.yaml`:

```yaml
APP:
  Name: "MRC Business Review Generator"
  Version: "2.0"
  Active_Theme: "premium"

Excel_Mapping:
  Rows:
    Categories_NS_kTL: 14
    Slide2_Contract: 15
    # ... row indices
  Columns:
    WIP_Col_WIP_TL: 98
    # ... column indices
  Offsets:
    MON_NS_Base: 3
    OI_Col_Base: 6
```

### Key Config Sections

- **APP**: Output naming, formatting (decimals, abbreviations)
- **THEMES**: Color palettes (premium/default)
- **BU_Colors**: ENG=#0EA5E9, MC=#10B981, TSI=#FAB611, NUC=#93C572
- **Excel_Mapping**: Row/column indices for each data source
- **WIP_Table**: Filtering thresholds (min 1M TL)

## Key Features

1. **Dynamic Row Offset Detection** - Automatically handles Excel file variations
2. **Month Column Selection** - Reads only the current month's OI data
3. **Stacking Logic** - NS charts: incremental deltas; EBIT: overlapping totals
4. **Caching** - 1-hour TTL via `@st.cache_data`
5. **WIP Filtering** - Excludes Nuclear BU, min threshold 1M TL

## Folder Structure

```
MRC Finance/
├── config/
│   └── config.yaml           # All settings
├── main/                     # Working directory
│   └── 260331_March/         # Monthly Excel files
├── 260331_March/             # Root-level data (also)
├── mbr_dashboard.py          # Streamlit app
├── generate_full_mbr_stacked.py  # PPTX generator
├── config_loader.py          # YAML loader
├── COMPARISON_ANALYSIS.md    # Dashboard vs PPTX comparison
└── .venv/                    # Python virtual environment
```

## Development Notes

- Dashboard uses pandas for data loading, plotly for charts
- PPTX generator uses matplotlib for charts, python-pptx for slides
- ConfigLoader provides dot-notation access: `config.get("THEMES.premium.Colors.Primary.Hex")`
- Colors are defined once in config and shared across both outputs

## Dependencies

```
pandas
plotly
streamlit
pyyaml
python-pptx
matplotlib
seaborn
numpy
```
