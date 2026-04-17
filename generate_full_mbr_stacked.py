"""
generate_full_mbr_stacked.py
============================
STACKED-BAR variant of the MRC MBR generator.

Key differences from generate_full_mbr.py:
  • All bar charts are STACKED (not grouped).
  • NS / OI charts use true component stacking (ENG+MC+T&SI+NUC, etc.)
  • NS confidence-tier charts (Contract / WP / WO) stack INCREMENTAL deltas
    so each band shows the *extra* value that tier adds over the one below.
  • EBIT charts keep a clean grouped layout because confidence tiers are
    overlapping totals, not additive components; however value labels are
    added above each bar.
  • All labels use human-readable thousands formatting  (e.g. "103,695",
    "1.2 M", "–45,000").  Segments shorter than 3 % of the tallest bar are
    left unlabelled to avoid clutter.
"""

import os
import re
import gc
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib.patheffects as path_effects
import seaborn as sns
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from config_loader import ConfigLoader

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Load configuration system
config = ConfigLoader()
theme_info = config.get_active_theme_template()
global_theme = theme_info.get("theme", "default")
theme_colors = config.get(f"THEMES.{global_theme}.Colors", {})


def load_rgb(key, default_arr):
    arr = theme_colors.get(key, {}).get("RGB", default_arr)
    return RGBColor(arr[0], arr[1], arr[2])


# BU colors loaded from config (consistent across all charts regardless of active theme)
_bu_color_map = config.get("BU_Colors", {})
DEFAULT_BU_COLORS = [
    _bu_color_map.get("ENG", "#00B4D8"),
    _bu_color_map.get("MC", "#FF006E"),
    _bu_color_map.get("TSI", "#8338EC"),
    _bu_color_map.get("NUC", "#FFBE0B"),
]
PREMIUM_COLORS = theme_colors.get("Series_Palette", DEFAULT_BU_COLORS)
ACCENT = load_rgb("Primary", [0, 180, 216])
WHITE = load_rgb("White", [255, 255, 255])
DARK = load_rgb("Dark", [43, 45, 66])
GRAY_LT = load_rgb("Gray_Light", [230, 230, 230])
GRAY_MID = load_rgb("Gray_Mid", [150, 150, 150])
slide_bg_color = (
    load_rgb("Background", [255, 255, 255]) if "Background" in theme_colors else WHITE
)

# Number formatting controls
fmt_abbr = config.get("APP.Formatting.Abbreviate_Millions", True)
fmt_cf = config.get("APP.Formatting.Chart_Decimals", 1)
fmt_ax = config.get("APP.Formatting.Axis_Decimals", 0)
fmt_lhex = config.get("APP.Formatting.Chart_Label_Hex", None)
fmt_thex = config.get("APP.Formatting.Chart_Total_Hex", None)
fmt_3dm = config.get("APP.Formatting.Chart_3D_Modern", False)
fmt_3dc = config.get("APP.Formatting.Chart_3D_Classic", False)
fmt_3di = config.get("APP.Formatting.Chart_3D_Isometric", False)
fmt_halo = config.get("APP.Formatting.Chart_Label_Halo", False)
fmt_thold = config.get("APP.Formatting.Chart_Label_Threshold", 0)

# ──────────────────────────────────────────────────────────────
# DESIGN SYSTEM
# ──────────────────────────────────────────────────────────────
plt.style.use("default")
text_hex = theme_colors.get("Dark", {}).get("Hex", "#2B2D42")
sns.set_theme(
    style="whitegrid",
    rc={
        "axes.facecolor": "#FFFFFF",
        "figure.facecolor": "#FFFFFF",
        "axes.edgecolor": theme_colors.get("Gray_Light", {}).get("Hex", "#E0E0E0"),
        "grid.color": theme_colors.get("Gray_Light_Blue", {}).get("Hex", "#F0F0F0"),
        "text.color": text_hex,
        "axes.labelcolor": text_hex,
        "xtick.color": text_hex,
        "ytick.color": text_hex,
        "font.family": "sans-serif",
        "font.sans-serif": ["Segoe UI", "Inter", "Arial"],
    },
)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.50)
FONT_FAMILY = config.get(f"THEMES.{global_theme}.Typography.Font_Family", "Segoe UI")

MONTHS_SHORT = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def safe_float(val):
    try:
        if pd.isna(val):
            return 0.0
        return float(val)
    except Exception:
        return 0.0


def clean_label(val):
    """Clean category labels, avoiding '2025.0' float formatting for years."""
    if pd.isna(val):
        return ""
    try:
        # If it's a numeric value that equals its integer version (like 2025.0), return int str
        f_val = float(val)
        if f_val == int(f_val):
            return str(int(f_val))
    except:
        pass
    return str(val).strip()


def row_range(df, row, col_start, col_end):
    return [safe_float(df.iloc[row, c]) for c in range(col_start, col_end)]


def human_k(v, abbrev=None):
    """
    Format a kTL/kEUR value for chart labels.
    Values are already in thousands (k), so:
      >= 1 000 000 k  -> show as "X.XB"  (billions) - unless abbrev=False
      >= 1 000 k      -> show as "X.XM"  (millions in original currency) - unless abbrev=False
      otherwise      -> show with thousands comma separator, e.g. "103,695"
    Negative values get a minus sign.
    Set abbrev=False to always show full numbers (e.g. "163,529" instead of "163.529M")
    """
    if abbrev is None:
        abbrev = fmt_abbr
        
    sign = "-" if v < 0 else ""
    av = abs(v)
    if abbrev and av >= 1_000_000:
        return f"{sign}{av / 1_000_000:.{fmt_cf}f}B"
    if abbrev and av >= 1_000:
        return f"{sign}{av / 1_000:.{fmt_cf}f}M"
    return f"{sign}{av:,.{fmt_ax}f}"


def human_tl(v):
    """
    Format a raw TL (or OC) value for the WIP table.
    Values are in actual TL (not thousands), so:
      >= 1 000 000 000  → "X.XB"
      >= 1 000 000      → "X.XM"
      >= 1 000          → "X.Xk"
      otherwise         → comma-separated integer
    """
    sign = "-" if v < 0 else ""
    av = abs(v)
    if av >= 1_000_000_000:
        return f"{sign}{av / 1_000_000_000:.1f}B"
    if av >= 1_000_000:
        return f"{sign}{av / 1_000_000:.1f}M"
    if av >= 1_000:
        return f"{sign}{av / 1_000:.1f}k"
    return f"{sign}{av:,.0f}"


def _label_color_for_bg(hex_color):
    """Return white or dark label colour depending on background brightness."""
    if fmt_lhex:
        return fmt_lhex
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return "white" if lum < 160 else text_hex


def _draw_3d_extrusion(ax, x, y, w, h, color, is_top=False):
    """Draw angled sides and top. x-depth and y-depth are now fixed for consistency."""
    if h == 0:
        return
    # darkened side and lightened top
    from matplotlib.colors import to_rgb, to_hex

    rgb = to_rgb(color)
    side_col = to_hex([c * 0.75 for c in rgb])
    top_col = to_hex([min(1, c * 1.2) for c in rgb])

    depth_x = w * 0.22
    depth_y = w * 0.25 if fmt_3dc else 0

    # Side Face
    ax.fill(
        [x + w, x + w + depth_x, x + w + depth_x, x + w],
        [y, y + depth_y, y + h + depth_y, y + h],
        color=side_col,
        linewidth=0,
        zorder=2,
    )
    # Top Face (only if this segment is at the very top of its stack)
    if is_top and h != 0:
        ax.fill(
            [x, x + depth_x, x + w + depth_x, x + w],
            [y + h, y + h + depth_y, y + h + depth_y, y + h],
            color=top_col,
            linewidth=0,
            zorder=2,
        )


# ──────────────────────────────────────────────────────────────
# CHART – STACKED BAR (component stacking)
# ──────────────────────────────────────────────────────────────
def save_stacked_chart(
    name,
    categories,
    series_dict,
    output_path,
    ylabel="kTL",
    use_default_colors=False,
    color_palette=None,
    label_threshold_override=None,
    extra_bar=None,
    n_yoy=2,
    extra_legend=None,
):
    """
    True stacked bar chart.  Each series is a separate stack layer.
    Segment labels are shown in human-readable format for segments
    that are at least 3 % of the tallest bar height.
    """
    n_cats = len(categories)
    # Spacing: group the first n_yoy bars (YoY) close together, then a larger
    # gap before the monthly bars, then regular spacing for months.
    x = []
    curr_x = 0
    for i in range(n_cats):
        x.append(curr_x)
        if i < n_yoy - 1:
            curr_x += 0.7   # within the YoY group — bars sit close
        elif i == n_yoy - 1:
            curr_x += 2.0   # wider gap after last YoY bar, before months
        else:
            curr_x += 1.0   # regular monthly spacing
    x = np.array(x)
    bar_w = 0.55
    # x-position of separator line = midpoint of the gap between last YoY and first month
    _sep_x = (x[n_yoy - 1] + x[n_yoy]) / 2 if n_cats > n_yoy else None

    fig, ax = plt.subplots(figsize=(13, 5.5))

    # Build stacks and track bottoms separately for positive and negative
    pos_bottoms = np.zeros(n_cats)
    neg_bottoms = np.zeros(n_cats)

    # Find each bar's total height and the overall max for reference
    series_list = list(series_dict.items())
    total_pos = np.zeros(n_cats)
    for _, vals in series_list:
        arr = np.array(vals, dtype=float)
        total_pos += np.where(arr > 0, arr, 0)
    max_height = total_pos.max() if total_pos.max() > 0 else 1
    _thold_pct = (
        label_threshold_override if label_threshold_override is not None else fmt_thold
    )
    # threshold is computed per-bar (relative to that bar's own total height)
    # so small bars still show labels for their segments proportionally.
    # Falls back to a global threshold for completely empty bars.
    def _bar_threshold(xi):
        h = total_pos[xi]
        return h * (_thold_pct / 100.0) if h > 0 else max_height * (_thold_pct / 100.0)

    # Choose color palette: CUSTOM, DEFAULT for Business Units, or PREMIUM for others
    if color_palette:
        pass  # use provided palette
    elif use_default_colors:
        color_palette = DEFAULT_BU_COLORS
    else:
        color_palette = PREMIUM_COLORS

    bar_handles = []
    for idx, (label, vals) in enumerate(series_list):
        color = color_palette[idx % len(color_palette)]
        arr = np.array(vals, dtype=float)
        pos = np.where(arr >= 0, arr, 0)
        neg = np.where(arr < 0, arr, 0)

        # Option B (Classic) or C (Isometric)
        if fmt_3dc or fmt_3di:
            for xi, (pv, nv, pb, nb) in enumerate(
                zip(pos, neg, pos_bottoms, neg_bottoms)
            ):
                # ONLY draw top-cap if this is the last series (usually WO)
                is_top = idx == len(series_list) - 1
                if pv > 0:
                    _draw_3d_extrusion(
                        ax, x[xi] - bar_w / 2, pb, bar_w, pv, color, is_top=is_top
                    )
                if nv < 0:
                    _draw_3d_extrusion(ax, x[xi] - bar_w / 2, nb, bar_w, nv, color)

        # Baseline bars
        bar_pos = ax.bar(
            x,
            pos,
            bar_w,
            bottom=pos_bottoms,
            label=label,
            color=color,
            alpha=0.92,
            linewidth=0,
            zorder=3,
        )
        bar_neg = ax.bar(
            x, neg, bar_w, bottom=neg_bottoms, color=color, alpha=0.65, linewidth=0, zorder=3
        )

        # Simulated 3D Pill Effect (White Highlights)
        if fmt_3dm:
            ax.bar(
                x,
                pos,
                bar_w * 0.35,
                bottom=pos_bottoms,
                color="white",
                alpha=0.25,
                linewidth=0,
                zorder=4,
            )
            ax.bar(
                x,
                neg,
                bar_w * 0.35,
                bottom=neg_bottoms,
                color="white",
                alpha=0.15,
                linewidth=0,
                zorder=4,
            )

        bar_handles.append(bar_pos)

        lbl_col = _label_color_for_bg(color)

        # Label each segment that is wide enough to read
        for xi, (pv, nv, pb, nb) in enumerate(zip(pos, neg, pos_bottoms, neg_bottoms)):
            # TOTAL value for this bar (full bar height, not just accumulated bottom)
            bar_total = total_pos[xi]

            # --- POSITIVE SEGMENTS ---
            # Rule 1: Don't show if zero
            # Rule 2: Don't show if smaller than threshold
            # Rule 3: Hide internal label for YoY bars only, OR if segment dominates the whole bar
            is_redundant = (xi < n_yoy) or (
                (pv >= bar_total * 0.98) if bar_total > 0 else False
            )

            if pv > 0 and pv >= _bar_threshold(xi) and not is_redundant:
                centre_y = pb + pv / 2
                t_obj = ax.text(
                    x[xi],
                    centre_y,
                    human_k(pv),
                    ha="center",
                    va="center",
                    fontsize=7.5,
                    fontweight="bold",
                    color=lbl_col,
                    zorder=5,
                )
                if fmt_halo:
                    t_obj.set_path_effects(
                        [path_effects.withStroke(linewidth=2, foreground="white")]
                    )

            # --- NEGATIVE SEGMENTS ---
            if abs(nv) > 0 and abs(nv) >= _bar_threshold(xi) and not is_redundant:
                centre_y = nb + nv / 2
                t_obj = ax.text(
                    x[xi],
                    centre_y,
                    human_k(nv),
                    ha="center",
                    va="center",
                    fontsize=7.5,
                    fontweight="bold",
                    color=lbl_col,
                    zorder=5,
                )
                if fmt_halo:
                    t_obj.set_path_effects(
                        [path_effects.withStroke(linewidth=2, foreground="white")]
                    )

        pos_bottoms += pos
        neg_bottoms += neg

    # Total label above the full bar
    total_lbl_col = fmt_thex if fmt_thex else (fmt_lhex if fmt_lhex else text_hex)
    for xi, tot in enumerate(pos_bottoms):
        if tot != 0:
            # Use angled labels for better visibility as requested
            t_obj = ax.text(
                x[xi],
                tot + max_height * 0.015 if tot > 0 else tot - max_height * 0.015,
                human_k(tot),
                ha="center",
                va="bottom" if tot > 0 else "top",
                rotation=45,
                fontsize=8.5,
                fontweight="bold",
                color=total_lbl_col,
                zorder=5,
            )
            if fmt_halo:
                t_obj.set_path_effects(
                    [path_effects.withStroke(linewidth=2, foreground="white")]
                )

    # ── Extra standalone bar (e.g. ABNS in red) ──────────────────
    if extra_bar is not None:
        eb_label, eb_val, eb_color = extra_bar
        eb_x = x[-1] + 1.6  # gap before it
        if eb_val > 0:
            ax.bar(eb_x, eb_val, bar_w, color=eb_color, alpha=0.92, linewidth=0, zorder=3)
            ax.text(
                eb_x,
                eb_val + max_height * 0.015,
                human_k(eb_val),
                ha="center",
                va="bottom",
                rotation=45,
                fontsize=8.5,
                fontweight="bold",
                color=eb_color,
                zorder=5,
            )
        x = np.append(x, eb_x)
        categories = list(categories) + [eb_label]

    # Separator line between YoY bars and monthly bars
    if _sep_x is not None:
        ax.axvline(_sep_x, color="#CBD5E1", linewidth=1.0, linestyle="--", zorder=1, alpha=0.8)

    chart_text_col = fmt_lhex if fmt_lhex else text_hex
    ax.set_title(
        name.replace("_", " "),
        fontsize=15,
        fontweight="bold",
        pad=18,
        color=chart_text_col,
    )
    ax.set_ylabel(ylabel, fontsize=11, color=chart_text_col)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, rotation=30, ha="right", fontsize=9)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.{fmt_ax}f}"))
    # Build legend handles: main series + optional extra entries (e.g. ABNS bar)
    legend_handles, legend_labels = ax.get_legend_handles_labels()
    if extra_legend:
        from matplotlib.patches import Patch
        for _lbl, _col in extra_legend:
            legend_handles.append(Patch(facecolor=_col, alpha=0.92, linewidth=0))
            legend_labels.append(_lbl)
    n_legend = len(legend_labels)
    ax.legend(
        legend_handles,
        legend_labels,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.28),
        ncol=min(4, n_legend),
        frameon=False,
        fontsize=10,
    )
    ax.axhline(0, color="#CBD5E1", linewidth=0.8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E1")
    ax.spines["bottom"].set_color("#CBD5E1")

    fig.tight_layout()
    fig.savefig(
        output_path,
        dpi=150,
        bbox_inches="tight",
        facecolor="#FFFFFF",
        transparent=False,
    )
    plt.close(fig)
    gc.collect()
    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError(f"Chart not written: {output_path}")
    return output_path


# ──────────────────────────────────────────────────────────────
# CHART – GROUPED BAR with value labels (used for EBIT tiers)
# ──────────────────────────────────────────────────────────────
def save_grouped_chart(
    name,
    categories,
    series_dict,
    output_path,
    ylabel="kTL",
    percents=None,
    color_palette=None,
):
    """
    Grouped bar chart with human-readable value labels above each bar.
    Used for EBIT where each series is an overlapping total (not additive).
    """
    n_cats = len(categories)
    n_ser = len(series_dict)
    # Spacing: group first 2 YoY bars close, then gap before months
    x = []
    curr_x = 0
    for i in range(n_cats):
        x.append(curr_x)
        if i < 2:
            curr_x += 0.7   # 2025 + 2026 Target + BL close together
        elif i == 2:
            curr_x += 2.0   # wider gap after BL, before Jan
        else:
            curr_x += 1.0
    x = np.array(x)
    bar_w = 0.75 / n_ser  # Slightly wider bars
    _sep_x_grp = (x[2] + x[3]) / 2 if n_cats > 3 else None

    fig, ax = plt.subplots(figsize=(13, 5.5))
    all_vals = [v for vals in series_dict.values() for v in vals]
    max_abs = max(abs(v) for v in all_vals) if all_vals else 1

    for i, (label, vals) in enumerate(series_dict.items()):
        pal = color_palette if color_palette else PREMIUM_COLORS
        color = pal[i % len(pal)]
        offset = (i - n_ser / 2 + 0.5) * bar_w
        arr = np.array(vals, dtype=float)
        # Option B (Classic) or C (Isometric)
        if fmt_3dc or fmt_3di:
            for xi, v in enumerate(arr):
                _draw_3d_extrusion(
                    ax, x[xi] + offset - bar_w / 2, 0, bar_w, v, color, is_top=True
                )

        bars = ax.bar(
            x + offset, arr, bar_w, label=label, color=color, alpha=0.90, linewidth=0, zorder=3
        )

        # Highlight (Pill Effect)
        if fmt_3dm:
            ax.bar(x + offset, arr, bar_w * 0.35, color="white", alpha=0.22, linewidth=0, zorder=4)

        # Value label inside the bar (for EBIT with percentages)
        for xi, val in enumerate(arr):
            # Only show label for the LAST series (Total Tier)
            if i < n_ser - 1:
                continue

            if abs(val) < 0.1:  # Hide zero labels
                continue

            # Use full numbers (no B/M suffix) inside bars
            txt = human_k(val, abbrev=False)

            # Include percentage if provided - show INSIDE bar
            pct_inside = ""
            if percents and xi < len(percents):
                p_v = percents[xi]
                if abs(p_v) > 0.0001:
                    try:
                        if abs(p_v) < 2.0:  # Heuristic for decimal percentages
                            pct_inside = f"{p_v:.1%}"
                        else:
                            pct_inside = f"{p_v:.0f}%"
                    except:
                        pct_inside = f"{p_v}"

            # Place percentage INSIDE the bar (white text on colored background)
            if pct_inside and val > 0:
                inside_y = val * 0.5  # Middle of bar
                t_obj = ax.text(
                    x[xi] + offset,
                    inside_y,
                    pct_inside,
                    ha="center",
                    va="center",
                    fontsize=7,
                    fontweight="bold",
                    color="#1E3A8A",  # Deep Sapphire Blue for readability
                    zorder=5,
                )
                if fmt_halo:
                    t_obj.set_path_effects(
                        [path_effects.withStroke(linewidth=1.5, foreground="black")]
                    )

            # Value label ABOVE the bar
            va = "bottom" if val >= 0 else "top"
            y_off = max_abs * 0.018 if val >= 0 else -max_abs * 0.018
            y = val + y_off

            lbl_color = fmt_thex if fmt_thex else (fmt_lhex if fmt_lhex else text_hex)

            t_obj = ax.text(
                x[xi] + offset,
                y,
                txt,
                ha="center",
                va=va,
                rotation=45,
                fontsize=8,
                fontweight="bold",
                color=lbl_color,
                zorder=5,
            )
            if fmt_halo:
                t_obj.set_path_effects(
                    [path_effects.withStroke(linewidth=2, foreground="white")]
                )

    # Fallback: label positions where the last series is 0 (was NaN in Excel)
    # but an earlier series has a real value (e.g., YoY bars only have Contract data)
    fallback_lbl_col = fmt_thex if fmt_thex else (fmt_lhex if fmt_lhex else text_hex)
    series_list = list(series_dict.items())
    last_series_arr = np.array(series_list[-1][1], dtype=float)
    for xi in range(len(categories)):
        if abs(last_series_arr[xi]) >= 0.1:
            continue  # Already labeled by main loop
        for si in range(n_ser - 2, -1, -1):
            ref_val = np.array(series_list[si][1], dtype=float)[xi]
            if abs(ref_val) < 0.1:
                continue
            ref_offset = (si - n_ser / 2 + 0.5) * bar_w
            txt = human_k(ref_val)
            if percents and xi < len(percents):
                p_v = percents[xi]
                if abs(p_v) > 0.0001:
                    txt += f" ({p_v:.1%})" if abs(p_v) < 2.0 else f" ({p_v:.0f}%)"
            y_off = max_abs * 0.018 if ref_val >= 0 else -max_abs * 0.018
            t_obj = ax.text(
                x[xi] + ref_offset,
                ref_val + y_off,
                txt,
                ha="center",
                va="bottom",
                rotation=45,
                fontsize=8,
                fontweight="bold",
                color=fallback_lbl_col,
                zorder=5,
            )
            if fmt_halo:
                t_obj.set_path_effects(
                    [path_effects.withStroke(linewidth=2, foreground="white")]
                )
            break

    # Separator line between YoY and monthly bars
    if _sep_x_grp is not None:
        ax.axvline(_sep_x_grp, color="#CBD5E1", linewidth=1.0, linestyle="--", zorder=1, alpha=0.8)

    chart_text_col = fmt_lhex if fmt_lhex else text_hex
    ax.set_title(
        name.replace("_", " "),
        fontsize=15,
        fontweight="bold",
        pad=18,
        color=chart_text_col,
    )
    ax.set_ylabel(ylabel, fontsize=11, color=chart_text_col)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, rotation=30, ha="right", fontsize=9)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.{fmt_ax}f}"))
    ax.axhline(0, color="#CBD5E1", linewidth=0.8)
    ax.legend(
        loc="lower center",
        bbox_to_anchor=(0.5, -0.28),
        ncol=min(4, len(series_dict)),
        frameon=False,
        fontsize=10,
    )
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E1")
    ax.spines["bottom"].set_color("#CBD5E1")

    fig.tight_layout()
    fig.savefig(
        output_path,
        dpi=150,
        bbox_inches="tight",
        facecolor="#FFFFFF",
        transparent=False,
    )
    plt.close(fig)
    gc.collect()
    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError(f"Chart not written: {output_path}")
    return output_path



# ──────────────────────────────────────────────────────────────
# CHART – BU EBIT  (cumulative bars matching reference PPTX)
# ──────────────────────────────────────────────────────────────
def save_bu_ebit_chart(
    name,
    categories,
    values,
    percents,
    output_path,
    ylabel="kTL",
    bu_color="#0EA5E9",
    bu_label="BU",
):
    """
    Single-series cumulative bar chart for BU EBIT slides.
    Replicates the reference MBR style:
      • First bar  → Forest Green  (2025 actuals)
      • Second bar → Deep Sapphire (2026 Target)
      • Remaining  → bu_color (monthly cumulative YTD)
      • EBIT % shown inside bars (white bold)
      • Value labels at top (angled)
      • Data table row at the bottom of the figure
    """
    # BU arrays (col52-65): [0]=2025, [1]=2026Target, [2]=Jan, …, [13]=Dec — no empty column.
    # Uniform spacing so ax.table column positions align exactly with bar positions.
    n = len(categories)
    x = np.arange(n, dtype=float)
    bar_w = 0.55
    _sep_x_ebit = (x[1] + x[2]) / 2 if n > 2 else None

    # Per-bar colors: 2025=dark blue, 2026 Target+monthly=bu_color
    bar_colors = ["#1E3A8A"] + [bu_color] * max(0, n - 1)

    fig, ax = plt.subplots(figsize=(13, 5.5))

    for xi, (val, col) in enumerate(zip(values, bar_colors)):
        if val == 0:
            continue
        ax.bar(x[xi], val, bar_w, color=col, alpha=0.88, zorder=3)

        # % label inside bar (white, bold)
        pct_str = ""
        if percents and xi < len(percents):
            pv = percents[xi]
            try:
                pv_f = float(pv)
                if abs(pv_f) > 0.0001:
                    pct_str = f"{pv_f:.1%}" if abs(pv_f) < 2 else f"{pv_f:.0f}%"
            except Exception:
                pass
        if pct_str and val > 0:
            inside_y = val * 0.5
            ax.text(
                x[xi], inside_y, pct_str,
                ha="center", va="center",
                fontsize=8, fontweight="bold",
                color="white", zorder=5,
            )

        # Value labels above bars intentionally omitted for BU EBIT charts

    # Separator line between YoY and monthly bars
    if _sep_x_ebit is not None:
        ax.axvline(_sep_x_ebit, color="#CBD5E1", linewidth=1.0, linestyle="--", zorder=1, alpha=0.8)

    chart_text_col = fmt_lhex if fmt_lhex else text_hex
    ax.set_title(
        name.replace("_", " "), fontsize=15, fontweight="bold",
        pad=18, color=chart_text_col,
    )
    ax.set_ylabel(ylabel, fontsize=11, color=chart_text_col)
    ax.set_xticks(x)
    ax.set_xticklabels([])   # table column headers serve as x-axis labels; hide duplicates
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v, _: f"{v:,.{fmt_ax}f}"))
    ax.axhline(0, color="#CBD5E1", linewidth=0.8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#CBD5E1")
    ax.spines["bottom"].set_color("#CBD5E1")

    # Data table row at the bottom (bu_label | v1 | v2 | …)
    table_vals = [[human_k(v, abbrev=False) if v != 0 else "-" for v in values]]
    the_table = ax.table(
        cellText=table_vals,
        rowLabels=[bu_label],
        colLabels=categories,
        loc="bottom",
        cellLoc="center",
    )
    the_table.auto_set_font_size(False)
    the_table.set_fontsize(8.5)
    the_table.scale(1.0, 1.45)
    # Style header row and row label
    for (row, col), cell in the_table.get_celld().items():
        cell.set_edgecolor("#CBD5E1")
        if row == 0:  # column-header row
            cell.set_facecolor(bu_color)
            cell.set_text_props(color="white", fontweight="bold")
        elif col == -1:  # row label
            cell.set_facecolor(bu_color)
            cell.set_text_props(color="white", fontweight="bold")
        else:
            cell.set_facecolor("#F4F6F9")

    # shift bottom margin to make room for table
    fig.subplots_adjust(bottom=0.28)
    fig.savefig(
        output_path, dpi=150, bbox_inches="tight",
        facecolor="#FFFFFF", transparent=False,
    )
    plt.close(fig)
    gc.collect()
    if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError(f"Chart not written: {output_path}")
    return output_path



# ──────────────────────────────────────────────────────────────
# INCREMENTAL DELTA HELPER
# ──────────────────────────────────────────────────────────────
def incremental_series(series_dict):
    """
    Given an ordered dict of confidence tiers like:
       {"Contract": [v1…], "Contract+WP": [v2…], "Contract+WP+WO": [v3…]}
    Returns a new dict with incremental deltas so they stack correctly:
       {"Contract": [...], "+WP": [...], "+WO": [...]}
    The first series is kept as-is; each subsequent series is (current - previous).
    """
    keys = list(series_dict.keys())
    vals = list(series_dict.values())
    result = {}
    for i, key in enumerate(keys):
        if i == 0:
            result[key] = vals[i][:]
        else:
            label = "+" + key.split("+")[-1] if "+" in key else f"Δ {key}"
            prev_arr = np.array(vals[i - 1], dtype=float)
            curr_arr = np.array(vals[i], dtype=float)
            result[label] = list(curr_arr - prev_arr)
    return result


# ──────────────────────────────────────────────────────────────
# PPTX SLIDE HELPERS  (identical to base script)
# ──────────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────────
# TEMPLATE-AWARE SLIDE HELPERS
# ──────────────────────────────────────────────────────────────
def duplicate_template_content_slide(prs, template_slide_idx=1):
    """
    Duplicate a template slide (e.g., content template) to preserve
    branding, logos, and decorative elements.

    Args:
        prs: Presentation object
        template_slide_idx: Index of template slide to duplicate (default 1 = content slide)

    Returns:
        New slide (copy of template)
    """
    if template_slide_idx >= len(prs.slides):
        # Fallback: create blank if template not available
        return prs.slides.add_slide(prs.slide_layouts[6])

    # Get the template slide
    template_slide = prs.slides[template_slide_idx]

    # Add a new blank slide
    blank_slide_layout = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank_slide_layout)

    # Copy background
    try:
        new_slide.background.fill.solid()
        # Copy from template if possible
        if template_slide.background.fill.type:
            new_slide.background.fill.fore_color.rgb = (
                template_slide.background.fill.fore_color.rgb
            )
    except:
        pass

    # Copy all shapes from template
    for shape in template_slide.shapes:
        el = shape.element
        newel = el.__copy__()
        new_slide.shapes._spTree.insert_element_before(newel, "p:extLst")

    return new_slide


def add_blank_slide(prs, use_template=True, clean_placeholders=True, source_slide_idx=1):
    """
    Create a new slide. If template available and use_template=True,
    copy branding/sidebar shapes from the specified source slide.

    Keeps only SIDEBAR / BRANDING shapes:
      - Shapes positioned in the left sidebar area (left < ~0.45 inch)
      - Background fills (full-width shapes with no text)

    Removes all CONTENT shapes:
      - Known placeholder texts: [CHART TITLE], [SUBTITLE], [CHART IMAGE], [KEY METRICS]
      - Any text shape in the main content area (real titles, subtitles added by add_header)
      - Picture shapes in the content area (embedded chart PNGs have broken rel-IDs)

    Args:
        source_slide_idx: Index of the slide to copy branding from (default=1).
                          Pass the WIP template slide index when creating WIP overflow slides
                          so branding is copied from the clean template, not a populated chart slide.
    """
    # Content area starts where add_header places elements (Inches(0.45))
    # Shapes with left >= this are in the main content area and must be filtered.
    _CONTENT_LEFT = Inches(0.45)
    _PLACEHOLDER_TEXTS = {"[CHART TITLE]", "[SUBTITLE]", "[CHART IMAGE]", "[KEY METRICS]"}
    _PICTURE_SHAPE_TYPE = 13  # MSO_SHAPE_TYPE.PICTURE

    # Try to duplicate template content slide
    if use_template and len(prs.slides) > source_slide_idx:
        try:
            template_slide = prs.slides[source_slide_idx]
            blank_slide_layout = prs.slide_layouts[6]
            new_slide = prs.slides.add_slide(blank_slide_layout)

            # Copy background
            try:
                new_slide.background.fill.solid()
                new_slide.background.fill.fore_color.rgb = (
                    template_slide.background.fill.fore_color.rgb
                )
            except:
                pass

            # Copy only sidebar/branding shapes, discard all content-area shapes
            shapes_data = []
            skip_count = 0
            for shape in template_slide.shapes:
                is_placeholder = False

                shape_left = getattr(shape, "left", 0) or 0
                in_content_area = shape_left >= _CONTENT_LEFT

                # 1. Known placeholder text markers
                if hasattr(shape, "text"):
                    text = shape.text.strip()
                    if text in _PLACEHOLDER_TEXTS:
                        is_placeholder = True
                    # 2. Any real (non-empty) text in the content area
                    #    — these are chart titles / subtitles added by add_header
                    elif text and in_content_area:
                        is_placeholder = True

                # 3. Picture shapes in the content area have embedded relationship IDs
                #    that do not transfer correctly and cause PowerPoint corruption errors.
                if not is_placeholder and shape.shape_type == _PICTURE_SHAPE_TYPE and in_content_area:
                    is_placeholder = True

                if is_placeholder:
                    skip_count += 1
                shapes_data.append((shape.element, is_placeholder))

            # Copy non-placeholder shapes only
            copy_count = 0
            for el, is_placeholder in shapes_data:
                if not is_placeholder:
                    newel = el.__copy__()
                    new_slide.shapes._spTree.insert_element_before(newel, "p:extLst")
                    copy_count += 1
            print(
                f"[add_blank_slide] {len(shapes_data)} shapes: copied {copy_count}, skipped {skip_count}"
            )

            return new_slide
        except Exception as e:
            pass

    # Fallback: create blank slide without template
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = slide_bg_color

    if global_theme == "premium":
        content_bg = slide.shapes.add_shape(
            1, Inches(0), Inches(0.9), Inches(13.33), Inches(6.6)
        )
        content_bg.fill.solid()
        content_bg.fill.fore_color.rgb = WHITE
        content_bg.line.fill.background()
    return slide


def add_header(slide, title_text, subtitle_text=""):
    """
    Add title and subtitle textboxes to slide.
    (Template placeholders are pre-removed by add_blank_slide)
    """
    tb = slide.shapes.add_textbox(
        Inches(0.45), Inches(0.18), Inches(12.4), Inches(0.72)
    )
    tf = tb.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(26)
    p.font.color.rgb = DARK if global_theme == "premium" else ACCENT

    div = slide.shapes.add_shape(
        1, Inches(0.45), Inches(0.92), Inches(12.4), Inches(0.025)
    )
    div.fill.solid()
    div.fill.fore_color.rgb = ACCENT if global_theme == "premium" else GRAY_LT
    div.line.fill.background()

    if subtitle_text:
        tb2 = slide.shapes.add_textbox(
            Inches(0.45), Inches(0.93), Inches(12.4), Inches(0.38)
        )
        p2 = tb2.text_frame.paragraphs[0]
        p2.text = subtitle_text
        p2.font.name = FONT_FAMILY
        p2.font.size = Pt(11)
        p2.font.color.rgb = GRAY_MID


def add_chart_image(
    slide, image_path, left=Inches(0.45), top=Inches(1.22), width=Inches(12.4)
):
    if os.path.exists(image_path) and os.path.getsize(image_path) > 0:
        slide.shapes.add_picture(image_path, left, top, width=width)
    else:
        print(f"  WARNING: chart image missing or empty — {image_path}")


def add_callout_box(
    slide, lines, left=Inches(10.5), top=Inches(6.15), width=Inches(2.5), height=None
):
    """
    Add a callout box to a slide with key metrics.
    Position parameters allow placing multiple callouts.
    """
    box_height = height if height else Inches(0.35 * len(lines) + 0.15)

    box = slide.shapes.add_shape(5, left, top, width, box_height)
    box.fill.solid()
    callout_bg = (
        load_rgb("Gray_Light_Blue", [240, 252, 255])
        if "Gray_Light_Blue" in theme_colors
        else RGBColor(240, 252, 255)
    )
    box.fill.fore_color.rgb = callout_bg
    box.line.color.rgb = ACCENT
    box.line.width = Pt(1)

    tb = slide.shapes.add_textbox(
        left + Inches(0.10),
        top + Inches(0.05),
        width - Inches(0.20),
        box_height - Inches(0.10),
    )
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line_item in enumerate(lines):
        label = line_item[0]
        value = line_item[1]
        # Optional 3rd element = hex color override for this value
        if len(line_item) > 2 and line_item[2]:
            h = line_item[2].lstrip("#")
            val_rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        else:
            val_rgb = ACCENT
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = label
        p.font.name = FONT_FAMILY
        p.font.size = Pt(8.0)
        p.font.color.rgb = GRAY_MID
        p.space_after = Pt(0)

        p2 = tf.add_paragraph()
        p2.text = value
        p2.font.name = FONT_FAMILY
        p2.font.bold = True
        p2.font.size = Pt(11.0)
        p2.font.color.rgb = val_rgb
        p2.space_after = Pt(2)


def add_second_callout_box(
    slide, lines, left=Inches(0.45), top=Inches(6.15), width=Inches(3.5), height=None,
    title="ABNS Projects", border_color=None,
):
    """
    ABNS project list callout box.  One line per project for maximum readability.
    Format per row: "▸ ProjectName (truncated)  —  1,234 kTL"
    Auto-scales font so the box never exceeds slide bottom (7.40").
    """
    _nuc_rgb = RGBColor(1, 107, 97)
    _border = border_color if border_color else _nuc_rgb

    MAX_BOT = Inches(7.40)
    available = MAX_BOT - top
    # 1 header line + N project lines, each 0.22" tall
    n_proj = len(lines)
    header_h = Inches(0.28)
    per_line_h = Inches(0.22)
    needed = header_h + per_line_h * n_proj + Inches(0.10)

    font_pt = 10.0
    if needed > available and n_proj > 0:
        per_line_h = max(Inches(0.18), (available - header_h - Inches(0.10)) / n_proj)
        scale = per_line_h / Inches(0.22)
        font_pt = max(7.0, round(10.0 * scale, 1))

    box_height = height if height else int(header_h + per_line_h * n_proj + Inches(0.10))

    box = slide.shapes.add_shape(5, left, top, width, box_height)
    box.fill.solid()
    callout_bg = (
        load_rgb("Gray_Light_Blue", [240, 252, 255])
        if "Gray_Light_Blue" in theme_colors
        else RGBColor(240, 252, 255)
    )
    box.fill.fore_color.rgb = callout_bg
    box.line.color.rgb = _border
    box.line.width = Pt(1.2)

    tb = slide.shapes.add_textbox(
        left + Inches(0.10),
        top + Inches(0.05),
        width - Inches(0.18),
        box_height - Inches(0.08),
    )
    tf = tb.text_frame
    tf.word_wrap = False  # prevent wrapping — truncation handles length

    # Header row
    p_hdr = tf.paragraphs[0]
    p_hdr.text = title
    p_hdr.font.name = FONT_FAMILY
    p_hdr.font.bold = True
    p_hdr.font.size = Pt(font_pt + 1.0)
    p_hdr.font.color.rgb = _border
    p_hdr.space_after = Pt(2)

    # One project per line: "▸ Name — value"
    max_name_chars = max(18, int(width / Inches(0.065)) - 12)
    for _lbl, _val_str in lines:
        # _lbl is "#N", _val_str is "Name: value" — split on last ": " to get name+val
        if ": " in _val_str:
            proj_name, val_part = _val_str.rsplit(": ", 1)
        else:
            proj_name, val_part = _val_str, ""
        proj_short = proj_name[:max_name_chars] + "…" if len(proj_name) > max_name_chars else proj_name
        row_text = f"▸ {proj_short}  —  {val_part}" if val_part else f"▸ {proj_short}"

        p = tf.add_paragraph()
        p.text = row_text
        p.font.name = FONT_FAMILY
        p.font.size = Pt(font_pt)
        p.font.color.rgb = _nuc_rgb
        p.space_after = Pt(0)


_tbl_cfg = config.get(f"THEMES.{global_theme}.Table", {})
_tbl_header_pt = Pt(_tbl_cfg.get("Header", {}).get("Font_Size", 10))
_tbl_data_pt = Pt(_tbl_cfg.get("Data_Text", {}).get("Font_Size", 9.5))


def style_table_header(cell):
    cell.fill.solid()
    cell.fill.fore_color.rgb = ACCENT
    for para in cell.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = FONT_FAMILY
            run.font.color.rgb = WHITE
            run.font.bold = True
            run.font.size = _tbl_header_pt


def style_table_data(cell, alt=False):
    cell.fill.solid()
    alt_bg = (
        load_rgb("Gray_Light_Blue", [240, 252, 255])
        if "Gray_Light_Blue" in theme_colors
        else RGBColor(240, 252, 255)
    )
    cell.fill.fore_color.rgb = alt_bg if alt else WHITE
    for para in cell.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = FONT_FAMILY
            run.font.color.rgb = DARK
            run.font.size = _tbl_data_pt


def style_table_total(cell):
    """Bold total row with accent background."""
    cell.fill.solid()
    total_bg = (
        load_rgb("Gray_Light", [230, 230, 230])
        if "Gray_Light" in theme_colors
        else RGBColor(230, 230, 230)
    )
    cell.fill.fore_color.rgb = total_bg
    for para in cell.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = FONT_FAMILY
            run.font.color.rgb = DARK
            run.font.bold = True
            run.font.size = _tbl_data_pt


def _set_cell_text(cell, text, font_size=Pt(9.5), bold=False, align=PP_ALIGN.LEFT):
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    if p.runs:
        run = p.runs[0]
    else:
        run = p.add_run()
    run.text = str(text)
    run.font.name = FONT_FAMILY
    run.font.size = font_size
    run.font.bold = bold


def update_slide_subtitle(slide, text):
    """Replace [SUBTITLE] placeholder on a slide with the given text."""
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip() == "[SUBTITLE]":
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = text
            run.font.name = FONT_FAMILY
            run.font.size = Pt(10)
            run.font.color.rgb = GRAY_MID
            break


def update_slide_title(slide, old_text, new_text):
    """Find a title shape matching old_text and replace with new_text, preserving formatting."""
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip() == old_text:
            tf = shape.text_frame
            # Preserve the first paragraph's existing font style
            if tf.paragraphs and tf.paragraphs[0].runs:
                run = tf.paragraphs[0].runs[0]
                run.text = new_text
            else:
                tf.clear()
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = new_text
            break


def add_oi_table_to_slide(slide, oi_projects, max_rows=15, chart_total=None):
    """
    Add Order Intake project details table to slide 6.
    Columns: Project | Customer | BU | Value (kTL)

    Args:
        slide: PPTX slide object
        oi_projects: List of (project, client, value, bu) tuples
        max_rows: Maximum rows to show
        chart_total: Optional total from chart for discrepancy note
    """
    rows_data = oi_projects
    n_data = len(rows_data)
    has_others = False

    n_rows = 1 + n_data + 1  # header + data + total row

    left = Inches(0.55)
    top = Inches(1.38)
    width = Inches(12.23)
    height = Inches(min(5.6, n_rows * 0.36 + 0.05))

    tbl_shape = slide.shapes.add_table(n_rows, 4, left, top, width, height)
    tbl = tbl_shape.table

    # Column widths: Project 42%, Customer 28%, BU 10%, Value 20%
    tbl.columns[0].width = int(width * 0.42)
    tbl.columns[1].width = int(width * 0.28)
    tbl.columns[2].width = int(width * 0.10)
    tbl.columns[3].width = int(width * 0.20)

    # Header row
    headers = ["Project", "Customer", "BU", "Value (kTL)"]
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        _set_cell_text(
            cell,
            h,
            font_size=Pt(11),
            bold=True,
            align=PP_ALIGN.RIGHT if ci == 3 else PP_ALIGN.LEFT,
        )
        style_table_header(cell)

    # Data rows
    total_val = 0.0
    for ri, (project, client, val, bu) in enumerate(rows_data):
        alt = ri % 2 == 1
        row_data = [project, client, bu, human_k(val)]
        aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
        for ci, (txt, aln) in enumerate(zip(row_data, aligns)):
            cell = tbl.cell(ri + 1, ci)
            _set_cell_text(cell, txt, font_size=Pt(10.5), align=aln)
            style_table_data(cell, alt=alt)
        total_val += val

    # Others row removed as requested

    # Unallocated row removed as requested

    # Total row
    total_row = n_rows - 1
    total_cells = ["", "TOTAL ORDER INTAKE", "", human_k(total_val)]
    total_aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
    for ci, (txt, aln) in enumerate(zip(total_cells, total_aligns)):
        cell = tbl.cell(total_row, ci)
        _set_cell_text(cell, txt, font_size=Pt(11), bold=True, align=aln)
        style_table_total(cell)

    print(
        f"  [Slide 6] OI table: {n_data} projects shown, total matches chart"
    )


def export_wip_to_excel(wip_rows, month, year, output_dir, slide_rows=None, negative_rows=None):
    """
    Export WIP detail table to an Excel file.
    Writes positive-WIP projects with slide markers, then a separate
    NEGATIVE WIP section at the bottom with subtotal and grand total.
    Returns the output file path.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    slide_names = {r["name"] for r in (slide_rows or [])}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"WIP {month} {year}"

    # ── Header row ─────────────────────────────────────────────
    headers = [
        "#",
        "Project Name",
        "Client",
        "Currency",
        "Total Invoice OC",
        "Total Prod OC",
        "WIP TL",
        "In Slide",
    ]
    hdr_fill = PatternFill("solid", fgColor="1E3A8A")
    hdr_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=10)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_side = Side(style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = border
    ws.row_dimensions[1].height = 28

    # ── Data rows ──────────────────────────────────────────────
    alt_fill = PatternFill("solid", fgColor="F4F6F9")
    data_font = Font(name="Segoe UI", size=9)
    num_fmt = "#,##0"

    for ri, proj in enumerate(wip_rows, 1):
        row_num = ri + 1
        on_slide = "YES" if proj["name"] in slide_names else ""
        row_vals = [
            ri,
            proj["name"],
            proj["client"],
            proj.get("orig_currency", ""),
            proj["inv_oc"],
            proj["prod_oc"],
            proj["wip_tl"],
            on_slide,
        ]
        fill = alt_fill if ri % 2 == 0 else None

        for ci, val in enumerate(row_vals, 1):
            cell = ws.cell(row=row_num, column=ci, value=val)
            cell.font = data_font
            cell.border = border
            cell.alignment = Alignment(
                horizontal="right" if ci >= 5 else "left",
                vertical="center",
                wrap_text=(ci == 2),
            )
            if fill:
                cell.fill = fill
            if ci in (5, 6, 7) and isinstance(val, (int, float)):
                cell.number_format = num_fmt

    # ── Total row ──────────────────────────────────────────────
    tot_row = len(wip_rows) + 2
    tot_font = Font(name="Segoe UI", bold=True, size=9)
    tot_fill = PatternFill("solid", fgColor="E2E8F0")
    totals = [
        "",
        "TOTAL",
        "",
        "",
        sum(r["inv_oc"] for r in wip_rows),
        sum(r["prod_oc"] for r in wip_rows),
        sum(r["wip_tl"] for r in wip_rows),
        "",
    ]
    for ci, val in enumerate(totals, 1):
        cell = ws.cell(row=tot_row, column=ci, value=val)
        cell.font = tot_font
        cell.fill = tot_fill
        cell.border = border
        cell.alignment = Alignment(
            horizontal="right" if ci >= 5 else "left", vertical="center"
        )
        if ci in (6, 7, 8) and isinstance(val, (int, float)):
            cell.number_format = num_fmt

    # ── Negative WIP section ───────────────────────────────────
    if negative_rows:
        neg_hdr_fill = PatternFill("solid", fgColor="E11D48")  # Deep Rose
        neg_hdr_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=10)
        neg_data_font = Font(name="Segoe UI", size=9, color="C0392B")
        neg_tot_font = Font(name="Segoe UI", bold=True, size=9, color="C0392B")
        neg_tot_fill = PatternFill("solid", fgColor="FFE4E8")
        neg_data_fill_odd = PatternFill("solid", fgColor="FFF0F3")
        grand_tot_fill = PatternFill("solid", fgColor="1E3A8A")
        grand_tot_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=10)

        # Blank separator row
        sep_row = tot_row + 1

        # Section header row
        neg_hdr_row = sep_row + 1
        neg_hdr_labels = ["", "NEGATIVE WIP TL PROJECTS", "", "", "", "", "", ""]
        for ci, val in enumerate(neg_hdr_labels, 1):
            cell = ws.cell(row=neg_hdr_row, column=ci, value=val)
            cell.font = neg_hdr_font
            cell.fill = neg_hdr_fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if ci == 2 else "left", vertical="center")
        ws.row_dimensions[neg_hdr_row].height = 22

        # Negative data rows
        for ni, proj in enumerate(negative_rows, 1):
            row_num = neg_hdr_row + ni
            row_vals = [
                ni,
                proj["name"],
                proj["client"],
                proj.get("orig_currency", ""),
                proj["inv_oc"],
                proj["prod_oc"],
                proj["wip_tl"],
                "",
            ]
            fill = neg_data_fill_odd if ni % 2 == 1 else None
            for ci, val in enumerate(row_vals, 1):
                cell = ws.cell(row=row_num, column=ci, value=val)
                cell.font = neg_data_font
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="right" if ci >= 5 else "left",
                    vertical="center",
                    wrap_text=(ci == 2),
                )
                if fill:
                    cell.fill = fill
                if ci in (5, 6, 7) and isinstance(val, (int, float)):
                    cell.number_format = num_fmt

        # Negative subtotal row
        neg_subtot_row = neg_hdr_row + len(negative_rows) + 1
        neg_subtotals = [
            "",
            "NEGATIVE SUBTOTAL",
            "",
            "",
            sum(r["inv_oc"] for r in negative_rows),
            sum(r["prod_oc"] for r in negative_rows),
            sum(r["wip_tl"] for r in negative_rows),
            "",
        ]
        for ci, val in enumerate(neg_subtotals, 1):
            cell = ws.cell(row=neg_subtot_row, column=ci, value=val)
            cell.font = neg_tot_font
            cell.fill = neg_tot_fill
            cell.border = border
            cell.alignment = Alignment(
                horizontal="right" if ci >= 5 else "left", vertical="center"
            )
            if ci in (5, 6, 7) and isinstance(val, (int, float)):
                cell.number_format = num_fmt

        # Grand total row (positive total + negative total)
        pos_wip_total = sum(r["wip_tl"] for r in wip_rows)
        neg_wip_total = sum(r["wip_tl"] for r in negative_rows)
        grand_tot_row = neg_subtot_row + 1
        grand_totals = [
            "",
            "GRAND TOTAL (incl. Negatives)",
            "",
            "",
            sum(r["inv_oc"] for r in wip_rows) + sum(r["inv_oc"] for r in negative_rows),
            sum(r["prod_oc"] for r in wip_rows) + sum(r["prod_oc"] for r in negative_rows),
            pos_wip_total + neg_wip_total,
            "",
        ]
        for ci, val in enumerate(grand_totals, 1):
            cell = ws.cell(row=grand_tot_row, column=ci, value=val)
            cell.font = grand_tot_font
            cell.fill = grand_tot_fill
            cell.border = border
            cell.alignment = Alignment(
                horizontal="right" if ci >= 5 else "left", vertical="center"
            )
            if ci in (5, 6, 7) and isinstance(val, (int, float)):
                cell.number_format = num_fmt
        ws.row_dimensions[grand_tot_row].height = 22

    # ── Column widths ──────────────────────────────────────────
    col_widths = [4, 48, 22, 12, 18, 18, 18, 9]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "C2"

    out_path = os.path.join(output_dir, f"WIP_Detail_{month}_{year}.xlsx")
    wb.save(out_path)
    print(f"  [WIP Excel] Saved: {os.path.basename(out_path)}  ({len(wip_rows)} rows, {len(negative_rows or [])} negatives)")
    return out_path


def embed_ole_excel_to_slide(
    pptx_path, xlsx_path, slide_xml_name, ole_x, ole_y, ole_cx, ole_cy
):
    """
    Post-process the PPTX zip to embed an Excel file as a double-clickable
    OLE object on a specific slide.
    ole_x/y/cx/cy are in EMU (integer).
    """
    import io, zipfile
    from lxml import etree
    from PIL import Image, ImageDraw

    # ── Build icon image displayed on the slide ─────────────────
    px_w = max(20, int(ole_cx / 914400 * 96))
    px_h = max(20, int(ole_cy / 914400 * 96))
    img = Image.new("RGB", (px_w, px_h), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, px_w - 1, px_h - 1], outline=(30, 58, 138), width=2)
    strip_w = max(8, min(px_w // 4, 52))
    draw.rectangle([2, 2, strip_w, px_h - 3], fill=(21, 128, 61))
    try:
        from PIL import ImageFont

        fnt_b = ImageFont.truetype("C:/Windows/Fonts/segoeuib.ttf", max(8, px_h // 4))
        fnt_s = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", max(7, px_h // 6))
    except Exception:
        fnt_b = fnt_s = None
    kw = {"font": fnt_b} if fnt_b else {}
    draw.text((strip_w // 2, px_h // 2), "XLS", fill=(255, 255, 255), **kw)
    kw_s = {"font": fnt_s} if fnt_s else {}
    fname = os.path.basename(xlsx_path)
    draw.text((strip_w + 6, max(4, px_h // 4 - 8)), fname, fill=(30, 58, 138), **kw_s)
    draw.text(
        (strip_w + 6, max(16, px_h - px_h // 3)),
        "Double-click to open",
        fill=(100, 116, 139),
        **kw_s,
    )
    icon_buf = io.BytesIO()
    img.save(icon_buf, format="PNG")
    icon_bytes = icon_buf.getvalue()

    # ── Internal zip paths ───────────────────────────────────────
    slide_path = f"ppt/slides/{slide_xml_name}"
    rels_path = f"ppt/slides/_rels/{slide_xml_name}.rels"
    embed_name = "ppt/embeddings/wip_detail.xlsx"
    icon_name = "ppt/media/wip_excel_icon.png"
    ct_path = "[Content_Types].xml"

    # ── XML namespaces ───────────────────────────────────────────
    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    OLE_RT = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
    )
    IMG_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    tmp_path = pptx_path + ".ole_tmp"
    with (
        zipfile.ZipFile(pptx_path, "r") as zin,
        zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout,
    ):
        names = zin.namelist()

        # ── Allocate new relationship IDs ────────────────────────
        rels_xml = zin.read(rels_path)
        rels_root = etree.fromstring(rels_xml)
        used_nums = [
            int(re.sub(r"\D", "", el.get("Id", "0")) or "0") for el in rels_root
        ]
        n = max(used_nums, default=0)
        rId_excel = f"rId{n + 1}"
        rId_icon = f"rId{n + 2}"
        etree.SubElement(
            rels_root,
            "Relationship",
            Id=rId_excel,
            Type=OLE_RT,
            Target="../embeddings/wip_detail.xlsx",
        )
        etree.SubElement(
            rels_root,
            "Relationship",
            Id=rId_icon,
            Type=IMG_RT,
            Target="../media/wip_excel_icon.png",
        )
        new_rels = etree.tostring(
            rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # ── Build graphicFrame element (OLE container) ───────────
        slide_xml = zin.read(slide_path)
        slide_root = etree.fromstring(slide_xml)
        spTree = slide_root.find(f".//{{{NS_P}}}spTree")

        # Add namespaces explicitly to avoid any validation errors
        NS_MAP = {"p": NS_P, "a": NS_A, "r": NS_R}

        gf = etree.SubElement(spTree, f"{{{NS_P}}}graphicFrame", nsmap=NS_MAP)
        nvGfPr = etree.SubElement(gf, f"{{{NS_P}}}nvGraphicFramePr")
        etree.SubElement(nvGfPr, f"{{{NS_P}}}cNvPr", id="9901", name="WIP Detail Excel")
        cNvGfPr = etree.SubElement(nvGfPr, f"{{{NS_P}}}cNvGraphicFramePr")
        etree.SubElement(cNvGfPr, f"{{{NS_A}}}graphicFrameLocks", noGrp="1")
        etree.SubElement(nvGfPr, f"{{{NS_P}}}nvPr")

        xfrm = etree.SubElement(gf, f"{{{NS_P}}}xfrm")
        etree.SubElement(xfrm, f"{{{NS_A}}}off", x=str(ole_x), y=str(ole_y))
        etree.SubElement(xfrm, f"{{{NS_A}}}ext", cx=str(ole_cx), cy=str(ole_cy))

        graphic = etree.SubElement(gf, f"{{{NS_A}}}graphic")
        gData = etree.SubElement(
            graphic,
            f"{{{NS_A}}}graphicData",
            uri="http://schemas.openxmlformats.org/presentationml/2006/ole",
        )

        oleObj = etree.SubElement(
            gData,
            f"{{{NS_P}}}oleObj",
            name="Worksheet",
            imgW=str(ole_cx),
            imgH=str(ole_cy),
            progId="Excel.Sheet.12",  # Restore progId
            showAsIcon="1",
        )
        oleObj.set(f"{{{NS_R}}}id", rId_excel)
        etree.SubElement(oleObj, f"{{{NS_P}}}embed")

        pic = etree.SubElement(oleObj, f"{{{NS_P}}}pic")
        nvPicPr = etree.SubElement(pic, f"{{{NS_P}}}nvPicPr")
        etree.SubElement(
            nvPicPr, f"{{{NS_P}}}cNvPr", id="0", name=""
        )  # Use id=0 as in ref
        etree.SubElement(nvPicPr, f"{{{NS_P}}}cNvPicPr")
        etree.SubElement(nvPicPr, f"{{{NS_P}}}nvPr")
        blipFill = etree.SubElement(pic, f"{{{NS_P}}}blipFill")
        blip = etree.SubElement(blipFill, f"{{{NS_A}}}blip")
        blip.set(f"{{{NS_R}}}embed", rId_icon)
        stretch = etree.SubElement(blipFill, f"{{{NS_A}}}stretch")
        etree.SubElement(stretch, f"{{{NS_A}}}fillRect")
        spPr = etree.SubElement(pic, f"{{{NS_P}}}spPr")
        xfrm2 = etree.SubElement(spPr, f"{{{NS_A}}}xfrm")
        etree.SubElement(xfrm2, f"{{{NS_A}}}off", x=str(ole_x), y=str(ole_y))
        etree.SubElement(xfrm2, f"{{{NS_A}}}ext", cx=str(ole_cx), cy=str(ole_cy))
        prstGeom = etree.SubElement(spPr, f"{{{NS_A}}}prstGeom", prst="rect")
        etree.SubElement(prstGeom, f"{{{NS_A}}}avLst")

        new_slide = etree.tostring(
            slide_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # ── Update [Content_Types].xml ───────────────────────────
        ct_xml = zin.read(ct_path)
        ct_root = etree.fromstring(ct_xml)
        tag_def = f"{{{CT_NS}}}Default"
        tag_ovr = f"{{{CT_NS}}}Override"
        ext_set = {el.get("Extension", "") for el in ct_root if el.tag == tag_def}
        parts_set = {el.get("PartName", "") for el in ct_root if el.tag == tag_ovr}
        if "png" not in ext_set:
            etree.SubElement(ct_root, tag_def, Extension="png", ContentType="image/png")
        embed_part = "/" + embed_name
        if embed_part not in parts_set:
            etree.SubElement(
                ct_root,
                tag_ovr,
                PartName=embed_part,
                ContentType="application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet",
            )
        new_ct = etree.tostring(
            ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # ── Copy all existing parts, replacing modified ones ─────
        for name in names:
            data = zin.read(name)
            if name == slide_path:
                data = new_slide
            elif name == rels_path:
                data = new_rels
            elif name == ct_path:
                data = new_ct
            zout.writestr(name, data)

        # ── Append new embedded parts ────────────────────────────
        with open(xlsx_path, "rb") as f:
            zout.writestr(embed_name, f.read())
        zout.writestr(icon_name, icon_bytes)

    os.replace(tmp_path, pptx_path)
    # Check if we actually succeeded in adding the files
    with zipfile.ZipFile(pptx_path, "r") as zcheck:
        zlist = zcheck.namelist()
        if embed_name in zlist and icon_name in zlist:
            print(
                f"  [OLE] Successfully embedded {os.path.basename(xlsx_path)}"
                f" in {slide_xml_name}"
            )
        else:
            print(
                f"  [OLE] ERROR: Files not found in ZIP after embedding!"
                f" (Missing: {[n for n in [embed_name, icon_name] if n not in zlist]})"
            )


def add_wip_table_to_slide(slide, wip_projects, slide_num=7):
    """
    WIP table for Slide 7 (and slide 8 if overflow).
    Columns: Project Name | Client | Original Currency | Total Invoice OC | Total Prod OC | WIP TL
    """
    rows_data = wip_projects
    n_data = len(rows_data)
    if n_data == 0:
        print(f"  [Slide {slide_num}] WIP table: no qualifying projects")
        return

    _wip_cfg = config.get("WIP_Table", {})
    _wip_cols = _wip_cfg.get(
        "Columns",
        [
            {
                "field": "name",
                "header": "Project Name",
                "width_pct": 0.32,
                "align": "LEFT",
            },
            {"field": "client", "header": "Client", "width_pct": 0.18, "align": "LEFT"},
            {
                "field": "orig_currency",
                "header": "Currency",
                "width_pct": 0.10,
                "align": "CENTER",
            },
            {
                "field": "inv_oc",
                "header": "Total Invoice OC",
                "width_pct": 0.15,
                "align": "RIGHT",
            },
            {
                "field": "prod_oc",
                "header": "Total Prod OC",
                "width_pct": 0.15,
                "align": "RIGHT",
            },
            {
                "field": "wip_tl",
                "header": "WIP TL",
                "width_pct": 0.10,
                "align": "RIGHT",
            },
        ],
    )
    _align_map = {
        "LEFT": PP_ALIGN.LEFT,
        "CENTER": PP_ALIGN.CENTER,
        "RIGHT": PP_ALIGN.RIGHT,
    }
    _numeric_fields = {"inv_oc", "prod_oc", "wip_tl"}

    _fs_cfg = _wip_cfg.get("Font_Size", {})
    _fs_header = Pt(_fs_cfg.get("Header", 10))
    _fs_data = Pt(_fs_cfg.get("Data", 9.5))
    _fs_total = Pt(_fs_cfg.get("Total", 10))

    from pptx.enum.text import MSO_ANCHOR
    n_cols = len(_wip_cols)
    n_rows = 1 + n_data + 1  # header + data + total

    left = Inches(0.45)
    top = Inches(1.38)
    width = Inches(12.43)
    # Row heights: header 0.46", data rows 0.44", total 0.46"
    _row_h_hdr  = Inches(0.46)
    _row_h_data = Inches(0.44)
    _row_h_tot  = Inches(0.46)
    height = _row_h_hdr + n_data * _row_h_data + _row_h_tot

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, int(height))
    tbl = tbl_shape.table

    # Explicit row heights for consistent spacing
    tbl.rows[0].height = int(_row_h_hdr)
    for ri in range(n_data):
        tbl.rows[ri + 1].height = int(_row_h_data)
    tbl.rows[n_rows - 1].height = int(_row_h_tot)

    for ci, col in enumerate(_wip_cols):
        tbl.columns[ci].width = int(width * col["width_pct"])

    def _vcenter(cell):
        cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Header row
    for ci, col in enumerate(_wip_cols):
        cell = tbl.cell(0, ci)
        aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
        _set_cell_text(cell, col["header"], font_size=_fs_header, bold=True, align=aln)
        style_table_header(cell)
        _vcenter(cell)
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = _fs_header

    # Data rows
    total_inv = total_prod = total_wip = 0.0
    for ri, proj in enumerate(rows_data):
        alt = ri % 2 == 1
        for ci, col in enumerate(_wip_cols):
            field = col["field"]
            raw = proj.get(field, 0 if field in _numeric_fields else "")
            txt = human_tl(raw) if field in _numeric_fields else str(raw)
            aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
            cell = tbl.cell(ri + 1, ci)
            _set_cell_text(cell, txt, font_size=_fs_data, align=aln)
            style_table_data(cell, alt=alt)
            _vcenter(cell)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = _fs_data
        total_inv += proj.get("inv_oc", 0)
        total_prod += proj.get("prod_oc", 0)
        total_wip += proj.get("wip_tl", 0)

    # Total row — fills only fields present in column list
    total_values = {
        "inv_oc": human_tl(total_inv),
        "prod_oc": human_tl(total_prod),
        "wip_tl": human_tl(total_wip),
    }
    tr = n_rows - 1
    for ci, col in enumerate(_wip_cols):
        field = col["field"]
        if field == "name":
            txt = f"TOP {n_data} TOTAL"
        else:
            txt = total_values.get(field, "")
        aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
        cell = tbl.cell(tr, ci)
        _set_cell_text(cell, txt, font_size=_fs_total, bold=True, align=aln)
        style_table_total(cell)
        _vcenter(cell)
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = _fs_total

    print(f"  [Slide 7] WIP table: {n_data} projects, WIP TL = {human_tl(total_wip)}")


def add_wip_negative_table_to_slide(slide, negative_rows, slide_total_wip, slide_num=8):
    """
    WIP Negative Projects table for its own dedicated slide.
    Uses the same column config as the positive WIP table.
    Appends a NEGATIVE SUBTOTAL row and a GRAND TOTAL (>=1M + Negatives) row.
    Negative WIP TL values are rendered in red.
    """
    from pptx.enum.text import MSO_ANCHOR
    from lxml import etree

    n_data = len(negative_rows)
    if n_data == 0:
        print(f"  [Slide {slide_num}] Negative WIP table: no negative projects")
        return

    _wip_cfg = config.get("WIP_Table", {})
    _wip_cols = _wip_cfg.get(
        "Columns",
        [
            {"field": "name",          "header": "Project Name",     "width_pct": 0.32, "align": "LEFT"},
            {"field": "client",        "header": "Client",           "width_pct": 0.18, "align": "LEFT"},
            {"field": "orig_currency", "header": "Currency",         "width_pct": 0.10, "align": "CENTER"},
            {"field": "inv_oc",        "header": "Total Invoice OC", "width_pct": 0.15, "align": "RIGHT"},
            {"field": "prod_oc",       "header": "Total Prod OC",    "width_pct": 0.15, "align": "RIGHT"},
            {"field": "wip_tl",        "header": "WIP TL",           "width_pct": 0.10, "align": "RIGHT"},
        ],
    )
    _align_map = {"LEFT": PP_ALIGN.LEFT, "CENTER": PP_ALIGN.CENTER, "RIGHT": PP_ALIGN.RIGHT}
    _numeric_fields = {"inv_oc", "prod_oc", "wip_tl"}

    _fs_cfg = _wip_cfg.get("Font_Size", {})
    _fs_header = Pt(_fs_cfg.get("Header", 10))
    _fs_data   = Pt(_fs_cfg.get("Data",   9.5))
    _fs_total  = Pt(_fs_cfg.get("Total",  10))

    n_cols = len(_wip_cols)
    # header + data rows + negative subtotal + grand total
    n_rows = 1 + n_data + 2

    left   = Inches(0.45)
    top    = Inches(1.38)
    width  = Inches(12.43)
    _row_h_hdr  = Inches(0.40)
    _row_h_data = Inches(0.37)
    _row_h_tot  = Inches(0.42)
    height = _row_h_hdr + n_data * _row_h_data + 2 * _row_h_tot

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, int(height))
    tbl = tbl_shape.table

    tbl.rows[0].height = int(_row_h_hdr)
    for ri in range(n_data):
        tbl.rows[ri + 1].height = int(_row_h_data)
    tbl.rows[n_rows - 2].height = int(_row_h_tot)
    tbl.rows[n_rows - 1].height = int(_row_h_tot)

    for ci, col in enumerate(_wip_cols):
        tbl.columns[ci].width = int(width * col["width_pct"])

    def _vcenter(cell):
        cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # ── Header row ──
    _header_red = RGBColor(0xE1, 0x1D, 0x48)  # deep rose red, matches Excel negative section
    for ci, col in enumerate(_wip_cols):
        cell = tbl.cell(0, ci)
        aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
        _set_cell_text(cell, col["header"], font_size=_fs_header, bold=True, align=aln)
        style_table_header(cell)
        cell.fill.fore_color.rgb = _header_red  # override ACCENT blue → red
        _vcenter(cell)
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = _fs_header

    # ── Data rows ──  colors match Excel export exactly
    _neg_red     = RGBColor(0xC0, 0x39, 0x2B)  # C0392B — all data text (matches Excel neg_data_font)
    _neg_red_bold = RGBColor(0xC0, 0x39, 0x2B)  # same, bold for WIP TL column
    _NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    def _set_cell_fill(cell, hex6):
        _tcPr = cell._tc.get_or_add_tcPr()
        sf = etree.SubElement(_tcPr, f"{_NS_A}solidFill")
        etree.SubElement(sf, f"{_NS_A}srgbClr").set("val", hex6)

    total_inv = total_prod = total_wip = 0.0
    for ri, proj in enumerate(negative_rows):
        odd = ri % 2 == 0  # 0-indexed: row 0,2,4… are "odd" in 1-indexed Excel terms
        for ci, col in enumerate(_wip_cols):
            field = col["field"]
            raw = proj.get(field, 0 if field in _numeric_fields else "")
            txt = human_tl(raw) if field in _numeric_fields else str(raw)
            aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
            cell = tbl.cell(ri + 1, ci)
            _set_cell_text(cell, txt, font_size=_fs_data, align=aln)
            if odd:
                _set_cell_fill(cell, "FFF0F3")  # matches Excel neg_data_fill_odd
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)  # white
            _vcenter(cell)
            is_wip = field == "wip_tl"
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = _neg_red
                    run.font.bold = is_wip
                    run.font.size = _fs_data
        total_inv  += proj.get("inv_oc", 0)
        total_prod += proj.get("prod_oc", 0)
        total_wip  += proj.get("wip_tl", 0)

    # ── Negative subtotal row ──  FFE4E8 fill, bold C0392B text (matches Excel)
    neg_subtot_ri = n_rows - 2
    neg_subtot_values = {
        "name":    f"NEGATIVE SUBTOTAL  ({n_data} project{'s' if n_data != 1 else ''})",
        "inv_oc":  human_tl(total_inv),
        "prod_oc": human_tl(total_prod),
        "wip_tl":  human_tl(total_wip),
    }
    for ci, col in enumerate(_wip_cols):
        field = col["field"]
        txt = neg_subtot_values.get(field, "")
        aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
        cell = tbl.cell(neg_subtot_ri, ci)
        _set_cell_text(cell, txt, font_size=_fs_total, bold=True, align=aln)
        _set_cell_fill(cell, "FFE4E8")  # matches Excel neg_tot_fill
        _vcenter(cell)
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.color.rgb = _neg_red
                run.font.bold = True
                run.font.size = _fs_total

    # ── Grand total row (>=1M positives + negatives) ──
    grand_total_wip = slide_total_wip + total_wip
    grand_tot_ri = n_rows - 1
    grand_tot_values = {
        "name":   f"GRAND TOTAL  (\u22651M on slides + Negatives)",
        "wip_tl": human_tl(grand_total_wip),
    }
    _grand_dark = RGBColor(0x1E, 0x3A, 0x8A)  # Deep Sapphire
    for ci, col in enumerate(_wip_cols):
        field = col["field"]
        txt = grand_tot_values.get(field, "")
        aln = _align_map.get(col["align"], PP_ALIGN.LEFT)
        cell = tbl.cell(grand_tot_ri, ci)
        _set_cell_text(cell, txt, font_size=_fs_total, bold=True, align=aln)
        _set_cell_fill(cell, "1E3A8A")  # matches Excel grand_tot_fill
        _vcenter(cell)
        for para in cell.text_frame.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.bold = True
                run.font.size = _fs_total

    print(
        f"  [Slide {slide_num}] Negative WIP table: {n_data} projects, "
        f"subtotal={human_tl(total_wip)}, grand total={human_tl(grand_total_wip)}"
    )


# ──────────────────────────────────────────────────────────────
# MAIN MBR FUNCTION
# ──────────────────────────────────────────────────────────────
def create_17_slide_mbr_stacked(excel_file, output_ppt=None):
    # ── Derive month / year from filename ─────────────────────
    month = "January"
    year = "2026"
    months_list = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ]
    for m in months_list:
        if m[:3].lower() in os.path.basename(excel_file).lower():
            month = m
            break
    m_year = re.search(r"^(\d{2})", os.path.basename(excel_file))
    if m_year:
        year = "20" + m_year.group(1)
    if output_ppt is None:
        output_ppt = f"MRC_MBR_Stacked_{month}_{year}.pptx"

    month_idx = months_list.index(month) if month in months_list else 0

    print(f"[Stacked] Reading {excel_file} …  ({month} {year})")

    # ── Load sheets ───────────────────────────────────────────
    try:
        df = pd.read_excel(excel_file, sheet_name="Summary TL", header=None)
    except Exception as e:
        print(f"Cannot read 'Summary TL': {e}")
        return

    try:
        df_oi = pd.read_excel(excel_file, sheet_name="Order Intake", header=None)
    except Exception:
        df_oi = None

    try:
        df_wip = pd.read_excel(excel_file, sheet_name="WIP & BL", header=None)
    except Exception:
        df_wip = None

    try:
        df_gm = pd.read_excel(excel_file, sheet_name="Gross Margin", header=None)
    except Exception:
        df_gm = None

    # ─────────────────────────────────────────────────────────
    # DATA MAP
    # ─────────────────────────────────────────────────────────
    def rv(row, c_start, c_end):
        return [safe_float(df.iloc[row, c]) for c in range(c_start, c_end)]

    R = config.get("Excel_Mapping.Rows", {})
    C = config.get("Excel_Mapping.Columns", {})
    O = config.get("Excel_Mapping.Offsets", {})

    cats_ns_ktl = [
        clean_label(df.iloc[R.get("Categories_NS_kTL", 14), c]) for c in range(C.get("NS_kTL_Start", 2), C.get("NS_kTL_End", 17))
    ]
    cats_ns_keur = [
        clean_label(df.iloc[R.get("Categories_NS_kEUR", 47), c]) for c in range(C.get("NS_kEUR_Start", 2), C.get("NS_kEUR_End", 17))
    ]
    cats_ebit_ktl = [
        clean_label(df.iloc[R.get("Categories_EBIT_kTL", 36), c]) for c in range(C.get("EBIT_kTL_Start", 1), C.get("EBIT_kTL_End", 16))
    ]
    cats_ebit_keur = [
        clean_label(df.iloc[R.get("Categories_EBIT_kEUR", 55), c]) for c in range(C.get("EBIT_kEUR_Start", 1), C.get("EBIT_kEUR_End", 16))
    ]
    # OI sheet categories row contains numeric values, not month labels.
    # Derive from main sheet NS labels: drop the "BL" entry (index 2) so
    # the 14 OI columns align as [2025, 2026 Target, Jan … Dec].
    cats_oi = cats_ns_ktl[:2] + cats_ns_ktl[3:]
    cats_bu_ns = [
        clean_label(df.iloc[R.get("Categories_BU_NS", 14), c]) for c in range(C.get("BU_NS_Start", 52), C.get("BU_NS_End", 66))
    ]
    cats_bu_ebit = [
        clean_label(df.iloc[R.get("Categories_BU_EBIT", 97), c]) for c in range(C.get("BU_EBIT_Start", 52), C.get("BU_EBIT_End", 66))
    ]

    s2_contract = rv(R.get("Slide2_Contract", 15), C.get("NS_kTL_Start", 2), C.get("NS_kTL_End", 17))
    s2_wp = rv(R.get("Slide2_WP", 16), C.get("NS_kTL_Start", 2), C.get("NS_kTL_End", 17))
    s2_wo = rv(R.get("Slide2_WO", 17), C.get("NS_kTL_Start", 2), C.get("NS_kTL_End", 17))

    s3_contract = rv(R.get("Slide3_Contract", 48), C.get("NS_kEUR_Start", 2), C.get("NS_kEUR_End", 17))
    s3_wp = rv(R.get("Slide3_WP", 49), C.get("NS_kEUR_Start", 2), C.get("NS_kEUR_End", 17))
    s3_wo = rv(R.get("Slide3_WO", 50), C.get("NS_kEUR_Start", 2), C.get("NS_kEUR_End", 17))

    s4_contract = rv(R.get("Slide4_Contract", 37), C.get("EBIT_kTL_Start", 1), C.get("EBIT_kTL_End", 16))
    s4_cwp = rv(R.get("Slide4_Contract_WP", 38), C.get("EBIT_kTL_Start", 1), C.get("EBIT_kTL_End", 16))
    s4_cwp_wo = rv(R.get("Slide4_Contract_WP_WO", 39), C.get("EBIT_kTL_Start", 1), C.get("EBIT_kTL_End", 16))

    s5_contract = rv(R.get("Slide5_Contract", 56), C.get("EBIT_kEUR_Start", 1), C.get("EBIT_kEUR_End", 16))
    s5_cwp = rv(R.get("Slide5_Contract_WP", 57), C.get("EBIT_kEUR_Start", 1), C.get("EBIT_kEUR_End", 16))
    s5_cwp_wo = rv(R.get("Slide5_Contract_WP_WO", 58), C.get("EBIT_kEUR_Start", 1), C.get("EBIT_kEUR_End", 16))

    # Dynamically find the OI categories row to handle Excel files that have
    # a different number of rows (e.g. January file has 1 extra row vs February).
    # The categories row is identified by having a year value (2025) at OI_kTL_Start
    # and "Jan" at OI_kTL_Start+2.
    _oi_col_start = C.get("OI_kTL_Start", 4)
    _oi_col_end   = C.get("OI_kTL_End", 18)
    _oi_cats_cfg  = R.get("Categories_OI_kTL", 543)
    _oi_cats_row  = _oi_cats_cfg  # default
    for _probe in range(_oi_cats_cfg - 2, _oi_cats_cfg + 5):
        if 0 <= _probe < len(df_oi):
            _c4 = str(df_oi.iloc[_probe, _oi_col_start]).strip()
            _c6 = str(df_oi.iloc[_probe, _oi_col_start + 2]).strip()
            # Categories row: col4 looks like a year (e.g. "2025"), col6 is "Jan"
            if _c4.split(".")[0].isdigit() and len(_c4.split(".")[0]) == 4 and "jan" in _c6.lower():
                _oi_cats_row = _probe
                break
    _oi_offset = _oi_cats_row - _oi_cats_cfg  # 0 for most files, ±N if file has extra/fewer rows
    if _oi_offset != 0:
        print(f"[OI] Row offset detected: {_oi_offset:+d} (file row count differs from config baseline)")

    def _oi_row(key, default):
        return [safe_float(df_oi.iloc[R.get(key, default) + _oi_offset, c])
                for c in range(_oi_col_start, _oi_col_end)]

    s6_eng = _oi_row("Slide6_ENG", 544)
    s6_mc  = _oi_row("Slide6_MC",  545)
    s6_tsi = _oi_row("Slide6_TSI", 546)
    s6_nuc = _oi_row("Slide6_NUC", 547)

    def bu_ns_row(xrow):
        return [safe_float(df.iloc[xrow - 1, c]) for c in range(C.get("BU_NS_Start", 52), C.get("BU_NS_End", 66))]

    bu_ns = {
        "ENG": {
            "Order": bu_ns_row(R.get("BU_NS_ENG_Order", 19)),
            "Offer": bu_ns_row(R.get("BU_NS_ENG_Offer", 20)),
            "Opp": bu_ns_row(R.get("BU_NS_ENG_Opp", 21)),
        },
        "MC": {
            "Order": bu_ns_row(R.get("BU_NS_MC_Order", 16)),
            "Offer": bu_ns_row(R.get("BU_NS_MC_Offer", 17)),
            "Opp": bu_ns_row(R.get("BU_NS_MC_Opp", 18)),
        },
        "T&SI": {
            "Order": bu_ns_row(R.get("BU_NS_TSI_Order", 22)),
            "Offer": bu_ns_row(R.get("BU_NS_TSI_Offer", 23)),
            "Opp": bu_ns_row(R.get("BU_NS_TSI_Opp", 24)),
        },
        "NUC": {
            "Order": bu_ns_row(R.get("BU_NS_NUC_Order", 25)),
            "Offer": bu_ns_row(R.get("BU_NS_NUC_Offer", 26)),
            "Opp": bu_ns_row(R.get("BU_NS_NUC_Opp", 27)),
        },
    }

    def bu_ebit_row(xrow):
        return [safe_float(df.iloc[xrow - 1, c]) for c in range(C.get("BU_EBIT_Start", 52), C.get("BU_EBIT_End", 66))]

    bu_ebit = {
        "ENG": bu_ebit_row(R.get("BU_EBIT_ENG", 102)),
        "MC": bu_ebit_row(R.get("BU_EBIT_MC", 99)),
        "T&SI": bu_ebit_row(R.get("BU_EBIT_TSI", 104)),
        "NUC": bu_ebit_row(R.get("BU_EBIT_NUC", 106)),
    }

    # ── Dynamic month offsets ─────────────────────────────────
    MON_NS = O.get("MON_NS_Base", 3) + month_idx
    MON_EBIT = O.get("MON_EBIT_Base", 3) + month_idx
    # BU NS/EBIT arrays (col52-65, 14 values): [2025, 2026Target, Jan…Dec] — Jan=index 2
    MON_BU = O.get("MON_BU_Base", 2) + month_idx
    MON_OI = O.get("MON_OI_Base", 2) + month_idx

    # ── Callout helpers: locate month column by scanning header row ────────
    def _find_month_col(hdr_row_idx, col_start, col_end, mon_name):
        """Return the pandas column index whose header cell matches mon_name."""
        target = mon_name[:3].lower()
        for ci in range(col_start, col_end):
            cell = str(df.iloc[hdr_row_idx, ci]).strip().lower()
            if cell == target or cell.startswith(target):
                return ci
        return None

    def _col_sum(row_indices, col_idx):
        """Sum a list of pandas row indices at a given column index."""
        return sum(safe_float(df.iloc[r, col_idx]) for r in row_indices)

    def _find_oi_month_col(hdr_row_idx, col_start, col_end, mon_name):
        """Like _find_month_col but scans the Order Intake sheet (df_oi)."""
        if df_oi is None:
            return None
        target = mon_name[:3].lower()
        for ci in range(col_start, col_end):
            cell = str(df_oi.iloc[hdr_row_idx, ci]).strip().lower()
            if cell == target or cell.startswith(target):
                return ci
        return None

    # ── NS kTL callout — 'Summary TL'!$C$15:$Q$18 ────────────────────────
    # Row 15 = header (2023/2024/Target/BL/Jan…Dec); rows 16-18 = Contract/WP/WO
    _ns_ktl_hdr  = R.get("Categories_NS_kTL", 14)
    _ns_ktl_cs   = C.get("NS_kTL_Start", 2)
    _ns_ktl_ce   = C.get("NS_kTL_End", 17)
    _ns_ktl_rows = [R.get("Slide2_Contract", 15), R.get("Slide2_WP", 16), R.get("Slide2_WO", 17)]
    _ns_ktl_col  = _find_month_col(_ns_ktl_hdr, _ns_ktl_cs, _ns_ktl_ce, month)
    _ns_ktl_jan  = _find_month_col(_ns_ktl_hdr, _ns_ktl_cs, _ns_ktl_ce, "Jan")
    _ns_ktl_dec  = _find_month_col(_ns_ktl_hdr, _ns_ktl_cs, _ns_ktl_ce, "Dec")
    if _ns_ktl_col is not None:
        _ktl_curr = _col_sum(_ns_ktl_rows, _ns_ktl_col)
        _ktl_prev = _col_sum(_ns_ktl_rows, _ns_ktl_col - 1) if month_idx > 0 else 0
        # YTD: sum of singular monthly values Jan → current month
        if _ns_ktl_jan is not None:
            _prev_sum, gr_ytd_ktl = 0, 0
            for _ci in range(_ns_ktl_jan, _ns_ktl_col + 1):
                _s = _col_sum(_ns_ktl_rows, _ci)
                gr_ytd_ktl += _s - _prev_sum
                _prev_sum = _s
        else:
            gr_ytd_ktl = _ktl_curr
        # December total: year-end value (sum of all singular monthly values)
        gr_dec_ktl = _col_sum(_ns_ktl_rows, _ns_ktl_dec) if _ns_ktl_dec is not None else gr_ytd_ktl
    else:  # fallback to offset-based index
        _ktl_curr = s2_contract[MON_NS] + s2_wp[MON_NS] + s2_wo[MON_NS]
        _ktl_prev = (s2_contract[MON_NS - 1] + s2_wp[MON_NS - 1] + s2_wo[MON_NS - 1]) if month_idx > 0 else 0
        gr_ytd_ktl = _ktl_curr
        gr_dec_ktl = _ktl_curr
    gr_mon_ktl = _ktl_curr - _ktl_prev   # singular monthly value

    # ── NS kEUR callout — 'Summary TL'!$C$48:$Q$51 ───────────────────────
    # Row 48 = header; rows 49-51 = Contract/WP/WO (same delta logic as kTL)
    _ns_keur_hdr  = R.get("Categories_NS_kEUR", 47)
    _ns_keur_cs   = C.get("NS_kEUR_Start", 2)
    _ns_keur_ce   = C.get("NS_kEUR_End", 17)
    _ns_keur_rows = [R.get("Slide3_Contract", 48), R.get("Slide3_WP", 49), R.get("Slide3_WO", 50)]
    _ns_keur_col  = _find_month_col(_ns_keur_hdr, _ns_keur_cs, _ns_keur_ce, month)
    _ns_keur_jan  = _find_month_col(_ns_keur_hdr, _ns_keur_cs, _ns_keur_ce, "Jan")
    _ns_keur_dec  = _find_month_col(_ns_keur_hdr, _ns_keur_cs, _ns_keur_ce, "Dec")
    if _ns_keur_col is not None:
        _keur_curr = _col_sum(_ns_keur_rows, _ns_keur_col)
        _keur_prev = _col_sum(_ns_keur_rows, _ns_keur_col - 1) if month_idx > 0 else 0
        if _ns_keur_jan is not None:
            _prev_sum, gr_ytd_keur = 0, 0
            for _ci in range(_ns_keur_jan, _ns_keur_col + 1):
                _s = _col_sum(_ns_keur_rows, _ci)
                gr_ytd_keur += _s - _prev_sum
                _prev_sum = _s
        else:
            gr_ytd_keur = _keur_curr
        gr_dec_keur = _col_sum(_ns_keur_rows, _ns_keur_dec) if _ns_keur_dec is not None else gr_ytd_keur
    else:
        _keur_curr = s3_contract[MON_NS] + s3_wp[MON_NS] + s3_wo[MON_NS]
        _keur_prev = (s3_contract[MON_NS - 1] + s3_wp[MON_NS - 1] + s3_wo[MON_NS - 1]) if month_idx > 0 else 0
        gr_ytd_keur = _keur_curr
        gr_dec_keur = _keur_curr
    gr_mon_keur = _keur_curr - _keur_prev

    # ── EBIT kTL callout — 'Summary TL'!$A$37:$P$40 ──────────────────────
    # Row 37 = header; row 40 (pandas 39) = Contract+WP+WO combined total.
    # Use only that single row; subtract previous month for singular value.
    _ebit_ktl_hdr  = R.get("Categories_EBIT_kTL", 36)
    _ebit_ktl_cs   = C.get("EBIT_kTL_Start", 1)
    _ebit_ktl_ce   = C.get("EBIT_kTL_End", 17)
    _ebit_ktl_row  = R.get("Slide4_Contract_WP_WO", 39)   # row 40 in Excel
    _ebit_ktl_col  = _find_month_col(_ebit_ktl_hdr, _ebit_ktl_cs, _ebit_ktl_ce, month)
    _ebit_ktl_jan  = _find_month_col(_ebit_ktl_hdr, _ebit_ktl_cs, _ebit_ktl_ce, "Jan")
    _ebit_ktl_dec  = _find_month_col(_ebit_ktl_hdr, _ebit_ktl_cs, _ebit_ktl_ce, "Dec")
    if _ebit_ktl_col is not None:
        _ebit_ktl_curr = safe_float(df.iloc[_ebit_ktl_row, _ebit_ktl_col])
        _ebit_ktl_prev = safe_float(df.iloc[_ebit_ktl_row, _ebit_ktl_col - 1]) if month_idx > 0 else 0
        # YTD: sum of singular monthly values Jan → current month
        if _ebit_ktl_jan is not None:
            _prev_val, ebit_ytd_ktl = 0, 0
            for _ci in range(_ebit_ktl_jan, _ebit_ktl_col + 1):
                _v = safe_float(df.iloc[_ebit_ktl_row, _ci])
                ebit_ytd_ktl += _v - _prev_val
                _prev_val = _v
        else:
            ebit_ytd_ktl = _ebit_ktl_curr
        # December total: year-end value
        ebit_dec_ktl = safe_float(df.iloc[_ebit_ktl_row, _ebit_ktl_dec]) if _ebit_ktl_dec is not None else ebit_ytd_ktl
    else:
        _ebit_ktl_curr = s4_cwp_wo[MON_EBIT]
        _ebit_ktl_prev = s4_cwp_wo[MON_EBIT - 1] if month_idx > 0 else 0
        ebit_ytd_ktl = _ebit_ktl_curr
        ebit_dec_ktl = _ebit_ktl_curr
    ebit_mon_ktl = _ebit_ktl_curr - _ebit_ktl_prev

    # ── EBIT kEUR callout — 'Summary TL'!$A$56:$P$59 ─────────────────────
    # Row 56 = header; row 59 (pandas 58) = Contract+WP+WO combined total.
    _ebit_keur_hdr  = R.get("Categories_EBIT_kEUR", 55)
    _ebit_keur_cs   = C.get("EBIT_kEUR_Start", 1)
    _ebit_keur_ce   = C.get("EBIT_kEUR_End", 17)
    _ebit_keur_row  = R.get("Slide5_Contract_WP_WO", 58)   # row 59 in Excel
    _ebit_keur_col  = _find_month_col(_ebit_keur_hdr, _ebit_keur_cs, _ebit_keur_ce, month)
    _ebit_keur_jan  = _find_month_col(_ebit_keur_hdr, _ebit_keur_cs, _ebit_keur_ce, "Jan")
    _ebit_keur_dec  = _find_month_col(_ebit_keur_hdr, _ebit_keur_cs, _ebit_keur_ce, "Dec")
    if _ebit_keur_col is not None:
        _ebit_keur_curr = safe_float(df.iloc[_ebit_keur_row, _ebit_keur_col])
        _ebit_keur_prev = safe_float(df.iloc[_ebit_keur_row, _ebit_keur_col - 1]) if month_idx > 0 else 0
        if _ebit_keur_jan is not None:
            _prev_val, ebit_ytd_keur = 0, 0
            for _ci in range(_ebit_keur_jan, _ebit_keur_col + 1):
                _v = safe_float(df.iloc[_ebit_keur_row, _ci])
                ebit_ytd_keur += _v - _prev_val
                _prev_val = _v
        else:
            ebit_ytd_keur = _ebit_keur_curr
        ebit_dec_keur = safe_float(df.iloc[_ebit_keur_row, _ebit_keur_dec]) if _ebit_keur_dec is not None else ebit_ytd_keur
    else:
        _ebit_keur_curr = s5_cwp_wo[MON_EBIT]
        _ebit_keur_prev = s5_cwp_wo[MON_EBIT - 1] if month_idx > 0 else 0
        ebit_ytd_keur = _ebit_keur_curr
        ebit_dec_keur = _ebit_keur_curr
    ebit_mon_keur = _ebit_keur_curr - _ebit_keur_prev

    # ── Retained for internal EBIT % computation ──────────────────────────
    gr_total_ktl   = s2_contract[14] + s2_wp[14] + s2_wo[14]
    ebit_total_ktl = sum(s4_cwp_wo[4:16])
    ebit_total_pct = ebit_total_ktl / gr_total_ktl if gr_total_ktl else 0

    # ── Order Intake ───────────────────────────────────────────────────────
    oi_mon_total = s6_eng[MON_OI] + s6_mc[MON_OI] + s6_tsi[MON_OI] + s6_nuc[MON_OI]
    oi_total = (
        sum(s6_eng[2:14]) + sum(s6_mc[2:14]) + sum(s6_tsi[2:14]) + sum(s6_nuc[2:14])
    )
    # OI data is non-cumulative (each cell = that month's value), so sum manually for YTD
    oi_ytd = (
        sum(s6_eng[2:MON_OI + 1]) + sum(s6_mc[2:MON_OI + 1])
        + sum(s6_tsi[2:MON_OI + 1]) + sum(s6_nuc[2:MON_OI + 1])
    )

    # ── OI callout: dynamic month-column scan on row 540 (pandas 539) ────────
    # 'Order Intake'!$D$544:$R$548 — month headers at row 540, cols G(6) to R(18)
    _oi_hdr_row = R.get("OI_Month_Header_Row", 539)  # Excel row 540 = pandas 539
    _oi_mon_col = _find_oi_month_col(_oi_hdr_row, 6, 18, month)
    if _oi_mon_col is not None and df_oi is not None:
        oi_callout_mon = sum(
            safe_float(df_oi.iloc[R.get(k, d) + _oi_offset, _oi_mon_col])
            for k, d in [
                ("Slide6_ENG", 544), ("Slide6_MC", 545),
                ("Slide6_TSI", 546), ("Slide6_NUC", 547),
            ]
        )
    else:
        oi_callout_mon = oi_mon_total  # fallback

    # ── BU NS callout: singular monthly values (cumulative delta) ────────────
    # Header row 15 (pandas 14), cols AZ:BN — same 'Summary TL' header as main NS
    _bu_ns_hdr = R.get("Categories_BU_NS", 14)
    _bu_ns_cs  = C.get("BU_NS_Start", 52)
    _bu_ns_ce  = C.get("BU_NS_End", 66)
    _bu_ns_col = _find_month_col(_bu_ns_hdr, _bu_ns_cs, _bu_ns_ce, month)

    def _bu_ns_sum(excel_rows, col_idx):
        """Sum Excel rows (1-indexed) at col_idx from Summary TL."""
        return sum(safe_float(df.iloc[r - 1, col_idx]) for r in excel_rows)

    _bu_ns_cfg = {
        "ENG":  [R.get("BU_NS_ENG_Order", 19), R.get("BU_NS_ENG_Offer", 20), R.get("BU_NS_ENG_Opp", 21)],
        "MC":   [R.get("BU_NS_MC_Order",   16), R.get("BU_NS_MC_Offer",  17), R.get("BU_NS_MC_Opp",  18)],
        "T&SI": [R.get("BU_NS_TSI_Order",  22), R.get("BU_NS_TSI_Offer", 23), R.get("BU_NS_TSI_Opp", 24)],
        "NUC":  [R.get("BU_NS_NUC_Order",  25), R.get("BU_NS_NUC_Offer", 26), R.get("BU_NS_NUC_Opp", 27)],
    }
    bu_ns_singular = {}
    for _bk, _rows in _bu_ns_cfg.items():
        if _bu_ns_col is not None:
            _bc = _bu_ns_sum(_rows, _bu_ns_col)
            _bp = _bu_ns_sum(_rows, _bu_ns_col - 1) if month_idx > 0 else 0
        else:  # fallback to pre-read arrays
            _bc = bu_ns[_bk]["Order"][MON_BU] + bu_ns[_bk]["Offer"][MON_BU] + bu_ns[_bk]["Opp"][MON_BU]
            _bp = (
                bu_ns[_bk]["Order"][MON_BU - 1] + bu_ns[_bk]["Offer"][MON_BU - 1] + bu_ns[_bk]["Opp"][MON_BU - 1]
            ) if month_idx > 0 else 0
        bu_ns_singular[_bk] = _bc - _bp

    # ── ABNS Extraction from Gross Margin sheet ────────────────────
    oi_abns_projects = []
    if df_gm is not None:
        GM_STATUS = C.get("GM_STATUS", 6)
        GM_ITEM   = C.get("GM_ITEM", 7)
        GM_PROJ   = C.get("GM_PROJ", 4)
        GM_CLIENT = C.get("GM_CLIENT", 2)
        GM_TOTAL  = C.get("GM_TOTAL", 25)

        for idx in range(4, df_gm.shape[0]):
            row = df_gm.iloc[idx]
            status = str(row[GM_STATUS]).strip().upper() if pd.notna(row[GM_STATUS]) else ""
            item   = str(row[GM_ITEM]).strip().upper()   if pd.notna(row[GM_ITEM])   else ""
            if status != "ABNS" or item != "GR":
                continue
            val = safe_float(row[GM_TOTAL])
            if val <= 0:
                continue
            project = str(row[GM_PROJ]).strip()  if pd.notna(row[GM_PROJ])   else ""
            client  = str(row[GM_CLIENT]).strip() if pd.notna(row[GM_CLIENT]) else ""
            if not project or project.lower() in ("nan", "0", ""):
                continue
            oi_abns_projects.append((project, client, val))

    oi_abns_projects.sort(key=lambda x: x[2], reverse=True)
    oi_mon_abns = sum(x[2] for x in oi_abns_projects)

    # Show ALL ABNS projects (no top-N limit)
    oi_abns_all = oi_abns_projects

    # BU NS arrays: [0]=2025, [1]=2026Target, [2]=Jan, …, [13]=Dec (14 values, col52-65)
    # Use December cumulative (index 13) as the FY 2026 total.
    bu_annual_ns = {
        bu: bu_ns[bu]["Order"][13] + bu_ns[bu]["Offer"][13]
        for bu in bu_ns
    }
    # BU EBIT arrays (col52-65, 14 values): [0]=2025, [1]=2026Target, [2]=Jan, …, [13]=Dec
    bu_mon_ebit = {bu: bu_ebit[bu][MON_BU] for bu in bu_ebit}

    # ── EBIT Percentages Extraction ───────────────────────────
    ebit_p_ktl  = rv(R.get("EBIT_P_kTL", 42), C.get("EBIT_kTL_Start", 1), C.get("EBIT_kTL_End", 16))
    ebit_p_keur = rv(R.get("EBIT_P_kTL", 42), C.get("EBIT_kEUR_Start", 1), C.get("EBIT_kEUR_End", 16))

    # ── Order Intake projects ─────────────────────────────────
    oi_projects = []
    if df_oi is not None:
        oi_col = O.get("OI_Col_Base", 6) + month_idx
        oi_data_end = R.get("Categories_OI_kTL", 543)

        for idx in range(2, oi_data_end):
            row = df_oi.iloc[idx]
            val = safe_float(row[oi_col])
            if val <= 0:
                continue
            
            project_col = C.get("OI_PROJ", 3)
            client_col = C.get("OI_CLIENT", 2)
            bu_col = C.get("OI_BU", 4)
            
            project = str(row[project_col]).strip() if pd.notna(row[project_col]) else ""
            client = str(row[client_col]).strip() if pd.notna(row[client_col]) else ""
            bu = str(row[bu_col]).strip() if pd.notna(row[bu_col]) else ""
            if not project or project.lower() == "nan" or project.strip() == "-":
                continue
            oi_projects.append((project, client, val, bu))
    oi_projects.sort(key=lambda x: x[2], reverse=True)

    # ── WIP & Backlog projects (GROSS FEES only, excl. Nuclear BU, WIP >= 1M) ──────
    wip_slide_rows = []   # All projects with WIP >= 1M → PPTX table(s)
    wip_excel_rows = []   # Positive projects → attached Excel
    wip_negative_rows = []  # Projects with negative WIP TL (after filters)

    if df_wip is not None:
        C_CODE = C.get("WIP_Col_Code", 1)
        C_NAME = C.get("WIP_Col_Name", 2)
        C_TYPE = C.get("WIP_Col_Type", 3)
        C_CLIENT = C.get("WIP_Col_Client", 4)
        C_BU = C.get("WIP_Col_BU", 10)
        C_ORIG_CURR = C.get("WIP_Col_OrigCurrency", 11)
        C_INV = C.get("WIP_Col_TotalInvoiceOC", 93)
        C_PROD = C.get("WIP_Col_TotalProdOC", 94)
        C_WIP = C.get("WIP_Col_WIP_TL", 98)
        START = R.get("WIP_DataStartRow", 3)
        MAX_ROWS_PER_SLIDE = config.get("WIP_Table.WIP_Slide_MaxRows", 10)
        MIN_TL = config.get("WIP_Table.WIP_Min_TL", 1000000)

        wip_excel_rows = []  # ALL projects (including < 1M, excluding negatives)
        wip_slide_rows = []  # Only projects >= 1M for slides
        wip_negative_rows = []  # Projects with negative WIP TL (after filters)

        for idx in range(START, len(df_wip)):
            row = df_wip.iloc[idx]

            if str(row[C_TYPE]).strip() != "GROSS FEES":
                continue

            bu = str(row[C_BU]).strip() if pd.notna(row[C_BU]) else ""
            if bu == "Nuclear":
                continue

            client = str(row[C_CLIENT]).strip() if pd.notna(row[C_CLIENT]) else ""
            if client.lower() == "nan":
                client = ""

            wip_tl = safe_float(row[C_WIP])

            orig_curr = (
                str(row[C_ORIG_CURR]).strip() if pd.notna(row[C_ORIG_CURR]) else ""
            )
            if orig_curr.lower() == "nan":
                orig_curr = ""

            entry = {
                "name": str(row[C_NAME]).strip() if pd.notna(row[C_NAME]) else "",
                "client": client,
                "orig_currency": orig_curr,
                "inv_oc": safe_float(row[C_INV]),
                "prod_oc": safe_float(row[C_PROD]),
                "wip_tl": wip_tl,
            }

            if wip_tl < -1000:
                wip_negative_rows.append(entry)  # Negatives tracked separately
                continue

            if wip_tl <= 0:  # skip zeros and near-zero negatives (> -1000 TL)
                continue

            wip_excel_rows.append(entry)  # All positive projects go to Excel

            if wip_tl >= MIN_TL:
                wip_slide_rows.append(entry)  # Only >= 1M go to slides

        # Sort by WIP TL descending; negatives ascending (most negative first)
        wip_excel_rows = sorted(wip_excel_rows, key=lambda x: x["wip_tl"], reverse=True)
        wip_slide_rows = sorted(wip_slide_rows, key=lambda x: x["wip_tl"], reverse=True)
        wip_negative_rows = sorted(wip_negative_rows, key=lambda x: x["wip_tl"])

        print(
            f"  [WIP] {len(wip_slide_rows)} projects >= 1M TL on slides, "
            f"{len(wip_excel_rows)} total in Excel export (incl. < 1M), "
            f"{len(wip_negative_rows)} negative WIP projects"
        )

    # ── Generate charts ───────────────────────────────────────
    chart_paths = []

    def stacked_chart(
        name,
        categories,
        series_dict,
        ylabel="kTL",
        use_default_colors=False,
        color_palette=None,
        label_threshold_override=None,
        extra_bar=None,
        n_yoy=2,
        extra_legend=None,
    ):
        path = os.path.join(SCRIPT_DIR, f"temp_stk_{name}.png")
        save_stacked_chart(
            name,
            categories,
            series_dict,
            path,
            ylabel,
            use_default_colors=use_default_colors,
            color_palette=color_palette,
            label_threshold_override=label_threshold_override,
            extra_bar=extra_bar,
            n_yoy=n_yoy,
            extra_legend=extra_legend,
        )
        chart_paths.append(path)
        return path

    def grouped_chart(
        name, categories, series_dict, ylabel="kTL", percents=None, color_palette=None
    ):
        path = os.path.join(SCRIPT_DIR, f"temp_grp_{name}.png")
        save_grouped_chart(
            name,
            categories,
            series_dict,
            path,
            ylabel,
            percents=percents,
            color_palette=color_palette,
        )
        chart_paths.append(path)
        return path

    def bu_ebit_chart(name, categories, values, percents, ylabel="kTL", bu_color="#0EA5E9", bu_label="BU"):
        path = os.path.join(SCRIPT_DIR, f"temp_buebit_{name}.png")
        save_bu_ebit_chart(name, categories, values, percents, path, ylabel, bu_color, bu_label)
        chart_paths.append(path)
        return path

    # Slide 2 — NS kTL  (add ABNS as a standalone red bar at the end)
    _nuc_color = _bu_color_map.get("NUC", "#016B61")
    _abns_bar = ("ABNS", oi_mon_abns, _nuc_color) if oi_mon_abns > 0 else None
    p_gr_ktl = stacked_chart(
        "Gross Revenue in kTL",
        cats_ns_ktl,
        {"Contract": s2_contract, "WP": s2_wp, "WO": s2_wo},
        extra_bar=_abns_bar,
        n_yoy=3,  # group 2025 + 2026 Target + BL together before monthly bars
        extra_legend=[("ABNS", _nuc_color)] if _abns_bar else None,
    )

    # Slide 3 — NS kEUR — group 2025 + 2026 Target + BL together before monthly bars
    p_gr_keur = stacked_chart(
        "Gross Revenue in kEUR",
        cats_ns_keur,
        {"Contract": s3_contract, "WP": s3_wp, "WO": s3_wo},
        ylabel="kEUR",
        n_yoy=3,
    )

    # Slide 4 — EBIT kTL
    p_ebit_ktl = grouped_chart(
        "EBIT in kTL",
        cats_ebit_ktl,
        {"Contract": s4_contract, "Contract+WP": s4_cwp, "Contract+WP+WO": s4_cwp_wo},
        percents=ebit_p_ktl,
    )

    # Slide 5 — EBIT kEUR
    p_ebit_keur = grouped_chart(
        "EBIT in kEUR",
        cats_ebit_keur,
        {"Contract": s5_contract, "Contract+WP": s5_cwp, "Contract+WP+WO": s5_cwp_wo},
        ylabel="kEUR",
        percents=ebit_p_keur,
    )

    # Slide 6 — OI kTL (use DEFAULT colors for Business Units)
    p_oi_ktl = stacked_chart(
        "Order Intake in kTL",
        cats_oi,
        {"Engineering": s6_eng, "MC": s6_mc, "T&SI": s6_tsi, "Nuclear": s6_nuc},
        use_default_colors=True,
    )

    # BU charts
    bu_chart_paths = {}
    bu_labels = [
        ("ENG", "Engineering"),
        ("MC", "MC"),
        ("T&SI", "T&SI"),
        ("NUC", "Nuclear"),
    ]

    for bu_idx, (bu, bu_label) in enumerate(bu_labels):
        # Generate 3 refined tints of the BU color
        base_hex = DEFAULT_BU_COLORS[bu_idx]
        import matplotlib.colors as mcolors

        base_rgb = mcolors.to_rgb(base_hex)
        base_hsv = list(mcolors.rgb_to_hsv(base_rgb))

        # Shades: Order (Pure BU Color), Offer (Mid Tint), Opp (Light Tint)
        # Multipliers are loaded from config BU_Shade_Tiers
        _tiers = config.get("BU_Shade_Tiers", {})
        _bri_floor = _tiers.get("Base_Brightness_Floor", 0.62)
        # Lift dark base colors to the floor so tints stay readable
        _base_v = max(base_hsv[2], _bri_floor)

        def _shade(sat_mult, bri_mult):
            return mcolors.to_hex(
                mcolors.hsv_to_rgb(
                    (
                        base_hsv[0],
                        base_hsv[1] * sat_mult,
                        min(1.0, _base_v * bri_mult),
                    )
                )
            )

        def _tier_color(tier_key, default_sat, default_bri):
            """Return color for a tier, using Color override if present."""
            tier = _tiers.get(tier_key, {})
            override = tier.get("Color")
            if override:
                return override
            return _shade(
                tier.get("Saturation_Multiplier", default_sat),
                tier.get("Brightness_Multiplier", default_bri),
            )

        bu_shades = [
            _tier_color("Order", 1.0,  1.0),
            _tier_color("Offer", 0.55, 1.1),
            _tier_color("Opp",   0.25, 1.15),
        ]

        # NS Chart: Use the refined departmental tints
        _bu_thold_map = config.get("APP.Formatting.BU_Label_Threshold", {})
        _bu_thold = _bu_thold_map.get(bu.replace("&", "").replace(" ", ""), None)
        p_ns = stacked_chart(
            f"Net Sales — {bu_label}",
            cats_bu_ns,
            {
                "Order": bu_ns[bu]["Order"],
                "Offer": bu_ns[bu]["Offer"],
                "Opp": bu_ns[bu]["Opp"],
            },
            color_palette=bu_shades,
            label_threshold_override=_bu_thold,
        )

        # EBIT Chart: use save_bu_ebit_chart — matches reference PPTX style
        # (green 2025 bar, dark-blue 2026 Target bar, BU-colour monthly bars,
        #  % labels inside bars, data table at the bottom)
        _bu_pct_rows = {
            "ENG":  R.get("BU_EBIT_P_ENG", 108),
            "MC":   R.get("BU_EBIT_P_MC",  107),
            "T&SI": R.get("BU_EBIT_P_TSI", 109),
            "NUC":  R.get("BU_EBIT_P_NUC", 110),
        }
        bu_pct_row = _bu_pct_rows.get(bu, 100)
        pcts = bu_ebit_row(bu_pct_row)

        p_eb = bu_ebit_chart(
            f"EBIT — {bu_label}",
            cats_bu_ebit,
            bu_ebit[bu],
            pcts,
            bu_color=bu_shades[0],  # Use brightness-floored BU color (matches NS Order shade)
            bu_label=bu_label,
        )
        bu_chart_paths[bu] = (p_ns, p_eb)

    # ── Build PPTX with pre-made template ──────────────────────────────
    theme_info = config.get_active_theme_template()
    theme = theme_info.get("theme", "default")
    template_path = theme_info.get("template")

    if template_path:
        template_path_norm = os.path.normpath(template_path)
        full_template_path = os.path.abspath(
            os.path.join(SCRIPT_DIR, template_path_norm)
        )

        if os.path.isfile(full_template_path):
            try:
                prs = Presentation(full_template_path)
                print(
                    f"[Theme] Active: {theme} | [Template] Loaded: {os.path.basename(full_template_path)}"
                )
                print(
                    f"[Template] {len(prs.slides)} slides ready (title + 15 content + thank you)"
                )
            except Exception as e:
                print(
                    f"WARNING: Unable to open PPTX template '{full_template_path}': {e}"
                )
                print("Trying to copy template locally...")
                try:
                    import shutil

                    temp_template = os.path.join(SCRIPT_DIR, "template_temp.pptx")
                    shutil.copy2(full_template_path, temp_template)
                    prs = Presentation(temp_template)
                    print(f"[Theme] Active: {theme} | [Template] Loaded from temp copy")
                    print(f"[Template] {len(prs.slides)} slides ready")
                    os.remove(temp_template)
                except Exception as e2:
                    print(f"Fallback: creating blank presentation: {e2}")
                    prs = Presentation()
        else:
            print(f"WARNING: template file not found: {full_template_path}")
            prs = Presentation()
            print(f"[Theme] {theme} | Using blank presentation")
    else:
        prs = Presentation()
        print(f"[Theme] {theme} | No template configured - using blank")

    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # ── Slide 0: Title ────────────────────────────────────────────────
    s0 = prs.slides[0]
    for shape in s0.shapes:
        if hasattr(shape, "text_frame"):
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if "[YEAR]" in run.text:
                        run.text = run.text.replace("[YEAR]", year)
                    if "[MONTH]" in run.text:
                        run.text = run.text.replace("[MONTH]", month)

    # ── Slide last: Thank You — replace [MONTH] / [YEAR] placeholders ─
    s_ty = prs.slides[-1]
    for shape in s_ty.shapes:
        if hasattr(shape, "text_frame"):
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if "[YEAR]" in run.text:
                        run.text = run.text.replace("[YEAR]", year)
                    if "[MONTH]" in run.text:
                        run.text = run.text.replace("[MONTH]", month)

    # ── Update [SUBTITLE] placeholder on all content slides (1-15) ────
    subtitle_text = f"{month} {year}  ·  Source: Project Budget Analysis"
    for si in range(1, 16):
        update_slide_subtitle(prs.slides[si], subtitle_text)

    # ── Update OI table slide title to include month ────────────────
    update_slide_title(prs.slides[6], "Order Intake", f"Order Intake ({month} {year})")

    # ── Slides 1-5: Financial charts ──────────────────────────────────
    chart_map = [
        (1, p_gr_ktl),
        (2, p_gr_keur),
        (3, p_ebit_ktl),
        (4, p_ebit_keur),
        (5, p_oi_ktl),
    ]

    for slide_idx, chart_path in chart_map:
        if chart_path and os.path.exists(chart_path):
            add_chart_image(
                prs.slides[slide_idx], chart_path, top=Inches(1.25), width=Inches(12.0)
            )

            # --- Add Key Metrics Callouts ---
            if slide_idx == 1:  # GR kTL (Net Sales)
                # Build ABNS project list for green callout (ALL projects)
                abns_lines = []
                for i, (proj, cl, val) in enumerate(oi_abns_all):
                    proj_short = proj[:28] + "..." if len(proj) > 28 else proj
                    abns_lines.append(
                        (f"#{i + 1}", f"{proj_short}: {human_k(val)}")
                    )
                if not abns_lines:
                    abns_lines = [("ABNS", "No ABNS projects this month")]

                callout_lines = [
                    (f"{month} NS", human_k(gr_mon_ktl)),
                    (f"FY {year} NS", human_k(gr_dec_ktl)),
                ]
                # ABNS total in red
                if oi_mon_abns > 0:
                    callout_lines.append(("FY ABNS", human_k(oi_mon_abns), _nuc_color))

                # Main callout box (right side)
                add_callout_box(
                    prs.slides[1],
                    callout_lines,
                )

                # Second callout (left side) — ALL ABNS projects, green border
                if oi_abns_all:
                    add_second_callout_box(
                        prs.slides[1],
                        abns_lines,
                        left=Inches(0.45),
                        top=Inches(6.15),
                        width=Inches(4.0),
                    )
            elif slide_idx == 2:  # GR kEUR
                add_callout_box(
                    prs.slides[2],
                    [
                        (f"{month} NS", human_k(gr_mon_keur)),
                        (f"FY {year} NS", human_k(gr_dec_keur)),
                    ],
                    left=Inches(10.5),
                )
            elif slide_idx == 3:  # EBIT kTL
                add_callout_box(
                    prs.slides[3],
                    [
                        (f"{month} EBIT", human_k(ebit_mon_ktl)),
                        (f"FY {year} EBIT", human_k(ebit_dec_ktl)),
                    ],
                )
            elif slide_idx == 4:  # EBIT kEUR
                add_callout_box(
                    prs.slides[4],
                    [
                        (f"{month} EBIT", human_k(ebit_mon_keur)),
                        (f"FY {year} EBIT", human_k(ebit_dec_keur)),
                    ],
                )
            elif slide_idx == 5:  # Order Intake
                callout_lines = [
                    (f"{month} Signed", human_k(oi_callout_mon)),
                    (f"FY {year} OI", human_k(oi_total)),
                ]
                if oi_mon_abns > 0:
                    callout_lines.append(("FY ABNS", human_k(oi_mon_abns)))

                add_callout_box(
                    prs.slides[5],
                    callout_lines,
                )

    # ── Slide 6: Order Intake project details table ───────────────────
    add_oi_table_to_slide(prs.slides[6], oi_projects, chart_total=oi_mon_total)

    # ── WIP Table(s): Split into multiple slides if needed ───────────
    # Strategy:
    #   1. Determine how many slides are needed (chunk count) FIRST.
    #   2. Pre-create all overflow slides by copying the CLEAN WIP template
    #      slide (index 7, before any content is written to it).  Using the
    #      WIP template slide — rather than a later populated chart slide —
    #      guarantees we copy only sidebar/branding, with no ghost titles or
    #      broken picture relationships.
    #   3. Reorder the new slides into their final positions immediately.
    #   4. Populate ALL WIP slides with content (headers + tables).
    #   5. Attach the Excel file to the LAST WIP slide only.
    MAX_ROWS_PER_SLIDE = config.get("WIP_Table.WIP_Slide_MaxRows", 10)
    from pptx.oxml.ns import qn as _qn

    extra_wip_slides = 0

    if wip_slide_rows:
        # ── Step 1: Calculate chunks ──────────────────────────────────
        wip_chunks = [
            wip_slide_rows[i : i + MAX_ROWS_PER_SLIDE]
            for i in range(0, len(wip_slide_rows), MAX_ROWS_PER_SLIDE)
        ]
        extra_wip_slides = len(wip_chunks) - 1
        print(
            f"  [WIP] {len(wip_slide_rows)} rows -> "
            f"{len(wip_chunks)} slide(s) ({MAX_ROWS_PER_SLIDE} rows/slide max)"
        )

        # ── Step 2 & 3: Pre-create overflow slides from clean WIP template ──
        # Copy from prs.slides[7] NOW, before it gets populated with content.
        # source_slide_idx=7 ensures we get the clean template sidebar branding.
        overflow_slides = []
        for chunk_idx in range(1, len(wip_chunks)):
            new_slide = add_blank_slide(prs, source_slide_idx=7)

            # Reorder: move the newly appended slide to position 7 + chunk_idx
            _sldIdLst = prs.part._element.find(_qn("p:sldIdLst"))
            _all_ids = list(_sldIdLst)
            _new_id = _all_ids[-1]
            target_pos = 7 + chunk_idx
            if target_pos < len(_all_ids) - 1:
                _sldIdLst.remove(_new_id)
                _sldIdLst.insert(target_pos, _new_id)

            overflow_slides.append(new_slide)
            print(f"  [WIP] Created overflow slide for page {chunk_idx + 1}")

        # ── Step 4: Populate all WIP slides with headers + tables ────
        # First slide (slide 7) — keep its existing template header; just add table
        add_wip_table_to_slide(prs.slides[7], wip_chunks[0], slide_num=7)

        for chunk_idx, (new_slide, chunk) in enumerate(
            zip(overflow_slides, wip_chunks[1:]), start=1
        ):
            add_header(
                new_slide,
                f"WIP (Work In Progress) - continued (Page {chunk_idx + 1})",
                f"{month} {year}",
            )
            add_wip_table_to_slide(new_slide, chunk, slide_num=7 + chunk_idx)
    else:
        add_wip_table_to_slide(prs.slides[7], [], slide_num=7)

    # ── Export full WIP detail to Excel ────────────────────────────────────────────────
    wip_ole_params = None

    if wip_excel_rows:
        wip_xlsx_path = export_wip_to_excel(
            wip_excel_rows, month, year, SCRIPT_DIR,
            slide_rows=wip_slide_rows, negative_rows=wip_negative_rows
        )

        _ole_y = Inches(6.92)
        wip_ole_params = (
            wip_xlsx_path,
            None,
            int(Inches(0.45)),
            int(_ole_y),
            int(Inches(2.1)),
            int(Inches(0.45)),
        )

    # ── Negative WIP slide (dedicated slide after all positive WIP slides) ────────
    _last_pos_wip_idx = 7 + extra_wip_slides  # last positive WIP slide index (before negative increment)
    if wip_negative_rows:
        _neg_slide = add_blank_slide(prs, source_slide_idx=7)

        # Reorder: insert immediately after the last positive WIP slide
        _sldIdLst2 = prs.part._element.find(_qn("p:sldIdLst"))
        _all_ids2 = list(_sldIdLst2)
        _new_neg_id = _all_ids2[-1]
        _neg_target_pos = 7 + extra_wip_slides + 1
        if _neg_target_pos < len(_all_ids2) - 1:
            _sldIdLst2.remove(_new_neg_id)
            _sldIdLst2.insert(_neg_target_pos, _new_neg_id)

        extra_wip_slides += 1  # negative slide counts as an extra WIP slide

        add_header(
            _neg_slide,
            "WIP (Work In Progress) — Negative Projects",
            f"{month} {year}",
        )
        _slide_total_wip = sum(r["wip_tl"] for r in wip_slide_rows)
        add_wip_negative_table_to_slide(
            _neg_slide,
            wip_negative_rows,
            _slide_total_wip,
            slide_num=7 + extra_wip_slides,
        )
        print(f"  [WIP] Created Negative Projects slide at position {7 + extra_wip_slides}")

    # ── Note on LAST POSITIVE WIP slide (not the negative slide) ────────────────
    last_wip_slide_num = 7 + extra_wip_slides

    if wip_excel_rows:
        n_slide_rows = len(wip_slide_rows)
        n_excel_remaining = len(wip_excel_rows) - n_slide_rows

        if _last_pos_wip_idx < len(prs.slides):
            tb_note = prs.slides[_last_pos_wip_idx].shapes.add_textbox(
                Inches(0.45), Inches(6.55), Inches(12.43), Inches(0.35)
            )
            tf = tb_note.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            run = p.add_run()
            if n_excel_remaining > 0:
                run.text = (
                    f"\u25b6  {n_excel_remaining} additional project(s) < 1M TL not shown. "
                    f"See attached Excel for full WIP list."
                )
            else:
                run.text = "\u25b6  See attached Excel for complete WIP list."
            run.font.size = Pt(8.5)
            run.font.italic = True
            run.font.color.rgb = GRAY_MID

    # ── Slides 8-15: Business Units (adjusted for extra WIP slides) ───────────────────────────────────
    # Adjust slide indices to account for extra WIP slides
    bu_slide_start = (
        8 + extra_wip_slides
    )  # First BU slide shifts if there are extra WIP slides
    bu_chart_configs = [
        (bu_slide_start, bu_chart_paths.get("ENG", (None, None))[0]),
        (bu_slide_start + 1, bu_chart_paths.get("ENG", (None, None))[1]),
        (bu_slide_start + 2, bu_chart_paths.get("MC", (None, None))[0]),
        (bu_slide_start + 3, bu_chart_paths.get("MC", (None, None))[1]),
        (bu_slide_start + 4, bu_chart_paths.get("T&SI", (None, None))[0]),
        (bu_slide_start + 5, bu_chart_paths.get("T&SI", (None, None))[1]),
        (bu_slide_start + 6, bu_chart_paths.get("NUC", (None, None))[0]),
        (bu_slide_start + 7, bu_chart_paths.get("NUC", (None, None))[1]),
    ]

    for bu_idx, (slide_idx, chart_path) in enumerate(bu_chart_configs):
        if chart_path and os.path.exists(chart_path) and slide_idx < len(prs.slides):
            s = prs.slides[slide_idx]
            add_chart_image(s, chart_path, top=Inches(1.25), width=Inches(12.0))

            # --- BU Monthly Highlights ---
            bu_key = bu_labels[bu_idx // 2][0]
            if bu_idx % 2 == 0:  # NS Slide (0, 2, 4, 6)
                mon_val = bu_ns_singular.get(bu_key, 0)
                add_callout_box(
                    s, [(f"{month} NS", human_k(mon_val))]
                )
            # EBIT slides (1, 3, 5, 7): no callout box

    # ── Save PPTX ──────────────────────────────────────────────────────
    output_ppt = os.path.join(SCRIPT_DIR, f"MRC_MBR_Stacked_{month}_{year}.pptx")
    prs.save(output_ppt)

    if wip_ole_params:
        # Fix 2: resolve slide filename using the correct XML order (sldIdLst)
        # Alphabetical sorting was causing us to patch the wrong slide.
        import zipfile as _zipfile
        from lxml import etree as _et2

        _slide_fn = None
        _target_slide_idx = _last_pos_wip_idx  # last positive WIP slide (OLE goes on positive slide, not negative)
        try:
            with _zipfile.ZipFile(output_ppt, "r") as _z:
                # 1. Map relationship IDs to targets from presentation.xml.rels
                _prels_xml = _z.read("ppt/_rels/presentation.xml.rels")
                _prels_root = _et2.fromstring(_prels_xml)
                _rid_to_target = {
                    rel.get("Id"): rel.get("Target") for rel in _prels_root
                }

                # 2. Find the LAST WIP slide in the presentation.xml sequence
                # After reordering, it's at index 7 + extra_wip_slides
                _pxml = _z.read("ppt/presentation.xml")
                _proot = _et2.fromstring(_pxml)
                _ns_m = {
                    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                }
                _sld_ids = _proot.findall(".//p:sldId", _ns_m)

                if len(_sld_ids) > _target_slide_idx:
                    _target_rid = _sld_ids[_target_slide_idx].get(f"{{{_ns_m['r']}}}id")
                    _target_path = _rid_to_target.get(_target_rid)
                    if _target_path:
                        _slide_fn = os.path.basename(_target_path)
                        print(
                            f"  [OLE] Resolved Slide {_target_slide_idx + 1} target: {_slide_fn} (via rId {_target_rid})"
                        )
        except Exception as _e:
            print(f"  [OLE] Error resolving slide filename: {_e}")

        _xlsx_path = wip_ole_params[0]
        if _slide_fn and os.path.isfile(_xlsx_path):
            embed_ole_excel_to_slide(
                output_ppt, _xlsx_path, _slide_fn, *wip_ole_params[2:]
            )
        else:
            print(
                f"  [OLE] Skipped embedding — slide_fn={_slide_fn!r}, xlsx exists={os.path.isfile(_xlsx_path)}"
            )

    print(f"\n[DONE] Saved:  {output_ppt}  ({len(prs.slides)} slides)")
    for p in chart_paths:
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass


# ──────────────────────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────────────────────
_MONTHS_FULL = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def _find_best_excel(month_filter=None):
    """
    Scan SCRIPT_DIR for Budget Analysis Excel files.
    For each month found, returns the file with the highest version number.
    If month_filter is given (e.g. "January"), returns only that month's best file.
    Returns the chosen file path or None.
    """
    candidates = [
        f for f in os.listdir(SCRIPT_DIR)
        if f.endswith(".xlsx") and not f.startswith("~")
        and ("Budget_Analysis" in f or "budget_analysis" in f.lower())
    ]

    # Parse (month, major, minor) from each candidate
    def _parse(fname):
        month = None
        for m in _MONTHS_FULL:
            if m.lower() in fname.lower():
                month = m
                break
        ver = re.search(r'[vV](\d+)[._]?(\d*)', fname)
        major = int(ver.group(1)) if ver else 0
        minor = int(ver.group(2)) if ver and ver.group(2) else 0
        return month, (major, minor)

    # Group by month → keep highest version
    best = {}  # month → (filename, version_tuple)
    no_month = []
    for f in candidates:
        month, ver = _parse(f)
        if month:
            if month not in best or ver > best[month][1]:
                best[month] = (f, ver)
        else:
            no_month.append(f)

    if month_filter:
        m_cap = month_filter.strip().capitalize()
        # Try full name match first
        if m_cap in best:
            chosen = os.path.join(SCRIPT_DIR, best[m_cap][0])
            print(f"[Auto-detect] Month='{m_cap}' -> {best[m_cap][0]} (v{best[m_cap][1][0]}.{best[m_cap][1][1]})")
            return chosen
        # Try abbreviation (e.g. "jan" → "January")
        for m in _MONTHS_FULL:
            if m.lower().startswith(m_cap.lower()):
                if m in best:
                    chosen = os.path.join(SCRIPT_DIR, best[m][0])
                    print(f"[Auto-detect] Month='{m}' -> {best[m][0]} (v{best[m][1][0]}.{best[m][1][1]})")
                    return chosen
        print(f"[Auto-detect] No Budget Analysis file found for month '{month_filter}'")
        print(f"  Available months: {sorted(best.keys(), key=lambda x: _MONTHS_FULL.index(x))}")
        return None

    # No month filter: pick the entry with the latest (date_prefix, version)
    # Use the YYMMDD prefix from the filename so Dec-2025 ranks below Jan-2026, etc.
    if best:
        def _sort_key(item):
            _, (fname, (maj, mn)) = item
            date_match = re.match(r'^(\d{6})', fname)
            date_prefix = int(date_match.group(1)) if date_match else 0
            return (date_prefix, maj, mn)
        latest_month, (latest_file, latest_ver) = max(best.items(), key=_sort_key)
        chosen = os.path.join(SCRIPT_DIR, latest_file)
        print(f"[Auto-detect] Latest: {latest_file}  (month={latest_month}, v{latest_ver[0]}.{latest_ver[1]})")
        if len(best) > 1:
            print(f"  Other available months: {[m for m in sorted(best.keys(), key=lambda x: _MONTHS_FULL.index(x)) if m != latest_month]}")
        return chosen

    # Fall back to most-recently-modified xlsx if no versioned files found
    if no_month:
        no_month.sort(key=lambda x: os.path.getmtime(os.path.join(SCRIPT_DIR, x)), reverse=True)
        chosen = os.path.join(SCRIPT_DIR, no_month[0])
        print(f"[Auto-detect] No versioned file found; using most-recently-modified: {no_month[0]}")
        return chosen

    return None


if __name__ == "__main__":
    import sys

    # Usage:
    #   python generate_full_mbr_stacked.py                   → auto-detect latest month
    #   python generate_full_mbr_stacked.py January           → pick highest-version January file
    #   python generate_full_mbr_stacked.py path/to/file.xlsx → use explicit file

    target = None
    month_arg = None

    if len(sys.argv) > 1:
        arg = sys.argv[1].strip()
        # Check if arg is a month name or abbreviation
        arg_cap = arg.capitalize()
        matched_month = None
        for m in _MONTHS_FULL:
            if m.lower() == arg_cap.lower() or m.lower().startswith(arg_cap.lower()):
                matched_month = m
                break
        if matched_month:
            month_arg = matched_month
        elif os.path.exists(os.path.abspath(arg)):
            target = os.path.abspath(arg)
        else:
            print(f"[ERROR] '{arg}' is not a valid file path or month name.")
            print(f"  Valid months: {', '.join(_MONTHS_FULL)}")
            sys.exit(1)

    if target is None:
        target = _find_best_excel(month_filter=month_arg)

    if target and os.path.exists(target):
        create_17_slide_mbr_stacked(target)
    else:
        print("[ERROR] No suitable Excel file found. Please pass a file path or month name.")
        sys.exit(1)
