"""
dashboard_shared.py
====================
Pure, side-effect-free constants and helper functions shared across every
dashboard page (Metals, Energy, Precious Metals, NGL, Home). No top-level
Streamlit rendering calls live here -- only things safe to import from a
page module without re-executing another page as a side effect.

Conventions replicated project-wide (see CLAUDE.md / dashboard_metals.py):
  - PnL = position x delta(F1_continuous), ratio back-adjusted.
  - Transaction costs charged on F1_raw (the actual traded price), not F1_continuous.
  - Sharpe = annualised, active-day convention (days with a non-zero position only).
  - Same-Day execution = shift-1 (position[t] = signal[t-1]).
    Lag-1 execution   = shift-2 (position[t] = signal[t-2]).
"""

import numpy as np
import pandas as pd
import streamlit as st

# ═══════════════════════════════════════════════
# GLOBAL CSS THEME (inject once from app.py)
# ═══════════════════════════════════════════════

DASHBOARD_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

    /* Global */
    .stApp { font-family: 'IBM Plex Sans', sans-serif; background-color: #0E0E0E; }

    /* Hide default streamlit elements */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    /* Metric cards */
    .metric-card {
        background: #161616;
        border: 1px solid #2A2A2A;
        border-left: 3px solid #B87333;
        border-radius: 4px;
        padding: 14px 18px;
        margin: 4px 0;
    }
    .metric-card h4 {
        color: #7A7068;
        font-size: 0.72rem;
        font-weight: 500;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin: 0 0 6px 0;
    }
    .metric-card .value {
        color: #D4CFC8;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.45rem;
        font-weight: 500;
        margin: 0;
    }
    .metric-card .delta-pos { color: #5BAD72; font-size: 0.82rem; }
    .metric-card .delta-neg { color: #B85450; font-size: 0.82rem; }

    /* Compact metric cards (fits many per row) */
    .metric-compact {
        background: #161616;
        border: 1px solid #2A2A2A;
        border-left: 3px solid #B87333;
        border-radius: 4px;
        padding: 7px 10px;
        margin: 3px 0;
    }
    .metric-compact h4 {
        color: #7A7068;
        font-size: 0.62rem;
        font-weight: 500;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        margin: 0 0 3px 0;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .metric-compact .value {
        color: #D4CFC8;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.1rem;
        font-weight: 500;
        margin: 0;
        white-space: nowrap;
    }

    /* Section headers */
    .section-header {
        font-family: 'IBM Plex Sans', sans-serif;
        color: #D4CFC8;
        font-size: 1rem;
        font-weight: 600;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        border-bottom: 1px solid #B87333;
        padding-bottom: 6px;
        margin: 24px 0 14px 0;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        background-color: #111111;
        padding: 0;
        border-bottom: 1px solid #2A2A2A;
        border-radius: 0;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 0;
        padding: 10px 22px;
        font-weight: 500;
        font-size: 0.88rem;
        letter-spacing: 0.03em;
        color: #7A7068;
        border-bottom: 2px solid transparent;
    }
    .stTabs [aria-selected="true"] {
        color: #B87333 !important;
        border-bottom: 2px solid #B87333 !important;
        background-color: transparent !important;
    }

    /* Backwardation / Contango badges */
    .badge-backwardation {
        background: rgba(91, 173, 114, 0.12);
        color: #5BAD72;
        padding: 3px 10px;
        border-radius: 2px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.04em;
        text-transform: uppercase;
    }
    .badge-contango {
        background: rgba(184, 84, 80, 0.12);
        color: #B85450;
        padding: 3px 10px;
        border-radius: 2px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.04em;
        text-transform: uppercase;
    }

    /* Title */
    .main-title {
        font-family: 'IBM Plex Sans', sans-serif;
        font-weight: 600;
        font-size: 1.6rem;
        color: #D4CFC8;
        margin-bottom: 0;
        letter-spacing: 0.02em;
    }
    .main-subtitle {
        color: #5A5248;
        font-size: 0.85rem;
        margin-top: 2px;
        letter-spacing: 0.04em;
    }
</style>
"""


def inject_css():
    st.markdown(DASHBOARD_CSS, unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# CHART THEME
# ═══════════════════════════════════════════════

CHART_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="#111111",
    font=dict(family="IBM Plex Sans, sans-serif", color="#8A8278"),
    xaxis=dict(gridcolor="rgba(50,46,42,0.6)", zerolinecolor="rgba(50,46,42,0.6)"),
    yaxis=dict(gridcolor="rgba(50,46,42,0.6)", zerolinecolor="rgba(50,46,42,0.6)"),
    margin=dict(l=60, r=30, t=50, b=50),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11, color="#8A8278")),
    hoverlabel=dict(bgcolor="#1C1A18", font_size=12, font_family="IBM Plex Mono"),
)

COLORS = {
    "primary": "#B87333",    # copper
    "secondary": "#C9A84C",  # gold
    "accent": "#3D8F8A",     # muted teal
    "green": "#5BAD72",      # muted green
    "red": "#B85450",        # muted red
    "amber": "#C9A84C",      # amber/gold
    "orange": "#B87333",     # copper-orange
    "pink": "#A07898",       # muted mauve
    "slate": "#6A6460",      # warm gray
}

METAL_COLORS = {
    "Copper":    "#B87333",
    "Aluminium": "#9BAAB3",
    "Zinc":      "#7A8E9A",
    "Nickel":    "#A0A5A8",
    "Lead":      "#6B7073",
    "Tin":       "#9A9EA0",
    "Gold":      "#C9A84C",
    "Silver":    "#B0B8C0",
    "Platinum":  "#C8D0D8",
    "Palladium": "#B8A898",
}


# ═══════════════════════════════════════════════
# GENERIC UI HELPERS
# ═══════════════════════════════════════════════

def metric_card(label, value, delta=None, unit=""):
    delta_html = ""
    if delta is not None:
        cls = "delta-pos" if delta >= 0 else "delta-neg"
        sign = "+" if delta >= 0 else ""
        delta_html = f'<span class="{cls}">{sign}{delta:.2f}%</span>'

    st.markdown(f"""
    <div class="metric-card">
        <h4>{label}</h4>
        <p class="value">{value}{unit}</p>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


def section_header(text):
    st.markdown(f'<div class="section-header">{text}</div>', unsafe_allow_html=True)


def _fmt_sh(x):  return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:+.2f}"
def _fmt_pct(x): return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:+.1f}%"
def _fmt_dd(x):  return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.0f}%"


def filter_date(df, start, end):
    if df.empty:
        return df
    try:
        if not isinstance(df.index, pd.DatetimeIndex):
            df = df.copy()
            df.index = pd.to_datetime(df.index, errors="coerce")
            df = df[df.index.notna()]
        if df.empty:
            return df
        mask = (df.index >= pd.Timestamp(start)) & (df.index <= pd.Timestamp(end))
        return df[mask]
    except Exception:
        return df


def find_curve_sheet(product_name, curve_data):
    """Return the best-matching sheet name in curve_data for a given product name."""
    clean = (product_name.lower()
             .replace("($/oz)", "").replace("($/lb)", "")
             .replace("comex", "").strip())
    for sheet_name in curve_data:
        sl = sheet_name.lower()
        if clean in sl or sl in clean:
            return sheet_name
    first = clean.split()[0] if clean.split() else ""
    for sheet_name in curve_data:
        if first and first in sheet_name.lower():
            return sheet_name
    return None


# ═══════════════════════════════════════════════
# GENERIC SIGNAL EXECUTION + PERFORMANCE METRICS
# (identical convention to the Metals dashboard: PnL on F1_continuous,
#  TC on F1_raw, active-day Sharpe, no look-ahead)
# ═══════════════════════════════════════════════

def exec_shift(sigbin, same_day):
    """Same-Day = shift 1, Lag-1 = shift 2 (no look-ahead)."""
    return sigbin.shift(1) if same_day else sigbin.shift(2)


def pos_metrics_generic(pos, f1r, f1c, tc_bps: int = 5) -> dict:
    """Active-day gross/net Sharpe, ann return %, max-DD % for a position series.
    PnL on F1_continuous; TC on F1_raw. Identical convention to the Metals tabs."""
    pos = pos.reindex(f1c.index).fillna(0.0)
    gp = pos * f1c.diff()
    chg = pos.diff().abs()
    if len(chg):
        chg.iloc[0] = abs(pos.iloc[0])
    tc = chg * (tc_bps / 10000.0 / 2.0) * f1r.reindex(f1c.index)
    net = gp - tc

    def _s(pnl):
        with np.errstate(invalid="ignore", divide="ignore"):
            r = (pnl / f1c.shift(1)).replace([np.inf, -np.inf], np.nan)
        a = r[pos != 0].dropna()
        return float(a.mean() / a.std(ddof=1) * np.sqrt(252)) if len(a) > 20 and a.std(ddof=1) > 0 else np.nan

    with np.errstate(invalid="ignore", divide="ignore"):
        gr = (gp / f1c.shift(1)).replace([np.inf, -np.inf], np.nan)
    cum = gr.fillna(0).cumsum() * 100
    return dict(gross=_s(gp), net=_s(net),
                ann=float(gr.dropna().mean() * 252 * 100) if gr.notna().any() else np.nan,
                mdd=float((cum - cum.cummax()).min()), nact=int((pos != 0).sum()),
                flat_pct=float(100 * (pos == 0).sum() / len(pos)) if len(pos) else np.nan)


def tc_label_map(last_price: float, unit_label: str = "/unit") -> dict:
    """TC selectbox options with bps and native-unit per-flip equivalent in each label.
    unit_label example: '/MT', '/bbl', '/MMBtu', '/oz', '/gal' -- pass the product's
    natural pricing unit so the $ estimate reads correctly for that asset class."""
    def _lbl(bps):
        flip_cost = (bps / 10000.0) * last_price
        return f"{bps} bps  (~${flip_cost:.2f}{unit_label} per flip)"
    return {
        "0 bps  (Gross)": 0,
        _lbl(5):  5,
        _lbl(10): 10,
        _lbl(20): 20,
    }
