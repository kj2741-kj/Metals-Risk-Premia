"""
Metals Risk Premia — Interactive Dashboard
============================================
Streamlit dashboard for exploring LME & CME metals data.

Local files (auto-loaded if present in same directory):
  1. Metals Cash and 3M.xlsx
  2. Metals Futures Curve.csv
"""

import io
import os
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots
from scipy import stats

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════

st.set_page_config(
    page_title="Metals Risk Premia Dashboard",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════
# STYLING
# ═══════════════════════════════════════════════

st.markdown("""
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

    /* Compact metric cards (momentum tab — fits 8 per row) */
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
""", unsafe_allow_html=True)


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
    "Copper":    "#B87333",  # real copper
    "Aluminium": "#9BAAB3",  # aluminum
    "Zinc":      "#7A8E9A",  # zinc blue-gray
    "Nickel":    "#A0A5A8",  # nickel silver
    "Lead":      "#6B7073",  # lead dark gray
    "Tin":       "#9A9EA0",  # tin
    "Gold":      "#C9A84C",  # gold
    "Silver":    "#B0B8C0",  # silver
    "Platinum":  "#C8D0D8",  # platinum
    "Palladium": "#B8A898",  # palladium warm gray
}


# ═══════════════════════════════════════════════
# LOCAL FILE HELPERS
# ═══════════════════════════════════════════════

LOCAL_CASH_PATH = os.path.join(os.path.dirname(__file__), "Metals Cash and 3M.xlsx")
LOCAL_CURVE_PATH = os.path.join(os.path.dirname(__file__), "Metals Futures Curve.csv")


def _local_bytesio(path):
    """Read a local file into a BytesIO buffer with a .name attribute."""
    with open(path, "rb") as f:
        buf = io.BytesIO(f.read())
    buf.name = os.path.basename(path)
    return buf


# ═══════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════

@st.cache_data(ttl=3600)
def load_cash_3m_data(file):
    """Load Metals Cash and 3M.xlsx — one sheet per metal."""
    xls = pd.ExcelFile(file)
    data = {}

    # LME sheets have a 3-row header:
    #   Row 0 = metal name (merged), Row 1 = instrument names, Row 2 = Price/Volume/Open Int
    lme_sheets = [s for s in xls.sheet_names if "LME" in s]
    for sheet in lme_sheets:
        df = pd.read_excel(xls, sheet_name=sheet, header=[0, 1, 2])
        # Flatten 3-level columns, dropping any "Unnamed" parts
        new_cols = []
        for c in df.columns:
            parts = [str(p).strip() for p in c
                     if str(p).strip() and "Unnamed" not in str(p)]
            new_cols.append("_".join(parts) if parts else str(c))
        df.columns = new_cols

        date_col = [c for c in df.columns if "date" in c.lower()]
        if date_col:
            df = df.rename(columns={date_col[0]: "Date"})
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date"])
            df = df.set_index("Date").sort_index()

        metal_name = sheet.replace("LME ", "").strip()
        data[metal_name] = df

    # CME Cash Prices: blank row 0, actual headers at row 1, blank row 2
    if "CME Cash Prices" in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name="CME Cash Prices", header=1)
        df = df.dropna(how="all")
        date_col = [c for c in df.columns if "date" in str(c).lower()]
        if date_col:
            df = df.rename(columns={date_col[0]: "Date"})
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date"])
            df = df.set_index("Date").sort_index()
        data["CME_Cash"] = df

    return data


@st.cache_data(ttl=3600)
def load_futures_curve_data(file):
    """
    Load Metals Futures Curve — one sheet per metal with F1-F27.
    Handles: .xlsx, .xls, .csv (with encoding fallbacks), and
    xlsx files incorrectly saved with .csv extension.
    """
    fname = file.name if hasattr(file, "name") else str(file)
    data = {}

    # Try reading as Excel first (even if extension is .csv)
    try:
        xls = pd.ExcelFile(file)
        return _parse_curve_excel(xls)
    except Exception:
        pass

    # If Excel fails, try CSV with multiple encodings
    if fname.lower().endswith(".csv"):
        for encoding in ["utf-8", "latin-1", "cp1252", "iso-8859-1", "utf-16"]:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=encoding)
                if not df.empty:
                    data["Sheet1"] = _parse_single_curve_df(df)
                    return data
            except Exception:
                continue

        # Last resort: read as bytes, decode, then parse
        try:
            file.seek(0)
            raw = file.read()
            # Check if it's actually xlsx bytes
            if raw[:4] == b"PK\x03\x04":
                file.seek(0)
                xls = pd.ExcelFile(io.BytesIO(raw))
                return _parse_curve_excel(xls)
            # Otherwise try as text
            for enc in ["utf-8", "latin-1", "cp1252"]:
                try:
                    text = raw.decode(enc)
                    df = pd.read_csv(io.StringIO(text))
                    if not df.empty:
                        data["Sheet1"] = _parse_single_curve_df(df)
                        return data
                except Exception:
                    continue
        except Exception:
            pass

    st.error(f"Could not read '{fname}'. Try saving it as .xlsx from Excel and re-uploading.")
    return data


def _parse_curve_excel(xls):
    """Parse an Excel file with one sheet per metal, multi-row headers."""
    data = {}

    for sheet in xls.sheet_names:
        try:
            # First pass: read raw to detect header structure
            df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=5)

            # Detect header rows by looking for "Date", "F1", "Price" etc.
            header_rows = []
            for i in range(min(4, len(df_raw))):
                row_vals = [str(v).strip().lower() for v in df_raw.iloc[i].values if pd.notna(v)]
                if any(kw in " ".join(row_vals) for kw in ["date", "f1", "f2", "price", "volume"]):
                    header_rows.append(i)

            if len(header_rows) >= 2:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_rows)
            elif len(header_rows) == 1:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_rows[0])
            else:
                df = pd.read_excel(xls, sheet_name=sheet, header=[0, 1, 2])

        except Exception:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=[0, 1])
            except Exception:
                df = pd.read_excel(xls, sheet_name=sheet)

        data[sheet] = _parse_single_curve_df(df)

    return data


def _parse_single_curve_df(df):
    """Parse a single dataframe with futures curve data into standardized format."""
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for col_tuple in df.columns:
            parts = [str(p).strip() for p in col_tuple
                     if pd.notna(p) and "Unnamed" not in str(p) and str(p).strip()]
            new_cols.append("_".join(parts) if parts else str(col_tuple))
        df.columns = new_cols

    df.columns = [str(c).strip() for c in df.columns]

    date_col = [c for c in df.columns if "date" in c.lower()]
    if date_col:
        df = df.rename(columns={date_col[0]: "Date"})
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df = df.dropna(subset=["Date"])
        df = df.set_index("Date").sort_index()

    prices = {}
    for col in df.columns:
        col_lower = col.lower().replace(" ", "_")
        for i in range(1, 28):
            patterns = [
                f"f{i}_price", f"f{i}_Price",
                f"F{i}_Price", f"F{i}_price",
            ]
            if any(p.lower() in col_lower for p in patterns):
                prices[f"F{i}"] = pd.to_numeric(df[col], errors="coerce")
                break
            elif col_lower.startswith(f"f{i}_") and "price" in col_lower:
                prices[f"F{i}"] = pd.to_numeric(df[col], errors="coerce")
                break

    result = {
        "raw": df,
        "prices": pd.DataFrame(prices, index=df.index) if prices else pd.DataFrame()
    }
    return result


def parse_cash_3m_columns(df, metal_name):
    """Parse the Cash & 3M dataframe columns into standardized names."""
    result = pd.DataFrame(index=df.index)

    for col in df.columns:
        cl = col.lower()
        # First-match-wins: don't overwrite a key once populated.
        # Check spread first — spread col names also contain "cash"/"3m"/"price".
        if "spread" in cl and "price" in cl and "spread_price" not in result.columns:
            result["spread_price"] = pd.to_numeric(df[col], errors="coerce")
        elif "spread" in cl and "volume" in cl and "spread_volume" not in result.columns:
            result["spread_volume"] = pd.to_numeric(df[col], errors="coerce")
        elif "spread" not in cl and (("cash" in cl and "price" in cl) or ("spot" in cl and "price" in cl)) and "cash_price" not in result.columns:
            result["cash_price"] = pd.to_numeric(df[col], errors="coerce")
        elif "spread" not in cl and (("3m" in cl and "price" in cl) or ("forward" in cl and "price" in cl)) and "3m_price" not in result.columns:
            result["3m_price"] = pd.to_numeric(df[col], errors="coerce")
        elif "spread" not in cl and "3m" in cl and "volume" in cl and "3m_volume" not in result.columns:
            result["3m_volume"] = pd.to_numeric(df[col], errors="coerce")
        elif "spread" not in cl and "3m" in cl and ("open" in cl or "oi" in cl or "int" in cl) and "3m_oi" not in result.columns:
            result["3m_oi"] = pd.to_numeric(df[col], errors="coerce")

    if "spread_price" not in result.columns and "cash_price" in result.columns and "3m_price" in result.columns:
        result["spread_price"] = result["cash_price"] - result["3m_price"]

    if "cash_price" in result.columns:
        result["cash_return"] = np.log(result["cash_price"] / result["cash_price"].shift(1))
    if "3m_price" in result.columns:
        result["3m_return"] = np.log(result["3m_price"] / result["3m_price"].shift(1))

    return result


# ═══════════════════════════════════════════════
# COPPER F1 CONTINUOUS LOADER (module-level)
# ═══════════════════════════════════════════════

LOCAL_F1_CONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "LME_Copper_Rolling_F1_v2.csv")


@st.cache_data(show_spinner=False)
def _load_copper_f1_data() -> pd.DataFrame:
    """Load pre-computed LME Copper rolling F1 continuous series from CSV.

    Returns a DataFrame with columns F1_raw and F1_continuous, indexed by Date.
    Returns an empty DataFrame if the file is not found.
    """
    if not os.path.exists(LOCAL_F1_CONT_PATH):
        return pd.DataFrame()
    df = pd.read_csv(LOCAL_F1_CONT_PATH, parse_dates=["Date"]).set_index("Date")
    df.index = df.index.normalize()
    df = df.sort_index()
    return df[["F1_raw", "F1_continuous"]]


# ═══════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════

with st.sidebar:
    st.markdown('<p class="main-title">⚙️ Metals Dashboard</p>', unsafe_allow_html=True)
    st.markdown('<p class="main-subtitle">Risk Premia & Market Structure</p>', unsafe_allow_html=True)
    st.divider()

    st.markdown("##### 📂 Data Files")

    # Auto-load local files; show uploaders as optional overrides
    cash_local_exists = os.path.exists(LOCAL_CASH_PATH)
    curve_local_exists = os.path.exists(LOCAL_CURVE_PATH)

    if cash_local_exists:
        st.success("✓ Metals Cash and 3M.xlsx (local)")
        cash_file_override = st.file_uploader("Override Cash & 3M file", type=["xlsx", "xls"], key="cash")
        cash_file = cash_file_override if cash_file_override else _local_bytesio(LOCAL_CASH_PATH)
    else:
        cash_file = st.file_uploader("Metals Cash and 3M", type=["xlsx", "xls"], key="cash")

    if curve_local_exists:
        st.success("✓ Metals Futures Curve.csv (local)")
        curve_file_override = st.file_uploader("Override Futures Curve file", type=["xlsx", "xls", "csv", "xlsm"], key="curve")
        curve_file = curve_file_override if curve_file_override else _local_bytesio(LOCAL_CURVE_PATH)
    else:
        curve_file = st.file_uploader("Metals Futures Curve", type=["xlsx", "xls", "csv", "xlsm"], key="curve")

    st.divider()

    st.markdown("##### 📅 Date Range")

    LME_METALS = ["Copper", "Aluminium", "Zinc", "Nickel", "Lead", "Tin"]

    if cash_file:
        cash_data = load_cash_3m_data(cash_file)
        available_metals = [m for m in LME_METALS if m in cash_data]
    else:
        available_metals = LME_METALS
        cash_data = {}

    DATE_CAP = pd.Timestamp("2025-12-31").date()

    if cash_data and available_metals:
        df_dates = cash_data[available_metals[0]]
        min_date = df_dates.index.min().date()
        max_date = min(df_dates.index.max().date(), DATE_CAP)
    else:
        min_date = pd.Timestamp("2006-01-01").date()
        max_date = DATE_CAP

    date_range = st.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=DATE_CAP
    )

    if len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = min_date, max_date

    st.divider()
    st.caption("NYU Financial Engineering")
    st.caption("Metals Risk Premia Project")


# ═══════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════

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


def _find_curve_sheet(metal_name, curve_data):
    """Return the best-matching sheet name in curve_data for a given metal."""
    clean = (metal_name.lower()
             .replace("($/oz)", "").replace("($/lb)", "")
             .replace("comex", "").strip())
    # exact-ish match first
    for sheet_name in curve_data:
        sl = sheet_name.lower()
        if clean in sl or sl in clean:
            return sheet_name
    # fallback: first word of clean name
    first = clean.split()[0] if clean.split() else ""
    for sheet_name in curve_data:
        if first and first in sheet_name.lower():
            return sheet_name
    return None


def _get_curve_price(metal_name, curve_data, contract, start_date, end_date):
    """Return filtered price series for a specific futures contract (e.g. 'F3')."""
    sheet = _find_curve_sheet(metal_name, curve_data)
    if sheet is None:
        return pd.Series(dtype=float)
    prices = curve_data[sheet].get("prices", pd.DataFrame())
    if prices.empty or contract not in prices.columns:
        return pd.Series(dtype=float)
    s = prices[contract].dropna()
    s = s[s > 0]
    if s.empty:
        return pd.Series(dtype=float)
    tmp = filter_date(pd.DataFrame({contract: s}), start_date, end_date)
    return tmp[contract] if contract in tmp.columns else pd.Series(dtype=float)


def _get_curve_raw_col(metal_name, curve_data, contract, keyword, start_date, end_date):
    """Return a raw column (volume or OI) for a specific contract from curve_data."""
    sheet = _find_curve_sheet(metal_name, curve_data)
    if sheet is None:
        return pd.Series(dtype=float)
    raw = curve_data[sheet].get("raw", pd.DataFrame())
    if raw.empty:
        return pd.Series(dtype=float)
    contract_lower = contract.lower()
    for col in raw.columns:
        cl = col.lower().replace(" ", "_")
        if cl.startswith(f"{contract_lower}_") and keyword in cl:
            s = pd.to_numeric(raw[col], errors="coerce").dropna()
            if not s.empty:
                tmp = filter_date(pd.DataFrame({"v": s}), start_date, end_date)
                return tmp["v"] if "v" in tmp.columns else pd.Series(dtype=float)
    return pd.Series(dtype=float)


# ═══════════════════════════════════════════════
# MAIN CONTENT
# ═══════════════════════════════════════════════

if not cash_file:
    st.markdown("## 📂 Upload Data to Begin")
    st.info("Upload **Metals Cash and 3M.xlsx** and optionally **Metals Futures Curve** file using the sidebar to explore the dashboard.")
    st.markdown("""
    **Expected file structure:**

    **File 1 — Metals Cash and 3M.xlsx:**
    One sheet per LME metal (LME Copper, LME Aluminium, ...) with columns for
    Cash Price, 3M Forward Price/Volume/OI, Cash-3M Spread Price/Volume.
    Plus a CME Cash Prices sheet for Gold, Silver, Platinum, Palladium, Copper ($/lb).

    **File 2 — Metals Futures Curve (.xlsx or .csv):**
    One sheet per metal with F1 through F27, each having Price, Volume, Open Interest columns.
    """)
    st.stop()


# Load futures curve data if available
curve_data = {}
if curve_file:
    curve_data = load_futures_curve_data(curve_file)

# Build extended metal list: LME metals + CME column names
CME_METALS_LIST = list(cash_data["CME_Cash"].columns) if "CME_Cash" in cash_data else []
ALL_METALS = available_metals + CME_METALS_LIST


# ═══════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📊 Market Overview",
    "📈 Term Structure",
    "💰 Cash vs 3M (Carry)",
    "📉 Volume & Open Interest",
    "🔗 Copper LME-CME Spread",
    "📋 Statistics",
    "⚡ Momentum Signals",
    "📐 Carry Signals",
])


# ══════════════════════════════════════════════════════
# TAB 1: MARKET OVERVIEW
# ══════════════════════════════════════════════════════

with tab1:
    st.markdown("### Market Overview")
    st.caption("Latest snapshot across all metals")

    summary_rows = []
    for metal in LME_METALS:
        if metal not in cash_data:
            continue
        mdf = parse_cash_3m_columns(cash_data[metal], metal)
        mdf = filter_date(mdf, start_date, end_date)
        if mdf.empty:
            continue

        last = mdf.iloc[-1]
        prev = mdf.iloc[-2] if len(mdf) > 1 else mdf.iloc[-1]

        cash_p = last.get("cash_price", np.nan)
        tm_p = last.get("3m_price", np.nan)
        spread = last.get("spread_price", np.nan)
        cash_chg = ((cash_p / prev.get("cash_price", np.nan)) - 1) * 100 if pd.notna(prev.get("cash_price")) else 0

        summary_rows.append({
            "Metal": metal,
            "Cash": cash_p,
            "3M Forward": tm_p,
            "Cash-3M Spread": spread,
            "Daily Chg (%)": cash_chg,
            "Structure": "Backwardation" if (pd.notna(spread) and spread > 0) else "Contango",
        })

    if summary_rows:
        cols = st.columns(min(len(summary_rows), 6))
        for i, row in enumerate(summary_rows):
            with cols[i % len(cols)]:
                badge = "backwardation" if row["Structure"] == "Backwardation" else "contango"
                metric_card(
                    row["Metal"],
                    f"${row['Cash']:,.0f}" if pd.notna(row["Cash"]) else "N/A",
                    row["Daily Chg (%)"] if pd.notna(row["Daily Chg (%)"]) else None,
                    ""
                )
                st.markdown(
                    f'<span class="badge-{badge}">{row["Structure"]}</span>',
                    unsafe_allow_html=True
                )

        st.markdown("")

        if "CME_Cash" in cash_data:
            section_header("Precious Metals & COMEX Copper (Latest)")
            cme = cash_data["CME_Cash"]
            cme = filter_date(cme, start_date, end_date)
            if not cme.empty:
                last_cme = cme.iloc[-1]
                cme_cols = st.columns(min(len(cme.columns), 5))
                for j, col_name in enumerate(cme.columns):
                    with cme_cols[j % len(cme_cols)]:
                        val = last_cme[col_name]
                        if "lb" in col_name.lower():
                            label = col_name.replace("($/lb)", "").strip()
                            unit_str = " $/lb"
                        else:
                            label = col_name.replace("($/oz)", "").strip()
                            unit_str = " $/oz"
                        if pd.notna(val):
                            metric_card(label, f"${val:,.2f}", unit=unit_str)

        st.divider()

        section_header("Cash-3M Spread — 1 Year")
        spread_metals = [m for m in LME_METALS if m in cash_data]
        selected_spread_metal = st.selectbox("Select Commodity", spread_metals, key="spread_metal_select")

        mdf_spread = parse_cash_3m_columns(cash_data[selected_spread_metal], selected_spread_metal)
        mdf_spread = filter_date(mdf_spread, start_date, end_date)

        if "spread_price" in mdf_spread.columns and not mdf_spread.empty:
            last_1y = mdf_spread["spread_price"].dropna().tail(252)
            is_positive = last_1y.iloc[-1] > 0

            fig_spread = go.Figure()
            fig_spread.add_trace(go.Scatter(
                x=last_1y.index, y=last_1y.values,
                mode="lines",
                fill="tozeroy",
                fillcolor="rgba(52,211,153,0.12)" if is_positive else "rgba(248,113,113,0.12)",
                line=dict(
                    color=COLORS["green"] if is_positive else COLORS["red"],
                    width=2
                ),
                name=selected_spread_metal,
                hovertemplate="%{x|%b %d, %Y}<br>Spread: $%{y:,.2f}<extra></extra>"
            ))
            fig_spread.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
            fig_spread.update_layout(
                **CHART_LAYOUT,
                height=320,
                title=dict(text=f"{selected_spread_metal} — Cash-3M Spread (Last 1 Year)", font=dict(size=14)),
                yaxis_title="Spread ($/MT)",
                xaxis_title=None,
                hovermode="x unified",
            )
            fig_spread.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
            fig_spread.update_yaxes(showspikes=True, spikecolor="#475569", spikethickness=1)
            st.plotly_chart(fig_spread, use_container_width=True)

            c1, c2, c3 = st.columns(3)
            backw_pct = (last_1y > 0).sum() / len(last_1y) * 100
            with c1:
                metric_card("Last Spread", f"${last_1y.iloc[-1]:,.1f}")
            with c2:
                metric_card("Backwardation", f"{backw_pct:.1f}%")
            with c3:
                metric_card("1Y Avg Spread", f"${last_1y.mean():,.1f}")


# ══════════════════════════════════════════════════════
# TAB 2: TERM STRUCTURE
# ══════════════════════════════════════════════════════

with tab2:
    st.markdown("### Term Structure (Futures Curve)")

    if not curve_data:
        st.info("Upload the **Metals Futures Curve** file to view term structure analysis.")
    else:
        curve_metals = list(curve_data.keys())
        if curve_metals:
            curve_metal = st.selectbox("Select Metal (Curve)", curve_metals, key="curve_metal")

            if curve_metal in curve_data and "prices" in curve_data[curve_metal]:
                prices_df = curve_data[curve_metal]["prices"]

                if not prices_df.empty:
                    if not isinstance(prices_df.index, pd.DatetimeIndex):
                        raw_df = curve_data[curve_metal].get("raw", pd.DataFrame())
                        if isinstance(raw_df.index, pd.DatetimeIndex):
                            prices_df.index = raw_df.index[:len(prices_df)]
                        else:
                            try:
                                prices_df.index = pd.to_datetime(prices_df.index, errors="coerce")
                                prices_df = prices_df[prices_df.index.notna()]
                            except Exception:
                                st.warning("Could not parse dates from futures curve data.")
                                prices_df = pd.DataFrame()

                prices_df = filter_date(prices_df, start_date, end_date)

                if not prices_df.empty and not prices_df.columns.empty:
                    available_dates = prices_df.dropna(how="all").index
                    if len(available_dates) > 0:
                        col1, col2 = st.columns([2, 1])

                        with col1:
                            min_d = available_dates.min().date()
                            max_d = available_dates.max().date()
                            picked = st.date_input(
                                "Select Date", value=max_d,
                                min_value=min_d, max_value=max_d,
                                key="curve_date_select",
                            )
                            # Snap to nearest available trading date
                            picked_ts = pd.Timestamp(picked)
                            nearest_idx = available_dates.get_indexer([picked_ts], method="nearest")[0]
                            selected_curve_date = available_dates[nearest_idx]

                        with col2:
                            st.markdown("##### Compare With")
                            compare_options = ["None", "1 Month Ago", "3 Months Ago", "6 Months Ago", "1 Year Ago", "2 Years Ago"]
                            compare_choice = st.selectbox("Historical snapshot", compare_options, index=0, key="compare_choice")

                        # Resolve comparison date
                        offsets = {
                            "1 Month Ago": pd.DateOffset(months=1),
                            "3 Months Ago": pd.DateOffset(months=3),
                            "6 Months Ago": pd.DateOffset(months=6),
                            "1 Year Ago": pd.DateOffset(years=1),
                            "2 Years Ago": pd.DateOffset(years=2),
                        }
                        compare_dates = [selected_curve_date]
                        if compare_choice != "None" and compare_choice in offsets:
                            target_dt = selected_curve_date - offsets[compare_choice]
                            nearest = available_dates[available_dates.get_indexer([target_dt], method="nearest")[0]]
                            if nearest != selected_curve_date:
                                compare_dates.append(nearest)

                        compare_dates = sorted(set(compare_dates))

                        fig = go.Figure()
                        compare_colors = [COLORS["primary"], COLORS["amber"], COLORS["accent"],
                                          COLORS["pink"], COLORS["green"]]

                        for k, dt in enumerate(compare_dates):
                            row = prices_df.loc[dt].dropna()
                            row = row[row > 0]  # drop zero/invalid prices
                            if row.empty:
                                continue
                            # Sort contracts numerically: F1 < F2 < ... < F27
                            try:
                                row = row.reindex(sorted(
                                    row.index,
                                    key=lambda c: int(c.upper().replace("F", "") or "0")
                                ))
                            except Exception:
                                pass

                            is_latest = (dt == selected_curve_date)
                            fig.add_trace(go.Scatter(
                                x=list(row.index),
                                y=row.values,
                                mode="lines+markers",
                                name=dt.strftime("%Y-%m-%d"),
                                line=dict(
                                    color=compare_colors[k % len(compare_colors)],
                                    width=3 if is_latest else 1.5,
                                ),
                                marker=dict(size=6 if is_latest else 4),
                                opacity=1 if is_latest else 0.6,
                                hovertemplate="%{x}: $%{y:,.2f}<extra>" + dt.strftime("%b %d, %Y") + "</extra>"
                            ))

                        fig.update_layout(
                            **CHART_LAYOUT,
                            height=500,
                            title=dict(text=f"{curve_metal} — Forward Curve", font=dict(size=16)),
                            xaxis_title="Contract",
                            yaxis_title="Price",
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        latest_row = prices_df.loc[selected_curve_date].dropna()
                        if len(latest_row) >= 2:
                            slope = latest_row.iloc[-1] - latest_row.iloc[0]
                            if slope > 0:
                                st.success(f"📈 **Contango** — Far month contracts are trading higher than near month ({curve_metal}, {selected_curve_date.strftime('%Y-%m-%d')})")
                            else:
                                st.warning(f"📉 **Backwardation** — Near month contracts are trading higher than far month ({curve_metal}, {selected_curve_date.strftime('%Y-%m-%d')})")
                else:
                    st.warning("Could not parse futures price columns. Check column naming (expecting F1, F2, ... pattern with Price).")
            else:
                st.warning(f"No price data found for {curve_metal}")


# ══════════════════════════════════════════════════════
# TAB 3: CASH VS 3M (CARRY)
# ══════════════════════════════════════════════════════

with tab3:
    selected_metal = st.selectbox("Select Metal", ALL_METALS, key="tab3_metal")

    # Load data: LME metals have their own dict key; CME metals live in CME_Cash columns
    if selected_metal in cash_data:
        metal_df = parse_cash_3m_columns(cash_data[selected_metal], selected_metal)
        metal_df = filter_date(metal_df, start_date, end_date)
    elif "CME_Cash" in cash_data and selected_metal in cash_data["CME_Cash"].columns:
        cme_raw = filter_date(cash_data["CME_Cash"], start_date, end_date)
        cash_s = pd.to_numeric(cme_raw[selected_metal], errors="coerce").dropna()
        metal_df = pd.DataFrame({"cash_price": cash_s})
        if curve_data:
            f3_p = _get_curve_price(selected_metal, curve_data, "F3", start_date, end_date)
            if not f3_p.empty:
                metal_df = metal_df.join(f3_p.rename("3m_price"), how="left")
                metal_df["spread_price"] = metal_df["cash_price"] - metal_df["3m_price"]
        metal_df["cash_return"] = np.log(metal_df["cash_price"] / metal_df["cash_price"].shift(1))
        metal_df = metal_df.dropna(subset=["cash_price"])
    else:
        st.warning(f"No data found for {selected_metal}")
        st.stop()

    st.markdown(f"### {selected_metal} — Cash vs 3M (Carry Analysis)")

    if "cash_price" not in metal_df.columns or metal_df.empty:
        st.warning("Price data not found for this metal.")
    else:
        has_3m = "3m_price" in metal_df.columns

        section_header("Price Comparison")
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=metal_df.index, y=metal_df["cash_price"],
            name="Cash (Spot)", line=dict(color=COLORS["amber"], width=2),
            hovertemplate="%{x|%b %d, %Y}<br>Cash: $%{y:,.2f}<extra></extra>"
        ))

        if has_3m:
            fig.add_trace(go.Scatter(
                x=metal_df.index, y=metal_df["3m_price"],
                name="3M Forward", line=dict(color=COLORS["primary"], width=2),
                hovertemplate="%{x|%b %d, %Y}<br>3M: $%{y:,.2f}<extra></extra>"
            ))

        fig.update_layout(
            **CHART_LAYOUT,
            height=400,
            title=dict(
                text=f"{selected_metal}: Cash vs 3M Forward" if has_3m else f"{selected_metal}: Cash Price",
                font=dict(size=14)
            ),
            yaxis_title="Price",
        )
        st.plotly_chart(fig, use_container_width=True)

        if not has_3m:
            st.info("F3 futures data not found for this CME metal in the futures curve file. Only cash price is available.")
        else:
            if "spread_price" in metal_df.columns:
                section_header("Cash-3M Spread")
                spread = metal_df["spread_price"].dropna()
                fig_sp = go.Figure()
                fig_sp.add_trace(go.Scatter(
                    x=spread.index, y=spread.values,
                    mode="lines",
                    fill="tozeroy",
                    fillcolor="rgba(52,211,153,0.10)",
                    line=dict(color=COLORS["green"], width=1.8),
                    name="Cash-3M Spread",
                    hovertemplate="%{x|%b %d, %Y}<br>Spread: $%{y:,.2f}<extra></extra>"
                ))
                fig_sp.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
                fig_sp.update_layout(
                    **CHART_LAYOUT,
                    height=350,
                    title=dict(text=f"{selected_metal} — Cash minus 3M Spread", font=dict(size=14)),
                    yaxis_title="Spread ($/MT)",
                )
                fig_sp.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
                fig_sp.update_yaxes(showspikes=True, spikecolor="#475569", spikethickness=1)
                st.plotly_chart(fig_sp, use_container_width=True)

            section_header("Annualized Carry (%)")
            if "spread_price" in metal_df.columns:
                carry_pct = (metal_df["spread_price"] / metal_df["3m_price"]) * (365 / 90) * 100
                carry_pct = carry_pct.dropna()

                fig_carry = go.Figure()
                fig_carry.add_trace(go.Scatter(
                    x=carry_pct.index, y=carry_pct.values,
                    fill="tozeroy",
                    fillcolor="rgba(59,130,246,0.1)",
                    line=dict(color=COLORS["primary"], width=1.5),
                    hovertemplate="%{x|%b %d, %Y}<br>Carry: %{y:.2f}%<extra></extra>"
                ))
                fig_carry.add_hline(y=0, line_dash="dash", line_color="#475569")
                fig_carry.update_layout(
                    **CHART_LAYOUT,
                    height=350,
                    title=dict(text="Annualized Carry (Spread / 3M × 365/90)", font=dict(size=13)),
                    yaxis_title="Carry (%)",
                )
                st.plotly_chart(fig_carry, use_container_width=True)

            section_header("Spread Distribution")
            if "spread_price" in metal_df.columns:
                spread_data = metal_df["spread_price"].dropna()
                col1, col2, col3 = st.columns(3)
                backw_pct = (spread_data > 0).sum() / len(spread_data) * 100
                with col1:
                    metric_card("Backwardation", f"{backw_pct:.1f}%")
                with col2:
                    metric_card("Contango", f"{100 - backw_pct:.1f}%")
                with col3:
                    metric_card("Avg Spread", f"${spread_data.mean():,.1f}")

                fig_hist = go.Figure()
                fig_hist.add_trace(go.Histogram(
                    x=spread_data.values,
                    nbinsx=80,
                    marker_color=COLORS["primary"],
                    opacity=0.7,
                    hovertemplate="Spread: $%{x:,.1f}<br>Count: %{y}<extra></extra>"
                ))
                fig_hist.add_vline(x=0, line_dash="dash", line_color=COLORS["amber"], line_width=2)
                fig_hist.update_layout(
                    **CHART_LAYOUT,
                    height=300,
                    title=dict(text="Distribution of Cash-3M Spread", font=dict(size=13)),
                    xaxis_title="Spread ($/MT)",
                    yaxis_title="Frequency",
                )
                st.plotly_chart(fig_hist, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 4: VOLUME & OPEN INTEREST
# ══════════════════════════════════════════════════════

with tab4:
    selected_metal = st.selectbox("Select Metal", ALL_METALS, key="tab4_metal")

    if selected_metal in cash_data:
        metal_df = parse_cash_3m_columns(cash_data[selected_metal], selected_metal)
        metal_df = filter_date(metal_df, start_date, end_date)
    elif "CME_Cash" in cash_data and selected_metal in cash_data["CME_Cash"].columns:
        cme_raw = filter_date(cash_data["CME_Cash"], start_date, end_date)
        cash_s = pd.to_numeric(cme_raw[selected_metal], errors="coerce").dropna()
        metal_df = pd.DataFrame({"cash_price": cash_s})
        if curve_data:
            f3_p = _get_curve_price(selected_metal, curve_data, "F3", start_date, end_date)
            if not f3_p.empty:
                metal_df = metal_df.join(f3_p.rename("3m_price"), how="left")
        metal_df = metal_df.dropna(subset=["cash_price"])
    else:
        st.warning(f"No data found for {selected_metal}")
        st.stop()

    st.markdown(f"### {selected_metal} — Volume & Open Interest")

    # Price chart: 3M Forward for LME; F3 front month for CME
    is_lme = selected_metal in cash_data
    if "3m_price" in metal_df.columns:
        price_series = metal_df["3m_price"].dropna()
        price_label = "3M Forward Price" if is_lme else "F3 Price"
        price_color = COLORS["primary"] if is_lme else COLORS["amber"]
    else:
        price_series = pd.Series(dtype=float)
        price_label = ""
        price_color = COLORS["primary"]

    if not price_series.empty:
        section_header(f"{price_label}")
        fig_price = go.Figure()
        fig_price.add_trace(go.Scatter(
            x=price_series.index, y=price_series.values,
            name=price_label,
            line=dict(color=price_color, width=2),
            fill="tozeroy",
            fillcolor=f"rgba({int(price_color[1:3],16)},{int(price_color[3:5],16)},{int(price_color[5:7],16)},0.1)",
            hovertemplate="%{x|%b %d, %Y}<br>" + price_label + ": $%{y:,.2f}<extra></extra>"
        ))
        fig_price.update_layout(
            **CHART_LAYOUT, height=350,
            title=dict(text=f"{selected_metal} — {price_label}", font=dict(size=14)),
            yaxis_title="Price",
        )
        fig_price.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        fig_price.update_yaxes(showspikes=True, spikecolor="#475569", spikethickness=1)
        st.plotly_chart(fig_price, use_container_width=True)

    if is_lme:
        # LME: V/OI from metal_df (3m_volume, 3m_oi)
        has_vol = "3m_volume" in metal_df.columns
        has_oi = "3m_oi" in metal_df.columns

        if not has_vol and not has_oi:
            st.info("3M Forward Volume and Open Interest data not available for this metal.")
        else:
            section_header("3M Forward — Volume & Open Interest")
            fig_vol = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                    row_heights=[0.5, 0.5], vertical_spacing=0.08)

            if has_vol:
                vol = metal_df["3m_volume"].dropna()
                fig_vol.add_trace(go.Bar(
                    x=vol.index, y=vol.values,
                    name="Volume", marker_color=COLORS["primary"], opacity=0.6,
                    hovertemplate="%{x|%b %d, %Y}<br>Vol: %{y:,.0f}<extra></extra>"
                ), row=1, col=1)
                vol_ma = vol.rolling(20).mean()
                fig_vol.add_trace(go.Scatter(
                    x=vol_ma.index, y=vol_ma.values,
                    name="20D Avg", line=dict(color=COLORS["amber"], width=2),
                ), row=1, col=1)

            if has_oi:
                oi = metal_df["3m_oi"].dropna()
                fig_vol.add_trace(go.Scatter(
                    x=oi.index, y=oi.values,
                    name="Open Interest", line=dict(color=COLORS["accent"], width=2),
                    fill="tozeroy", fillcolor="rgba(6,182,212,0.1)",
                    hovertemplate="%{x|%b %d, %Y}<br>OI: %{y:,.0f}<extra></extra>"
                ), row=2, col=1)

            fig_vol.update_layout(
                **CHART_LAYOUT,
                height=500,
                title=dict(text=f"{selected_metal} — 3M Forward", font=dict(size=14)),
            )
            fig_vol.update_yaxes(title_text="Volume", row=1, col=1)
            fig_vol.update_yaxes(title_text="Open Interest", row=2, col=1)
            st.plotly_chart(fig_vol, use_container_width=True)

    else:
        # CME: contract-month dropdown → V/OI from curve data
        if not curve_data:
            st.info("Futures curve data not available for Volume & Open Interest.")
        else:
            cme_sheet = _find_curve_sheet(selected_metal, curve_data)
            if not cme_sheet:
                st.info(f"No futures curve data found for {selected_metal}.")
            else:
                raw_all = filter_date(curve_data[cme_sheet].get("raw", pd.DataFrame()), start_date, end_date)
                avail_contracts = []
                for i in range(1, 28):
                    for col in raw_all.columns:
                        cl = col.lower().replace(" ", "_")
                        if cl.startswith(f"f{i}_") and "volume" in cl:
                            avail_contracts.append(f"F{i}")
                            break

                if not avail_contracts:
                    st.info("No volume data found in futures curve file for this metal.")
                else:
                    sel_contract = st.selectbox(
                        "Select Contract Month",
                        avail_contracts,
                        index=min(2, len(avail_contracts) - 1),
                        key="tab4_contract"
                    )
                    sel_vol = _get_curve_raw_col(selected_metal, curve_data, sel_contract, "volume", start_date, end_date)
                    sel_oi = _get_curve_raw_col(selected_metal, curve_data, sel_contract, "open", start_date, end_date)

                    if sel_vol.empty and sel_oi.empty:
                        st.info(f"No volume/OI data found for {sel_contract}.")
                    else:
                        section_header(f"{sel_contract} — Volume & Open Interest")
                        fig_vol_c = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                                  row_heights=[0.5, 0.5], vertical_spacing=0.08)

                        if not sel_vol.empty:
                            fig_vol_c.add_trace(go.Bar(
                                x=sel_vol.index, y=sel_vol.values,
                                name="Volume", marker_color=COLORS["primary"], opacity=0.6,
                                hovertemplate="%{x|%b %d, %Y}<br>Vol: %{y:,.0f}<extra></extra>"
                            ), row=1, col=1)
                            cme_vol_ma = sel_vol.rolling(20).mean()
                            fig_vol_c.add_trace(go.Scatter(
                                x=cme_vol_ma.index, y=cme_vol_ma.values,
                                name="20D Avg", line=dict(color=COLORS["amber"], width=2),
                            ), row=1, col=1)

                        if not sel_oi.empty:
                            fig_vol_c.add_trace(go.Scatter(
                                x=sel_oi.index, y=sel_oi.values,
                                name="Open Interest", line=dict(color=COLORS["accent"], width=2),
                                fill="tozeroy", fillcolor="rgba(6,182,212,0.1)",
                                hovertemplate="%{x|%b %d, %Y}<br>OI: %{y:,.0f}<extra></extra>"
                            ), row=2, col=1)

                        fig_vol_c.update_layout(
                            **CHART_LAYOUT,
                            height=500,
                            title=dict(text=f"{selected_metal} — {sel_contract}", font=dict(size=14)),
                        )
                        fig_vol_c.update_yaxes(title_text="Volume", row=1, col=1)
                        fig_vol_c.update_yaxes(title_text="Open Interest", row=2, col=1)
                        st.plotly_chart(fig_vol_c, use_container_width=True)

    # Futures Strip Volume Heatmap — works for both LME and CME
    if curve_data:
        section_header("Futures Strip Volume Heatmap")
        curve_match = _find_curve_sheet(selected_metal, curve_data)

        if curve_match and "raw" in curve_data[curve_match]:
            raw_curve = curve_data[curve_match]["raw"]
            raw_curve = filter_date(raw_curve, start_date, end_date)

            vol_cols = [c for c in raw_curve.columns if "volume" in c.lower()]
            if vol_cols:
                vol_df = raw_curve[vol_cols].copy()
                vol_df.columns = [c.split("_")[0] if "_" in c else c for c in vol_df.columns]

                vol_monthly = vol_df.resample("ME").mean()
                vol_monthly = vol_monthly.tail(36)

                if not vol_monthly.empty:
                    fig_hm = go.Figure(data=go.Heatmap(
                        z=vol_monthly.values.T,
                        x=vol_monthly.index.strftime("%Y-%m"),
                        y=vol_monthly.columns,
                        colorscale="Viridis",
                        hovertemplate="Date: %{x}<br>Contract: %{y}<br>Avg Volume: %{z:,.0f}<extra></extra>"
                    ))
                    fig_hm.update_layout(
                        **CHART_LAYOUT,
                        height=400,
                        title=dict(text=f"{selected_metal} — Monthly Average Volume by Contract", font=dict(size=13)),
                        xaxis_title="Month",
                        yaxis_title="Contract",
                    )
                    st.plotly_chart(fig_hm, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 5: CROSS-METAL (LME Copper vs COMEX Copper)
# ══════════════════════════════════════════════════════

with tab5:
    st.markdown("### LME Copper vs COMEX Copper (HG)")
    st.caption("Location arbitrage: LME $/MT vs COMEX ¢/lb")

    has_lme_cu = "Copper" in cash_data
    has_cme_cu = "CME_Cash" in cash_data

    if not has_lme_cu or not has_cme_cu:
        st.info("Both **LME Copper** and **CME Cash Prices** sheets are needed for this analysis.")
    else:
        lme_cu = parse_cash_3m_columns(cash_data["Copper"], "Copper")
        lme_cu = filter_date(lme_cu, start_date, end_date)

        cme_cash = cash_data["CME_Cash"]
        cme_cash = filter_date(cme_cash, start_date, end_date)

        cu_cme_col = [c for c in cme_cash.columns if "copper" in c.lower() or "cu" in c.lower()]

        if not cu_cme_col:
            st.warning("Copper column not found in CME Cash Prices sheet.")
        else:
            cu_cme_col = cu_cme_col[0]
            cme_cu_price = cme_cash[cu_cme_col].dropna()

            LBS_PER_MT = 2204.62
            cme_cu_mt = cme_cu_price * LBS_PER_MT

            combined = pd.DataFrame({
                "LME_Cash": lme_cu["cash_price"] if "cash_price" in lme_cu.columns else lme_cu.get("3m_price"),
                "COMEX_MT": cme_cu_mt,
            }).dropna()

            if combined.empty:
                st.warning("No overlapping dates between LME and COMEX copper data.")
            else:
                combined["Spread"] = combined["LME_Cash"] - combined["COMEX_MT"]
                combined["Ratio"] = combined["LME_Cash"] / combined["COMEX_MT"]

                col1, col2, col3, col4 = st.columns(4)
                last = combined.iloc[-1]
                with col1:
                    metric_card("LME Cash", f"${last['LME_Cash']:,.0f}", unit=" /MT")
                with col2:
                    metric_card("COMEX (conv.)", f"${last['COMEX_MT']:,.0f}", unit=" /MT")
                with col3:
                    metric_card("Spread", f"${last['Spread']:,.0f}", unit=" /MT")
                with col4:
                    metric_card("Ratio", f"{last['Ratio']:.4f}")

                section_header("LME vs COMEX — Price in $/MT")
                fig_xm = go.Figure()

                fig_xm.add_trace(go.Scatter(
                    x=combined.index, y=combined["LME_Cash"],
                    name="LME Copper Cash", line=dict(color=COLORS["orange"], width=2),
                    hovertemplate="%{x|%b %d, %Y}<br>LME: $%{y:,.0f}/MT<extra></extra>"
                ))

                fig_xm.add_trace(go.Scatter(
                    x=combined.index, y=combined["COMEX_MT"],
                    name="COMEX Copper (conv. $/MT)", line=dict(color=COLORS["primary"], width=2),
                    hovertemplate="%{x|%b %d, %Y}<br>COMEX: $%{y:,.0f}/MT<extra></extra>"
                ))

                fig_xm.update_layout(
                    **CHART_LAYOUT, height=400,
                    title=dict(text="LME vs COMEX Copper", font=dict(size=14)),
                    yaxis_title="Price ($/MT)",
                )
                st.plotly_chart(fig_xm, use_container_width=True)

                section_header("LME Cash − COMEX (conv.) Spread")
                spread_vals = combined["Spread"]
                spread_pos = spread_vals.clip(lower=0)
                spread_neg = spread_vals.clip(upper=0)
                fig_spread_bar = go.Figure()
                # Filled area: positive (LME premium)
                fig_spread_bar.add_trace(go.Scatter(
                    x=spread_pos.index, y=spread_pos.values,
                    mode="none", fill="tozeroy",
                    fillcolor="rgba(52,211,153,0.25)",
                    name="LME Cash", showlegend=True,
                    hoverinfo="skip",
                ))
                # Filled area: negative (COMEX premium)
                fig_spread_bar.add_trace(go.Scatter(
                    x=spread_neg.index, y=spread_neg.values,
                    mode="none", fill="tozeroy",
                    fillcolor="rgba(248,113,113,0.25)",
                    name="CME Cash", showlegend=True,
                    hoverinfo="skip",
                ))
                # Main spread line (always visible)
                fig_spread_bar.add_trace(go.Scatter(
                    x=spread_vals.index, y=spread_vals.values,
                    mode="lines", name="Spread",
                    line=dict(color=COLORS["accent"], width=1.5),
                    hovertemplate="%{x|%b %d, %Y}<br>Spread: $%{y:,.0f}/MT<extra></extra>",
                    showlegend=False,
                ))
                fig_spread_bar.update_layout(
                    **CHART_LAYOUT, height=350,
                    title=dict(text="LME Cash − COMEX (conv.) Spread ($/MT)", font=dict(size=14)),
                    yaxis_title="Spread ($/MT)",
                )
                fig_spread_bar.add_hline(y=0, line_dash="dash", line_color="white", line_width=1, opacity=0.4)
                st.plotly_chart(fig_spread_bar, use_container_width=True)

                section_header("Rolling 60-Day Correlation")
                rolling_corr = combined["LME_Cash"].rolling(60).corr(combined["COMEX_MT"])

                corr_min = rolling_corr.dropna().min()
                corr_min_date = rolling_corr.dropna().idxmin()

                fig_corr = go.Figure()
                fig_corr.add_trace(go.Scatter(
                    x=rolling_corr.index, y=rolling_corr.values,
                    fill="tozeroy", fillcolor="rgba(139,92,246,0.15)",
                    line=dict(color=COLORS["secondary"], width=2),
                    hovertemplate="%{x|%b %d, %Y}<br>Correlation: %{y:.4f}<extra></extra>"
                ))
                # Mark the minimum point
                fig_corr.add_trace(go.Scatter(
                    x=[corr_min_date], y=[corr_min],
                    mode="markers+text",
                    marker=dict(color=COLORS["red"], size=9, symbol="circle"),
                    text=[f"  Min: {corr_min:.4f}"],
                    textposition="middle right",
                    textfont=dict(color=COLORS["red"], size=11),
                    hovertemplate=f"{corr_min_date.strftime('%b %d, %Y')}<br>Min Correlation: {corr_min:.4f}<extra></extra>",
                    showlegend=False,
                ))
                fig_corr.update_layout(
                    **CHART_LAYOUT, height=360,
                    title=dict(text="LME-COMEX Rolling Correlation (60D)", font=dict(size=13)),
                    yaxis_title="Correlation",
                    hovermode="x unified",
                )
                fig_corr.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
                fig_corr.update_yaxes(showspikes=True, spikecolor="#475569", spikethickness=1)
                st.plotly_chart(fig_corr, use_container_width=True)

                st.info(
                    "**Why the sharp drop?** The rolling correlation fell sharply around mid-2025 due to the "
                    "**US copper tariff shock** — COMEX copper (US domestic) priced in a large import tariff premium "
                    "and diverged from LME copper (global benchmark), temporarily breaking the historically tight "
                    "relationship. Any visible data gap reflects periods where CME settlement prices were unavailable. "
                    "The correlation recovered once the tariff premium stabilised."
                )

                section_header("Spread Statistics")
                col1, col2 = st.columns(2)
                with col1:
                    spread_stats = combined["Spread"].describe()
                    st.dataframe(spread_stats.to_frame("LME-COMEX Spread ($/MT)").style.format("{:,.2f}"))
                with col2:
                    fig_sp_hist = go.Figure()
                    fig_sp_hist.add_trace(go.Histogram(
                        x=combined["Spread"].values, nbinsx=60,
                        marker_color=COLORS["secondary"], opacity=0.7,
                    ))
                    fig_sp_hist.add_vline(x=0, line_dash="dash", line_color=COLORS["amber"])
                    fig_sp_hist.update_layout(
                        **CHART_LAYOUT, height=300,
                        title=dict(text="Spread Distribution", font=dict(size=13)),
                        xaxis_title="Spread ($/MT)",
                    )
                    st.plotly_chart(fig_sp_hist, use_container_width=True)


# ══════════════════════════════════════════════════════
# TAB 6: STATISTICS
# ══════════════════════════════════════════════════════

with tab6:
    st.markdown("### Descriptive Statistics")

    section_header("Summary Statistics — All LME Metals")

    stats_rows = []
    for metal in LME_METALS:
        if metal not in cash_data:
            continue
        mdf = parse_cash_3m_columns(cash_data[metal], metal)
        mdf = filter_date(mdf, start_date, end_date)
        if "cash_price" not in mdf.columns:
            continue

        cash = mdf["cash_price"].dropna()
        rets = mdf.get("cash_return", pd.Series(dtype=float)).dropna()
        spread = mdf.get("spread_price", pd.Series(dtype=float)).dropna()

        row = {
            "Metal": metal,
            "Obs": len(cash),
            "Start": cash.index.min().strftime("%Y-%m-%d") if len(cash) > 0 else "",
            "End": cash.index.max().strftime("%Y-%m-%d") if len(cash) > 0 else "",
            "Mean Price": cash.mean(),
            "Std Price": cash.std(),
            "Min Price": cash.min(),
            "Max Price": cash.max(),
        }

        if len(rets) > 10:
            row["Ann. Return"] = rets.mean() * 252 * 100
            row["Ann. Vol"] = rets.std() * np.sqrt(252) * 100
            row["Skew"] = rets.skew()
            row["Kurtosis"] = rets.kurtosis()

        if len(spread) > 0:
            row["Avg Spread"] = spread.mean()
            row["Backw. %"] = (spread > 0).sum() / len(spread) * 100

        stats_rows.append(row)

    if stats_rows:
        stats_df = pd.DataFrame(stats_rows).set_index("Metal")
        fmt_dict = {c: "{:,.2f}" for c in stats_df.columns if stats_df[c].dtype in ["float64", "float32"]}
        fmt_dict.update({"Obs": "{:,.0f}", "Backw. %": "{:.1f}%"})
        st.dataframe(stats_df.style.format(fmt_dict, na_rep="—"), use_container_width=True)

    section_header("Summary Statistics — CME Metals")
    if "CME_Cash" in cash_data:
        cme_df_all = filter_date(cash_data["CME_Cash"], start_date, end_date)
        cme_stats_rows = []
        for col in cme_df_all.columns:
            series = pd.to_numeric(cme_df_all[col], errors="coerce").dropna()
            if len(series) < 10:
                continue
            rets_c = np.log(series / series.shift(1)).dropna()
            row_c = {
                "Metal": col,
                "Obs": len(series),
                "Start": series.index.min().strftime("%Y-%m-%d"),
                "End": series.index.max().strftime("%Y-%m-%d"),
                "Mean Price": series.mean(),
                "Std Price": series.std(),
                "Min Price": series.min(),
                "Max Price": series.max(),
            }
            if len(rets_c) > 10:
                row_c["Ann. Return"] = rets_c.mean() * 252 * 100
                row_c["Ann. Vol"] = rets_c.std() * np.sqrt(252) * 100
                row_c["Skew"] = rets_c.skew()
                row_c["Kurtosis"] = rets_c.kurtosis()
            cme_stats_rows.append(row_c)

        if cme_stats_rows:
            cme_stats_df = pd.DataFrame(cme_stats_rows).set_index("Metal")
            fmt_cme = {c: "{:,.2f}" for c in cme_stats_df.columns if cme_stats_df[c].dtype in ["float64", "float32"]}
            fmt_cme.update({"Obs": "{:,.0f}"})
            st.dataframe(cme_stats_df.style.format(fmt_cme, na_rep="—"), use_container_width=True)
    else:
        st.info("CME Cash Prices data not available.")

    selected_metal = st.selectbox("Select Metal for Detailed Analysis", available_metals, key="tab6_metal")
    if selected_metal in cash_data:
        metal_df = parse_cash_3m_columns(cash_data[selected_metal], selected_metal)
        metal_df = filter_date(metal_df, start_date, end_date)
    else:
        metal_df = pd.DataFrame()

    section_header(f"{selected_metal} — Rolling Volatility")
    if not metal_df.empty and "cash_return" in metal_df.columns:
        rets = metal_df["cash_return"].dropna()

        fig_vol = go.Figure()
        for window, color, wname in [(30, COLORS["primary"], "30D"), (60, COLORS["amber"], "60D"), (90, COLORS["accent"], "90D")]:
            rv = rets.rolling(window).std() * np.sqrt(252) * 100
            fig_vol.add_trace(go.Scatter(
                x=rv.index, y=rv.values,
                name=wname, line=dict(color=color, width=1.5),
                hovertemplate="%{x|%b %d, %Y}<br>" + wname + ": %{y:.1f}%<extra></extra>"
            ))

        fig_vol.update_layout(
            **CHART_LAYOUT, height=400,
            title=dict(text=f"{selected_metal} — Annualized Realized Volatility", font=dict(size=14)),
            yaxis_title="Volatility (%)",
        )
        st.plotly_chart(fig_vol, use_container_width=True)

    section_header(f"{selected_metal} — Return Distribution")
    if not metal_df.empty and "cash_return" in metal_df.columns:
        rets = metal_df["cash_return"].dropna()

        col1, col2 = st.columns([2, 1])
        with col1:
            fig_rd = go.Figure()
            fig_rd.add_trace(go.Histogram(
                x=rets.values * 100, nbinsx=100,
                marker_color=COLORS["primary"], opacity=0.7,
                name="Daily Returns",
            ))
            fig_rd.add_vline(x=0, line_dash="dash", line_color=COLORS["amber"])
            fig_rd.update_layout(
                **CHART_LAYOUT, height=350,
                title=dict(text="Daily Log Return Distribution (%)", font=dict(size=13)),
                xaxis_title="Return (%)",
                yaxis_title="Frequency",
            )
            st.plotly_chart(fig_rd, use_container_width=True)

        with col2:
            st.markdown("##### Return Statistics")
            ret_stats = {
                "Mean (daily)": f"{rets.mean() * 100:.4f}%",
                "Std (daily)": f"{rets.std() * 100:.4f}%",
                "Skewness": f"{rets.skew():.4f}",
                "Kurtosis": f"{rets.kurtosis():.4f}",
                "Min": f"{rets.min() * 100:.2f}%",
                "Max": f"{rets.max() * 100:.2f}%",
                "Ann. Return": f"{rets.mean() * 252 * 100:.2f}%",
                "Ann. Volatility": f"{rets.std() * np.sqrt(252) * 100:.2f}%",
            }
            for k, v in ret_stats.items():
                st.markdown(f"**{k}:** `{v}`")


# ══════════════════════════════════════════════════════
# MOMENTUM: comparison helper (module-level)
# ══════════════════════════════════════════════════════

def _mom_cum_pnl(f1r: pd.Series, f1c: pd.Series, spec: dict) -> pd.Series:
    """Compute gross cumulative PnL for a momentum variant spec dict."""
    def _ew(s, n): return s.ewm(com=n - 1, adjust=False).mean()
    t  = spec["type"]
    sd = spec.get("same_day", False)
    if t == "ma":
        sig = np.sign(f1r.rolling(spec["m"]).mean() - f1r.rolling(spec["n"]).mean()).values.astype(float)
    elif t == "cta_paper":
        pv, us = f1r.rolling(63).std(), []
        for sk, lk in zip((8, 16, 32), (24, 48, 96)):
            x = _ew(f1r, sk) - _ew(f1r, lk); y = x / pv
            with np.errstate(invalid="ignore"):
                z = (y / y.rolling(252).std()).values
                us.append(z * np.exp(-z ** 2 / 4) / 0.89)
        sig = np.sign(np.nanmean(np.stack(us, axis=1), axis=1))
    else:
        x = _ew(f1r, spec["s"]) - _ew(f1r, spec["l"]); y = x / f1r.rolling(63).std()
        with np.errstate(invalid="ignore"):
            z = (y / y.rolling(252).std()).values
            sig = np.sign(z * np.exp(-z ** 2 / 4) / 0.89)
    T = len(sig); pos = np.empty(T)
    if sd:
        pos[:] = np.where(np.isfinite(sig), sig, 0.0)
    else:
        pos[0] = 0.0; pos[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
    return pd.Series(pos * f1c.diff().values.astype(float), index=f1r.index).cumsum()


# All variant+timing combinations available for comparison dropdown
_MOM_CMP_OPTIONS = {
    "N/A": None,
    "MA(35,43) — Lag-1":    {"type": "ma", "m": 35, "n": 43, "same_day": False},
    "MA(35,43) — Same-Day": {"type": "ma", "m": 35, "n": 43, "same_day": True},
    "MA(33,48) — Lag-1":    {"type": "ma", "m": 33, "n": 48, "same_day": False},
    "MA(33,48) — Same-Day": {"type": "ma", "m": 33, "n": 48, "same_day": True},
    "MA(35,44) — Lag-1":    {"type": "ma", "m": 35, "n": 44, "same_day": False},
    "MA(35,44) — Same-Day": {"type": "ma", "m": 35, "n": 44, "same_day": True},
    "MA(34,47) — Lag-1":    {"type": "ma", "m": 34, "n": 47, "same_day": False},
    "MA(34,47) — Same-Day": {"type": "ma", "m": 34, "n": 47, "same_day": True},
    "MA(36,44) — Lag-1":    {"type": "ma", "m": 36, "n": 44, "same_day": False},
    "MA(36,44) — Same-Day": {"type": "ma", "m": 36, "n": 44, "same_day": True},
    "MA(1,5) — Lag-1":      {"type": "ma", "m": 1,  "n": 5,  "same_day": False},
    "MA(1,5) — Same-Day":   {"type": "ma", "m": 1,  "n": 5,  "same_day": True},
    "MA(5,20) — Lag-1":     {"type": "ma", "m": 5,  "n": 20, "same_day": False},
    "MA(5,20) — Same-Day":  {"type": "ma", "m": 5,  "n": 20, "same_day": True},
    "MA(10,60) — Lag-1":    {"type": "ma", "m": 10, "n": 60, "same_day": False},
    "MA(10,60) — Same-Day": {"type": "ma", "m": 10, "n": 60, "same_day": True},
    "CTA(9,21) — Lag-1":    {"type": "cta_single", "s": 9,  "l": 21, "same_day": False},
    "CTA(9,21) — Same-Day": {"type": "cta_single", "s": 9,  "l": 21, "same_day": True},
    "CTA(9,20) — Lag-1":    {"type": "cta_single", "s": 9,  "l": 20, "same_day": False},
    "CTA(9,20) — Same-Day": {"type": "cta_single", "s": 9,  "l": 20, "same_day": True},
    "CTA(10,19) — Lag-1":   {"type": "cta_single", "s": 10, "l": 19, "same_day": False},
    "CTA(10,19) — Same-Day":{"type": "cta_single", "s": 10, "l": 19, "same_day": True},
    "CTA(8,21) — Lag-1":    {"type": "cta_single", "s": 8,  "l": 21, "same_day": False},
    "CTA(8,21) — Same-Day": {"type": "cta_single", "s": 8,  "l": 21, "same_day": True},
    "CTA(14,15) — Lag-1":   {"type": "cta_single", "s": 14, "l": 15, "same_day": False},
    "CTA(14,15) — Same-Day":{"type": "cta_single", "s": 14, "l": 15, "same_day": True},
    "CTA Paper — Lag-1":    {"type": "cta_paper", "same_day": False},
    "CTA Paper — Same-Day": {"type": "cta_paper", "same_day": True},
}

_MA_OPT_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "MA_Crossover_Optimization.csv")
_CTA_OPT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "CTA_Optimization.csv")


@st.cache_data(show_spinner=False)
def _load_ma_opt() -> pd.DataFrame:
    if not os.path.exists(_MA_OPT_PATH):
        return pd.DataFrame()
    return pd.read_csv(_MA_OPT_PATH)


@st.cache_data(show_spinner=False)
def _load_cta_opt() -> pd.DataFrame:
    if not os.path.exists(_CTA_OPT_PATH):
        return pd.DataFrame()
    return pd.read_csv(_CTA_OPT_PATH)


# ══════════════════════════════════════════════════════
# CARRY: signal helpers (module-level)
# ══════════════════════════════════════════════════════

def _carry_raw_signal(curve_prices: pd.DataFrame, cash_parsed: pd.DataFrame, spec: dict) -> pd.Series:
    """Return raw carry signal series (continuous value, not binarized) for a spec dict."""
    v = spec.get("variant", "v1")
    if v in ("v1", "v2"):
        st_type = spec.get("signal_type", "f1f2")
        if st_type == "cash3m":
            if cash_parsed.empty or "cash_price" not in cash_parsed.columns or "3m_price" not in cash_parsed.columns:
                return pd.Series(dtype=float)
            cp = cash_parsed["cash_price"].dropna()
            tm = cash_parsed["3m_price"].dropna()
            idx = cp.index.intersection(tm.index)
            raw = (cp.reindex(idx) - tm.reindex(idx)) / cp.reindex(idx)
        elif "f3" in st_type:
            if "F1" not in curve_prices.columns or "F3" not in curve_prices.columns:
                return pd.Series(dtype=float)
            f1 = curve_prices["F1"].dropna(); f3 = curve_prices["F3"].dropna()
            idx = f1.index.intersection(f3.index)
            raw = (f1.reindex(idx) - f3.reindex(idx)) / f1.reindex(idx)
        else:
            if "F1" not in curve_prices.columns or "F2" not in curve_prices.columns:
                return pd.Series(dtype=float)
            f1 = curve_prices["F1"].dropna(); f2 = curve_prices["F2"].dropna()
            idx = f1.index.intersection(f2.index)
            raw = (f1.reindex(idx) - f2.reindex(idx)) / f1.reindex(idx)
        if v == "v2":
            raw = raw * (6 if "f3" in st_type else 12)
        return raw.replace([np.inf, -np.inf], np.nan).dropna()
    elif v == "v3":
        j, k = spec.get("j", 3), spec.get("k", 15)
        fj_col, fk_col = f"F{j}", f"F{k}"
        if fj_col not in curve_prices.columns or fk_col not in curve_prices.columns:
            return pd.Series(dtype=float)
        fj = curve_prices[fj_col].dropna(); fk = curve_prices[fk_col].dropna()
        idx = fj.index.intersection(fk.index)
        raw = (fj.reindex(idx) - fk.reindex(idx)) / fk.reindex(idx)
        return raw.replace([np.inf, -np.inf], np.nan).dropna()
    elif v == "v4":
        window = spec.get("window", 252)
        if "F1" not in curve_prices.columns or "F2" not in curve_prices.columns:
            return pd.Series(dtype=float)
        f1 = curve_prices["F1"].dropna(); f2 = curve_prices["F2"].dropna()
        idx = f1.index.intersection(f2.index)
        base = (f1.reindex(idx) - f2.reindex(idx)) / f1.reindex(idx)
        z = (base - base.rolling(window).mean()) / base.rolling(window).std()
        return z.replace([np.inf, -np.inf], np.nan).dropna()
    return pd.Series(dtype=float)


def _carry_cum_pnl(curve_prices: pd.DataFrame, cash_parsed: pd.DataFrame,
                   f1c: pd.Series, spec: dict) -> pd.Series:
    """Gross cumulative PnL (USD/MT) for a carry spec dict."""
    cr = _carry_raw_signal(curve_prices, cash_parsed, spec)
    if cr.empty:
        return pd.Series(dtype=float)
    idx = cr.index.intersection(f1c.index)
    if len(idx) < 10:
        return pd.Series(dtype=float)
    cr = cr.reindex(idx); f1c_a = f1c.reindex(idx)
    sig = np.sign(cr.values); T = len(sig)
    pos = np.empty(T)
    if spec.get("same_day", True):
        pos[:] = np.where(np.isfinite(sig), sig, 0.0)
    else:
        pos[0] = 0.0
        pos[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
    return (pd.Series(pos, index=idx) * f1c_a.diff()).cumsum()


_CARRY_CMP_OPTIONS = {
    "N/A": None,
    "V1: (F1-F2)/F1 - Lag-1":        {"variant": "v1", "signal_type": "f1f2",   "same_day": False},
    "V1: (F1-F2)/F1 - Same-Day":     {"variant": "v1", "signal_type": "f1f2",   "same_day": True},
    "V1: (F1-F3)/F1 - Lag-1":        {"variant": "v1", "signal_type": "f1f3",   "same_day": False},
    "V1: (F1-F3)/F1 - Same-Day":     {"variant": "v1", "signal_type": "f1f3",   "same_day": True},
    "V1: (Cash-3M)/Cash - Lag-1":    {"variant": "v1", "signal_type": "cash3m", "same_day": False},
    "V1: (Cash-3M)/Cash - Same-Day": {"variant": "v1", "signal_type": "cash3m", "same_day": True},
    "V3: F3-F15 - Lag-1":            {"variant": "v3", "j": 3,  "k": 15, "same_day": False},
    "V3: F3-F15 - Same-Day":         {"variant": "v3", "j": 3,  "k": 15, "same_day": True},
    "V3: F4-F16 - Lag-1":            {"variant": "v3", "j": 4,  "k": 16, "same_day": False},
    "V3: F4-F16 - Same-Day":         {"variant": "v3", "j": 4,  "k": 16, "same_day": True},
    "V3: F5-F17 - Lag-1":            {"variant": "v3", "j": 5,  "k": 17, "same_day": False},
    "V3: F5-F17 - Same-Day":         {"variant": "v3", "j": 5,  "k": 17, "same_day": True},
    "V3: F6-F18 - Lag-1":            {"variant": "v3", "j": 6,  "k": 18, "same_day": False},
    "V3: F6-F18 - Same-Day":         {"variant": "v3", "j": 6,  "k": 18, "same_day": True},
    "V3: F7-F19 - Lag-1":            {"variant": "v3", "j": 7,  "k": 19, "same_day": False},
    "V3: F7-F19 - Same-Day":         {"variant": "v3", "j": 7,  "k": 19, "same_day": True},
    "V3: F8-F20 - Lag-1":            {"variant": "v3", "j": 8,  "k": 20, "same_day": False},
    "V3: F8-F20 - Same-Day":         {"variant": "v3", "j": 8,  "k": 20, "same_day": True},
    "V3: F9-F21 - Lag-1":            {"variant": "v3", "j": 9,  "k": 21, "same_day": False},
    "V3: F9-F21 - Same-Day":         {"variant": "v3", "j": 9,  "k": 21, "same_day": True},
    "V3: F10-F22 - Lag-1":           {"variant": "v3", "j": 10, "k": 22, "same_day": False},
    "V3: F10-F22 - Same-Day":        {"variant": "v3", "j": 10, "k": 22, "same_day": True},
    "V3: F11-F23 - Lag-1":           {"variant": "v3", "j": 11, "k": 23, "same_day": False},
    "V3: F11-F23 - Same-Day":        {"variant": "v3", "j": 11, "k": 23, "same_day": True},
    "V3: F12-F24 - Lag-1":           {"variant": "v3", "j": 12, "k": 24, "same_day": False},
    "V3: F12-F24 - Same-Day":        {"variant": "v3", "j": 12, "k": 24, "same_day": True},
    "V4: Z-score (252d) - Lag-1":    {"variant": "v4", "window": 252, "same_day": False},
    "V4: Z-score (252d) - Same-Day": {"variant": "v4", "window": 252, "same_day": True},
}


# ══════════════════════════════════════════════════════
# TAB 7: MOMENTUM SIGNALS
# ══════════════════════════════════════════════════════

with tab7:
    st.markdown("### Momentum Signals — LME Copper")
    st.caption(
        "Baz-Granger CTA trend signal (Eqs 29-33) and MA Crossover. "
        "Signal computed from F1_raw only; PnL from F1_continuous (roll costs captured). "
        "Returns expressed as % of notional. Transaction costs applied on every position change."
    )

    # ── Controls ──────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns([1.6, 1.8, 1.4, 1.4])

    with c1:
        sig_type = st.selectbox(
            "Signal Type",
            ["MA Crossover", "CTA (Baz-Granger)"],
            key="mom_sig_type",
        )

    with c2:
        if sig_type == "MA Crossover":
            variant_opts = {
                "MA(35,43) — Best Sharpe [default]": (35, 43),
                "MA(33,48)": (33, 48),
                "MA(35,44)": (35, 44),
                "MA(34,47)": (34, 47),
                "MA(36,44)": (36, 44),
                "MA(1,5)":   (1, 5),
                "MA(5,20)":  (5, 20),
                "MA(10,60)": (10, 60),
            }
            default_idx = 0
        else:
            variant_opts = {
                "CTA(8,21) — Best Same-Day Sharpe": ("cta_single", 8, 21),
                "CTA(9,21) — Best Lag-1 Sharpe [default]": ("cta_single", 9, 21),
                "CTA(9,20)": ("cta_single", 9, 20),
                "CTA(10,19)": ("cta_single", 10, 19),
                "CTA(14,15)": ("cta_single", 14, 15),
                "CTA Paper (8-16-32 / 24-48-96)": ("cta_paper",),
            }
            default_idx = 1
        variant_label = st.selectbox(
            "Strategy Variant", list(variant_opts.keys()),
            index=default_idx, key="mom_variant",
        )
        variant_params = variant_opts[variant_label]

    with c3:
        timing_label = st.selectbox(
            "Position Entry",
            ["Lag-1 (Next-Day)", "Same-Day"],
            index=0, key="mom_timing",
        )
        same_day = timing_label == "Same-Day"

    with c4:
        tc_bps_map = {
            "0 bps  (No TC)":   0,
            "5 bps  Round Trip":  5,
            "10 bps Round Trip": 10,
            "20 bps Round Trip": 20,
        }
        tc_label = st.selectbox(
            "TC (bps, round-trip)", list(tc_bps_map.keys()),
            index=0, key="mom_tc",
        )
        tc_bps = tc_bps_map[tc_label]   # round-trip cost in basis points

    # ── Custom parameter override ─────────────────────────────────────────────
    with st.expander("Custom Parameters (override dropdown selection)", expanded=False):
        if sig_type == "MA Crossover":
            cp1, cp2, cp3 = st.columns([1, 1, 2])
            with cp1:
                m_cust = st.number_input("Fast window m", min_value=1, max_value=124, value=35, step=1, key="cust_m")
            with cp2:
                n_cust = st.number_input("Slow window n", min_value=m_cust + 1, max_value=126, value=43, step=1, key="cust_n")
            with cp3:
                use_custom = st.checkbox("Use custom MA(m,n)", value=False, key="cust_ma_on")
            if use_custom and n_cust > m_cust:
                variant_params = (m_cust, n_cust)
                variant_label  = f"Custom MA({m_cust},{n_cust})"
        else:
            cp1, cp2, cp3 = st.columns([1, 1, 2])
            with cp1:
                s_cust = st.number_input("Short EWMA S", min_value=1, max_value=49, value=9,  step=1, key="cust_s")
            with cp2:
                l_cust = st.number_input("Long EWMA L",  min_value=s_cust + 1, max_value=100, value=21, step=1, key="cust_l")
            with cp3:
                use_custom = st.checkbox("Use custom CTA(S,L)", value=False, key="cust_cta_on")
            if use_custom and l_cust > s_cust:
                variant_params = ("cta_single", s_cust, l_cust)
                variant_label  = f"Custom CTA({s_cust},{l_cust})"

    # ── Data loading ──────────────────────────────────────────────────────────
    _f1_df = _load_copper_f1_data()
    if _f1_df.empty:
        st.error(
            "LME_Copper_Rolling_F1_v2.csv not found. "
            "Ensure it is in the same directory as app.py."
        )
        st.stop()
    f1r: pd.Series = _f1_df["F1_raw"]
    f1c: pd.Series = _f1_df["F1_continuous"]

    # ── Signal & position computation ─────────────────────────────────────────
    def _ewma(s: pd.Series, n: int) -> pd.Series:
        return s.ewm(com=n - 1, adjust=False).mean()

    if sig_type == "MA Crossover":
        m_val, n_val = variant_params
        sig_raw = np.sign(f1r.rolling(m_val).mean() - f1r.rolling(n_val).mean()).values.astype(float)

    elif isinstance(variant_params, tuple) and variant_params[0] == "cta_paper":
        pv = f1r.rolling(63).std()
        us = []
        for sk, lk in zip((8, 16, 32), (24, 48, 96)):
            x = _ewma(f1r, sk) - _ewma(f1r, lk)
            y = x / pv
            with np.errstate(invalid="ignore"):
                z = (y / y.rolling(252).std()).values
                u = z * np.exp(-z ** 2 / 4) / 0.89
            us.append(u)
        sig_raw = np.sign(np.nanmean(np.stack(us, axis=1), axis=1))

    else:  # cta_single
        _, s_val, l_val = variant_params
        x = _ewma(f1r, s_val) - _ewma(f1r, l_val)
        y = x / f1r.rolling(63).std()
        with np.errstate(invalid="ignore"):
            z = (y / y.rolling(252).std()).values
            u = z * np.exp(-z ** 2 / 4) / 0.89
        sig_raw = np.sign(u)

    T = len(sig_raw)
    pos_np = np.empty(T)
    if same_day:
        pos_np[:] = np.where(np.isfinite(sig_raw), sig_raw, 0.0)
    else:
        pos_np[0] = 0.0
        pos_np[1:] = np.where(np.isfinite(sig_raw[:-1]), sig_raw[:-1], 0.0)
    sig_np = sig_raw

    # ── PnL with transaction costs ─────────────────────────────────────────────
    delta_np  = f1c.diff().values.astype(float)
    pos_s     = pd.Series(pos_np, index=f1r.index)
    delta_s   = pd.Series(delta_np, index=f1r.index)
    gross_pnl = pos_s * delta_s

    # TC in bps: cost per position change = |Δpos| × (bps/10000 / 2) × F1_price
    pos_change   = pos_s.diff().abs()
    pos_change.iloc[0] = abs(pos_s.iloc[0])
    tc_cost_s    = pos_change * (tc_bps / 10000.0 / 2.0) * f1c
    net_pnl      = gross_pnl - tc_cost_s

    cum_pnl_gross = gross_pnl.cumsum()
    cum_pnl_net   = net_pnl.cumsum()

    # ── Pre-compute daily returns (full period, needed for rolling Sharpe) ─────
    f1_prev_full  = f1c.shift(1)
    gross_ret_all = (gross_pnl / f1_prev_full).replace([np.inf, -np.inf], np.nan)
    net_ret_all   = (net_pnl   / f1_prev_full).replace([np.inf, -np.inf], np.nan)

    # ── Date filter for performance metrics ───────────────────────────────────
    st.divider()
    pf_c1, pf_c2 = st.columns([3, 1])
    with pf_c1:
        perf_dates = st.date_input(
            "Performance period  (signal uses full history — only metrics & charts below update)",
            value=(f1r.index[0].date(), f1r.index[-1].date()),
            min_value=f1r.index[0].date(), max_value=f1r.index[-1].date(),
            key="mom_perf_dates",
        )
    p_start = pd.Timestamp(perf_dates[0]) if len(perf_dates) >= 1 else f1r.index[0]
    p_end   = pd.Timestamp(perf_dates[1]) if len(perf_dates) == 2 else f1r.index[-1]
    pmask   = (gross_pnl.index >= p_start) & (gross_pnl.index <= p_end)

    gross_pnl_f = gross_pnl[pmask];   net_pnl_f = net_pnl[pmask]
    pos_s_f     = pos_s[pmask]
    gross_ret_f = gross_ret_all[pmask]; net_ret_f = net_ret_all[pmask]

    # ── Performance metrics ────────────────────────────────────────────────────
    def _perf(daily_pnl: pd.Series, daily_ret: pd.Series,
              position: pd.Series, label: str) -> dict:
        active  = daily_ret[position != 0].dropna()
        pnl_act = daily_pnl[position != 0].dropna()
        n = len(active)
        if n < 20:
            return {}
        ann_r  = float(active.mean() * 252 * 100)
        ann_sd = float(active.std()  * np.sqrt(252) * 100)
        sharpe = ann_r / ann_sd if ann_sd > 0 else np.nan
        down   = active[active < 0]
        srt_d  = float(down.std() * np.sqrt(252) * 100) if len(down) > 1 else np.nan
        sortino = ann_r / srt_d if srt_d and srt_d > 0 else np.nan
        cum_r   = daily_ret.fillna(0).cumsum() * 100
        mdd_pct = float((cum_r - cum_r.cummax()).min())
        calmar  = ann_r / abs(mdd_pct) if mdd_pct != 0 else np.nan
        wins, losses = pnl_act[pnl_act > 0], pnl_act[pnl_act < 0]
        return {
            "label": label, "n": n,
            "sharpe": sharpe, "sortino": sortino,
            "ann_ret_pct": ann_r, "ann_std_pct": ann_sd,
            "mdd_pct": mdd_pct, "calmar": calmar,
            "hit_rate": float((active > 0).mean()) * 100,
            "profit_factor": abs(wins.sum() / losses.sum()) if len(losses) > 0 else np.nan,
            "total_pnl_usdmt": float(pnl_act.sum()),
        }

    m_gross = _perf(gross_pnl_f, gross_ret_f, pos_s_f, "Gross (No TC)")
    m_net   = _perf(net_pnl_f,   net_ret_f,   pos_s_f, f"Net ({tc_label})")

    # ── Metric cards ──────────────────────────────────────────────────────────
    def _mcard(col, label, val, fmt=".2f", suffix="", good_high=True):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            col.markdown(f'<div class="metric-compact"><h4>{label}</h4><p class="value">—</p></div>',
                         unsafe_allow_html=True)
            return
        v_str = f"{val:{fmt}}{suffix}"
        col.markdown(f'<div class="metric-compact"><h4>{label}</h4><p class="value">{v_str}</p></div>',
                     unsafe_allow_html=True)

    section_header = lambda t: st.markdown(f'<div class="section-header">{t}</div>', unsafe_allow_html=True)

    st.divider()
    section_header("PERFORMANCE METRICS")
    st.caption(f"**Strategy:** {variant_label}  |  **Entry:** {timing_label}  |  **TC:** {tc_label}")

    if m_gross and m_net:
        # Row 1: Sharpe, Sortino, Ann Return, Std Dev, Active Days
        cols = st.columns(8)
        _mcard(cols[0], "Sharpe Gross",      m_gross.get("sharpe"),       ".2f")
        _mcard(cols[1], "Sharpe Net",        m_net.get("sharpe"),         ".2f")
        _mcard(cols[2], "Sortino Gross",     m_gross.get("sortino"),      ".2f")
        _mcard(cols[3], "Sortino Net",       m_net.get("sortino"),        ".2f")
        _mcard(cols[4], "Ann Ret% Gross",    m_gross.get("ann_ret_pct"),  ".2f", "%")
        _mcard(cols[5], "Ann Ret% Net",      m_net.get("ann_ret_pct"),    ".2f", "%")
        _mcard(cols[6], "Ann Std Dev%",      m_gross.get("ann_std_pct"),  ".2f", "%")
        _mcard(cols[7], "Active Days",       float(m_gross.get("n", 0)), ",.0f")

        # Row 2: MaxDD, Calmar, Hit Rate, Profit Factor, Total PnL
        cols2 = st.columns(8)
        _mcard(cols2[0], "Max DD% Gross",    m_gross.get("mdd_pct"),      ".2f", "%", good_high=False)
        _mcard(cols2[1], "Max DD% Net",      m_net.get("mdd_pct"),        ".2f", "%", good_high=False)
        _mcard(cols2[2], "Calmar Gross",     m_gross.get("calmar"),       ".2f")
        _mcard(cols2[3], "Calmar Net",       m_net.get("calmar"),         ".2f")
        _mcard(cols2[4], "Hit Rate",         m_gross.get("hit_rate"),     ".2f", "%")
        _mcard(cols2[5], "Profit Factor",    m_gross.get("profit_factor"),".2f")
        _mcard(cols2[6], "PnL Gross $/MT",   m_gross.get("total_pnl_usdmt"), ",.2f")
        _mcard(cols2[7], "PnL Net $/MT",     m_net.get("total_pnl_usdmt"),   ",.2f")
    else:
        st.warning("Insufficient active trading days to compute metrics.")

    # ── Rolling Sharpe ─────────────────────────────────────────────────────────
    st.divider()
    section_header("ROLLING SHARPE RATIO")
    rs_c1, rs_c2 = st.columns([3, 1])
    with rs_c2:
        rs_window = st.radio("Window", ["1 Year (252d)", "2 Years (504d)", "Both"],
                             index=2, key="rs_window", horizontal=False)

    _dr = gross_ret_all.fillna(0)
    roll_252 = (_dr.rolling(252).mean() / _dr.rolling(252).std() * np.sqrt(252))
    roll_504 = (_dr.rolling(504).mean() / _dr.rolling(504).std() * np.sqrt(252))

    with rs_c1:
        fig_rs = go.Figure()
        if rs_window in ("1 Year (252d)", "Both"):
            fig_rs.add_trace(go.Scatter(
                x=roll_252.index, y=roll_252.values, name="Rolling Sharpe (1yr)",
                mode="lines", line=dict(color=COLORS["primary"], width=1.6),
                hovertemplate="%{x|%b %Y}<br>Sharpe (1yr): %{y:.2f}<extra></extra>",
            ))
        if rs_window in ("2 Years (504d)", "Both"):
            fig_rs.add_trace(go.Scatter(
                x=roll_504.index, y=roll_504.values, name="Rolling Sharpe (2yr)",
                mode="lines", line=dict(color=COLORS["amber"], width=1.6, dash="dot"),
                hovertemplate="%{x|%b %Y}<br>Sharpe (2yr): %{y:.2f}<extra></extra>",
            ))
        # Shade positive regions
        fig_rs.add_hline(y=0,  line_dash="dash", line_color="#475569", line_width=1)
        fig_rs.add_hline(y=0.5, line_dash="dot", line_color=COLORS["green"],
                         line_width=0.8, annotation_text="0.5", annotation_position="right")
        fig_rs.update_layout(
            **CHART_LAYOUT, height=320,
            title=dict(text=f"{variant_label} — Rolling Sharpe ({timing_label})", font=dict(size=13)),
            yaxis_title="Annualised Sharpe", xaxis_title=None, hovermode="x unified",
        )
        fig_rs.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_rs, use_container_width=True)
    st.caption("Signal computed over full 2006-2025 history. Rolling Sharpe uses gross returns. "
               "Positive/negative swings show regime dependence — a consistently positive curve indicates robustness.")

    # ── Cumulative PnL chart ──────────────────────────────────────────────────
    st.divider()
    section_header("CUMULATIVE PnL (USD/MT)")

    # Comparison strategy pickers
    cmp_keys = list(_MOM_CMP_OPTIONS.keys())
    cc1, cc2 = st.columns(2)
    with cc1:
        cmp_a_label = st.selectbox(
            "Strategy A", cmp_keys, index=0, key="cmp_a",
            help="First strategy to plot. Select N/A to hide.",
        )
    with cc2:
        cmp_b_label = st.selectbox(
            "Strategy B", cmp_keys, index=0, key="cmp_b",
            help="Second strategy to compare. Select N/A to hide.",
        )

    _CMP_COLORS = [COLORS["primary"], COLORS["amber"], COLORS["green"], "#A78BFA"]
    fig_cum = go.Figure()
    _cmp_plotted = 0
    for _lbl, _cidx in [(cmp_a_label, 0), (cmp_b_label, 1)]:
        _spec = _MOM_CMP_OPTIONS.get(_lbl)
        if _spec is None:
            continue
        _cpnl = _mom_cum_pnl(f1r, f1c, _spec)
        fig_cum.add_trace(go.Scatter(
            x=_cpnl.index, y=_cpnl.values,
            name=_lbl, mode="lines",
            line=dict(color=_CMP_COLORS[_cidx], width=1.8,
                      dash="dot" if _cidx == 1 else "solid"),
            hovertemplate=f"%{{x|%b %d, %Y}}<br>{_lbl}: $%{{y:,.1f}}/MT<extra></extra>",
        ))
        _cmp_plotted += 1

    # Also show gross/net for the currently selected main strategy
    fig_cum.add_trace(go.Scatter(
        x=cum_pnl_gross.index, y=cum_pnl_gross.values,
        name=f"{variant_label} Gross", mode="lines",
        line=dict(color="#64748B", width=1.2, dash="solid"),
        hovertemplate="%{x|%b %d, %Y}<br>Gross: $%{y:,.1f}/MT<extra></extra>",
        visible="legendonly",
    ))
    if tc_bps > 0:
        fig_cum.add_trace(go.Scatter(
            x=cum_pnl_net.index, y=cum_pnl_net.values,
            name=f"{variant_label} Net ({tc_label})", mode="lines",
            line=dict(color="#94A3B8", width=1.2, dash="dot"),
            hovertemplate="%{x|%b %d, %Y}<br>Net: $%{y:,.1f}/MT<extra></extra>",
            visible="legendonly",
        ))

    fig_cum.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
    fig_cum.update_layout(
        **CHART_LAYOUT, height=420,
        title=dict(text="Strategy Comparison — Cumulative PnL (Gross, USD/MT)", font=dict(size=13)),
        yaxis_title="Cumulative PnL (USD/MT)",
        xaxis_title=None, hovermode="x unified",
    )
    fig_cum.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_cum, use_container_width=True)

    # ── Annual PnL bar chart ──────────────────────────────────────────────────
    st.divider()
    section_header("ANNUAL PnL BREAKDOWN (Gross, USD/MT)")

    annual_pnl = gross_pnl.resample("YE").sum()
    annual_pnl.index = annual_pnl.index.year
    bar_colors = [COLORS["green"] if v >= 0 else COLORS["red"] for v in annual_pnl.values]

    fig_ann = go.Figure(go.Bar(
        x=annual_pnl.index.astype(str), y=annual_pnl.values,
        marker_color=bar_colors, name="Annual PnL",
        hovertemplate="%{x}<br>PnL: $%{y:,.1f}/MT<extra></extra>",
    ))
    fig_ann.update_layout(
        **CHART_LAYOUT, height=300,
        title=dict(text="Annual PnL", font=dict(size=13)),
        yaxis_title="PnL (USD/MT)", xaxis_title=None, showlegend=False,
    )
    st.plotly_chart(fig_ann, use_container_width=True)

    # ── Parameter Optimization Heatmap ────────────────────────────────────────
    st.divider()
    section_header("PARAMETER OPTIMIZATION SURFACE")

    if sig_type == "MA Crossover":
        _df_opt = _load_ma_opt()
        if _df_opt.empty:
            st.info("MA_Crossover_Optimization.csv not found in data/ folder.")
        else:
            hm_c1, hm_c2 = st.columns([3, 1])
            with hm_c2:
                hm_metric = st.selectbox("Colour by", ["sharpe", "ann_return", "hit_rate"],
                                         format_func=lambda x: {"sharpe": "Sharpe Ratio",
                                                                 "ann_return": "Ann Return ($/MT)",
                                                                 "hit_rate": "Hit Rate %"}[x],
                                         key="hm_metric")
                hm_n_max = st.slider("Max n (slow window)", 20, 126, 80, step=5, key="hm_nmax")
            _df_filt = _df_opt[_df_opt["n"] <= hm_n_max].copy()
            _pivot = _df_filt.pivot_table(index="m", columns="n", values=hm_metric)
            import plotly.express as _px
            fig_hm = _px.imshow(
                _pivot,
                color_continuous_scale="RdYlGn",
                zmin=_df_filt[hm_metric].quantile(0.05),
                zmax=_df_filt[hm_metric].quantile(0.95),
                labels=dict(x="Slow window n", y="Fast window m", color=hm_metric),
                aspect="auto",
            )
            # Mark top-5 pairs
            _top5 = _df_opt.nlargest(5, "sharpe")
            fig_hm.add_trace(go.Scatter(
                x=_top5["n"], y=_top5["m"],
                mode="markers+text",
                marker=dict(symbol="star", size=10, color="#FFFFFF", line=dict(color="#B87333", width=1.5)),
                text=[f"#{i+1}" for i in range(len(_top5))],
                textposition="top center", textfont=dict(size=8, color="#FFFFFF"),
                name="Top 5", showlegend=True,
                hovertemplate="MA(%{y},%{x})<br>Sharpe: %{customdata:.3f}<extra>Top 5</extra>",
                customdata=_top5["sharpe"].values,
            ))
            fig_hm.update_layout(
                **CHART_LAYOUT, height=480,
                title=dict(text=f"MA Crossover — Sharpe surface  (m=fast, n=slow, n≤{hm_n_max})", font=dict(size=13)),
                coloraxis_colorbar=dict(title=hm_metric, thickness=14),
            )
            with hm_c1:
                st.plotly_chart(fig_hm, use_container_width=True)
            st.caption("White stars = current top-5 by Sharpe (Lag-1). A wide green plateau means the "
                       "strategy is robust to parameter choice. An isolated peak suggests overfitting.")

    else:  # CTA
        _df_cta = _load_cta_opt()
        if _df_cta.empty:
            st.info("CTA_Optimization.csv not found in data/ folder.")
        else:
            hm_c1, hm_c2 = st.columns([3, 1])
            with hm_c2:
                cta_metric = st.selectbox("Colour by", ["sharpe", "ann_return", "hit_rate"],
                                          format_func=lambda x: {"sharpe": "Sharpe Ratio",
                                                                  "ann_return": "Ann Return ($/MT)",
                                                                  "hit_rate": "Hit Rate %"}[x],
                                          key="cta_hm_metric")
            import plotly.express as _px
            fig_cta = _px.scatter(
                _df_cta, x="L", y="S",
                color=cta_metric,
                color_continuous_scale="RdYlGn",
                range_color=[_df_cta[cta_metric].quantile(0.05), _df_cta[cta_metric].quantile(0.95)],
                size_max=8,
                labels=dict(S="Short EWMA (S)", L="Long EWMA (L)", color=cta_metric),
                hover_data={"S": True, "L": True, cta_metric: ":.3f"},
            )
            _top5c = _df_cta.nlargest(5, "sharpe")
            fig_cta.add_trace(go.Scatter(
                x=_top5c["L"], y=_top5c["S"],
                mode="markers+text",
                marker=dict(symbol="star", size=12, color="#FFFFFF", line=dict(color="#B87333", width=1.5)),
                text=[f"#{i+1}" for i in range(len(_top5c))],
                textposition="top center", textfont=dict(size=8, color="#FFFFFF"),
                name="Top 5",
                hovertemplate="CTA(S=%{y},L=%{x})<extra>Top 5</extra>",
            ))
            fig_cta.update_layout(
                **CHART_LAYOUT, height=460,
                title=dict(text="CTA (Baz-Granger) — Sharpe scatter  (S=short EWMA, L=long EWMA)", font=dict(size=13)),
            )
            with hm_c1:
                st.plotly_chart(fig_cta, use_container_width=True)
            st.caption("White stars = top-5 by Sharpe (Lag-1). Each dot is one (S,L) pair.")

    # ── Signal & Position chart ────────────────────────────────────────────────
    st.divider()
    section_header("SIGNAL & POSITION OVER TIME")

    # Full date range — no filter widget
    f1r_w = f1r
    pos_w = pos_s

    fig_sig = make_subplots(
        rows=2, cols=1, shared_xaxes=True, row_heights=[0.65, 0.35],
        vertical_spacing=0.04,
    )
    # F1_raw price
    fig_sig.add_trace(go.Scatter(
        x=f1r_w.index, y=f1r_w.values, name="F1 Price",
        line=dict(color=COLORS["primary"], width=1.5),
        hovertemplate="%{x|%b %d, %Y}<br>F1: $%{y:,.1f}<extra></extra>",
    ), row=1, col=1)

    # Separate bar traces for Long / Short so legend squares are clearly colored
    pos_long  = pos_w.where(pos_w > 0, 0.0)
    pos_short = pos_w.where(pos_w < 0, 0.0)

    fig_sig.add_trace(go.Bar(
        x=pos_w.index, y=pos_long.values,
        name="Long (+1)", marker_color="#00E676", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Long<extra></extra>",
    ), row=2, col=1)
    fig_sig.add_trace(go.Bar(
        x=pos_w.index, y=pos_short.values,
        name="Short (-1)", marker_color="#FF1744", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Short<extra></extra>",
    ), row=2, col=1)

    fig_sig.update_layout(
        **CHART_LAYOUT, height=500, barmode="overlay",
        title=dict(text=f"{variant_label} — Price & Position ({timing_label})", font=dict(size=13)),
        hovermode="x unified", showlegend=True,
        xaxis2_title=None,
    )
    fig_sig.update_yaxes(title_text="F1 Price ($/MT)", row=1, col=1)
    fig_sig.update_yaxes(title_text="Position", tickvals=[-1, 0, 1],
                          ticktext=["Short", "Flat", "Long"], row=2, col=1)
    fig_sig.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_sig, use_container_width=True)

    # ── Signal flip table (most recent 20) ────────────────────────────────────
    st.divider()
    section_header("RECENT SIGNAL FLIPS (Last 20)")

    pos_full   = pos_s.copy()
    flips_mask = pos_full.diff().abs() > 0
    flips_mask.iloc[0] = pos_full.iloc[0] != 0
    flip_dates = pos_full[flips_mask].tail(20)

    if not flip_dates.empty:
        flip_df = pd.DataFrame({
            "Date":       flip_dates.index.strftime("%Y-%m-%d"),
            "Position":   flip_dates.values.astype(int),
            "Direction":  ["LONG" if v > 0 else "SHORT" for v in flip_dates.values],
            "F1_raw":     f1r.reindex(flip_dates.index).round(1).values,
            "Gross PnL on flip day": gross_pnl.reindex(flip_dates.index).round(2).values,
            "TC cost":    tc_cost_s.reindex(flip_dates.index).round(2).values,
            "Net PnL":    net_pnl.reindex(flip_dates.index).round(2).values,
        })
        st.dataframe(flip_df, use_container_width=True, hide_index=True)
    else:
        st.info("No signal flips found in data.")

    # ── Methodology note ──────────────────────────────────────────────────────
    st.divider()
    with st.expander("Methodology Notes", expanded=False):
        st.markdown("""
**Signal sources (no leakage)**
- All signals computed exclusively from `F1_raw` (closing prices, backward-looking only)
- PnL computed exclusively from `F1_continuous` (return-based roll-adjusted series)

**MA Crossover**
- Signal = sign(SMA(m) − SMA(n)) where m < n
- +1 (long) when fast MA is above slow MA; −1 (short) otherwise

**CTA Signal — Baz-Granger Eqs 29-33**
- x = EWMA(S) − EWMA(L)  [EWMA convention: com = n−1, i.e. λ = (n−1)/n]
- y = x / σ₆₃(price)     [63-day price volatility normalisation]
- z = y / σ₂₅₂(y)        [252-day signal normalisation]
- u = z · exp(−z²/4) / 0.89  [response function — shrinks extreme signals]
- Signal = sign(u)
- CTA Paper uses 3 timescales (S,L) = (8,24), (16,48), (32,96); S_CTA = mean(u₁,u₂,u₃)

**Position timing**
- *Lag-1*: position[t] = signal[t−1]  → entered at close t−1, PnL from t−1→t
- *Same-Day*: position[t] = signal[t]  → entered at close t, PnL from t−1→t
- MA Crossover: Lag-1 outperforms (4/5 top pairs, avg Sharpe delta −0.04)
- CTA: Same-Day substantially outperforms (5/5 top pairs, avg Sharpe delta +0.30)

**Transaction costs**
- Expressed in basis points (bps) of notional, round-trip
- TC_cost[t] = |Δposition[t]| × (bps / 10000 / 2) × F1_cont[t]
- Flip (+1→−1): |Δ|=2 → cost = 1 full round trip × price
- Entry (0→±1): |Δ|=1 → cost = ½ round trip × price
- Cost is time-varying (scales with copper price level)

**Returns & risk metrics**
- daily_ret[t] = position[t] × ΔF1_cont[t] / F1_cont[t−1]
- All % metrics (Ann Return, Std Dev, Max DD, Calmar, Sortino) computed from daily_ret
- Sharpe = Ann_ret / Ann_std (unitless, consistent across gross/net)
        """)


# ══════════════════════════════════════════════════════
# TAB 8: CARRY SIGNALS
# ══════════════════════════════════════════════════════

with tab8:
    st.markdown("### Carry Signals — LME Copper")
    st.caption("Term structure carry: Long in backwardation, Short in contango. Signal from curve shape; PnL always from F1_continuous.")

    # ── Controls ──────────────────────────────────────────────────────────────
    c8_c1, c8_c2, c8_c3, c8_c4 = st.columns([2, 2, 1.5, 1.5])
    with c8_c1:
        carry_vgroup = st.selectbox(
            "Variant Group",
            ["V1 — Roll Yield", "V2 — Annualized", "V3 — Long Slope", "V4 — Z-score"],
            key="carry_vgroup",
        )
    with c8_c2:
        if carry_vgroup == "V1 — Roll Yield":
            _c8_sub_opts = {"(F1-F2)/F1": "f1f2", "(F1-F3)/F1": "f1f3", "(Cash-3M)/Cash": "cash3m"}
        elif carry_vgroup == "V2 — Annualized":
            _c8_sub_opts = {"(F1-F2)/F1 x12 (ann.)": "f1f2_ann", "(F1-F3)/F1 x6 (ann.)": "f1f3_ann"}
        elif carry_vgroup == "V3 — Long Slope":
            _c8_sub_opts = {f"F{j}-F{k} Slope": (j, k)
                            for j, k in [(3,15),(4,16),(5,17),(6,18),(7,19),(8,20),(9,21),(10,22),(11,23),(12,24)]}
        else:
            _c8_sub_opts = {"Z-score (252d window)": 252}
        carry_sub_label = st.selectbox("Sub-Variant", list(_c8_sub_opts.keys()), key="carry_subv")
        carry_sub_val = _c8_sub_opts[carry_sub_label]
    with c8_c3:
        carry_timing = st.selectbox("Position Entry", ["Same-Day", "Lag-1 (Next-Day)"], index=0, key="carry_timing")
        carry_same_day = carry_timing == "Same-Day"
    with c8_c4:
        carry_tc_map = {"0 bps (No TC)": 0, "5 bps": 5, "10 bps": 10, "20 bps": 20}
        carry_tc_label = st.selectbox("TC (bps, round-trip)", list(carry_tc_map.keys()), index=0, key="carry_tc")
        carry_tc_bps = carry_tc_map[carry_tc_label]

    # Build spec dict
    if carry_vgroup == "V1 — Roll Yield":
        carry_spec = {"variant": "v1", "signal_type": carry_sub_val, "same_day": carry_same_day}
    elif carry_vgroup == "V2 — Annualized":
        carry_spec = {"variant": "v2", "signal_type": carry_sub_val, "same_day": carry_same_day}
    elif carry_vgroup == "V3 — Long Slope":
        carry_spec = {"variant": "v3", "j": carry_sub_val[0], "k": carry_sub_val[1], "same_day": carry_same_day}
    else:
        carry_spec = {"variant": "v4", "window": carry_sub_val, "same_day": carry_same_day}

    # ── Data loading ──────────────────────────────────────────────────────────
    _f1_df_c8 = _load_copper_f1_data()
    if _f1_df_c8.empty:
        st.error("LME_Copper_Rolling_F1_v2.csv not found. Place it in the same directory as app.py.")
        st.stop()
    cf1c = _f1_df_c8["F1_continuous"]
    cf1r = _f1_df_c8["F1_raw"]

    _cu_sheet_c8 = _find_curve_sheet("Copper", curve_data) if curve_data else None
    if not curve_data or _cu_sheet_c8 is None:
        st.error("Futures Curve data not loaded. Upload Metals Futures Curve file in the sidebar.")
        st.stop()
    c8_curve_px = curve_data[_cu_sheet_c8]["prices"].copy()
    c8_curve_px.index = pd.to_datetime(c8_curve_px.index).normalize()
    c8_curve_px = c8_curve_px.sort_index()

    _cash_cu8 = parse_cash_3m_columns(cash_data.get("Copper", pd.DataFrame()), "Copper")
    if not _cash_cu8.empty:
        _cash_cu8.index = pd.to_datetime(_cash_cu8.index).normalize()

    # ── Carry signal computation ───────────────────────────────────────────────
    carry_raw = _carry_raw_signal(c8_curve_px, _cash_cu8, carry_spec)
    if carry_raw.empty:
        st.error("Could not compute carry signal. Check that the required curve columns (F1, F2, F3, etc.) exist in the uploaded futures curve file.")
        st.stop()

    _c8_idx = carry_raw.index.intersection(cf1c.index)
    carry_raw = carry_raw.reindex(_c8_idx).dropna()
    _c8_idx = carry_raw.index
    cf1c_a = cf1c.reindex(_c8_idx)

    carry_sig_arr = np.sign(carry_raw.values)
    T_c8 = len(carry_sig_arr)
    carry_pos_np = np.empty(T_c8)
    if carry_same_day:
        carry_pos_np[:] = np.where(np.isfinite(carry_sig_arr), carry_sig_arr, 0.0)
    else:
        carry_pos_np[0] = 0.0
        carry_pos_np[1:] = np.where(np.isfinite(carry_sig_arr[:-1]), carry_sig_arr[:-1], 0.0)
    carry_pos = pd.Series(carry_pos_np, index=_c8_idx)

    c8_delta = cf1c_a.diff()
    c8_gross_pnl = carry_pos * c8_delta
    c8_pos_change = carry_pos.diff().abs()
    c8_pos_change.iloc[0] = abs(carry_pos.iloc[0])
    c8_tc_cost = c8_pos_change * (carry_tc_bps / 10000.0 / 2.0) * cf1c_a
    c8_net_pnl = c8_gross_pnl - c8_tc_cost
    c8_cum_gross = c8_gross_pnl.cumsum()
    c8_cum_net = c8_net_pnl.cumsum()

    cf1_prev8 = cf1c_a.shift(1)
    c8_gross_ret_all = (c8_gross_pnl / cf1_prev8).replace([np.inf, -np.inf], np.nan)
    c8_net_ret_all = (c8_net_pnl / cf1_prev8).replace([np.inf, -np.inf], np.nan)

    # Regime signal series (used by multiple sections below)
    _c8_sig_bin = pd.Series(np.sign(carry_raw.values), index=_c8_idx)
    _c8_flip_mask = _c8_sig_bin.diff().abs() > 0
    _c8_flip_mask.iloc[0] = True
    _c8_regime_id = _c8_flip_mask.cumsum()
    _c8_back_mask = (_c8_sig_bin > 0)
    _c8_cont_mask = (_c8_sig_bin < 0)

    # Regime durations
    _c8_regime_vals = _c8_sig_bin.groupby(_c8_regime_id).first()
    _c8_regime_lens = _c8_sig_bin.groupby(_c8_regime_id).count()
    _back_durs = _c8_regime_lens[_c8_regime_vals > 0].tolist()
    _cont_durs = _c8_regime_lens[_c8_regime_vals < 0].tolist()
    avg_back_dur = int(np.mean(_back_durs)) if _back_durs else 0
    avg_cont_dur = int(np.mean(_cont_durs)) if _cont_durs else 0

    section_header = lambda t: st.markdown(f'<div class="section-header">{t}</div>', unsafe_allow_html=True)
    def _cmcard(col, label, val, fmt=".2f", suffix=""):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            col.markdown(f'<div class="metric-compact"><h4>{label}</h4><p class="value">—</p></div>', unsafe_allow_html=True)
            return
        col.markdown(f'<div class="metric-compact"><h4>{label}</h4><p class="value">{val:{fmt}}{suffix}</p></div>', unsafe_allow_html=True)

    # ── Section 2: Live Regime Badge ──────────────────────────────────────────
    last_date_c8 = carry_raw.index[-1]
    last_carry_val = float(carry_raw.iloc[-1])
    is_back_c8 = last_carry_val > 0
    _rcolor = "#5BAD72" if is_back_c8 else "#B85450"
    _rword  = "BACKWARDATION" if is_back_c8 else "CONTANGO"
    _rsig   = "LONG  (+1)" if is_back_c8 else "SHORT (-1)"
    _rbg    = "rgba(91,173,114,0.08)" if is_back_c8 else "rgba(184,84,80,0.08)"
    carry_pct_last = last_carry_val * 100

    _flip_pos_series = _c8_sig_bin[_c8_flip_mask]
    if len(_flip_pos_series) > 1:
        last_flip_dt = _flip_pos_series.index[-2]
        days_since = (last_date_c8 - last_flip_dt).days
        flip_str = f"{last_flip_dt.strftime('%Y-%m-%d')}  ({days_since}d ago)"
    else:
        flip_str = "N/A"

    st.divider()
    st.markdown(f"""
    <div style="background:{_rbg}; border:1px solid {_rcolor}; border-left:5px solid {_rcolor};
                border-radius:4px; padding:18px 24px; margin-bottom:4px;">
      <div style="display:flex; gap:48px; align-items:center; flex-wrap:wrap;">
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Current Regime</div>
          <div style="color:{_rcolor};font-size:1.6rem;font-weight:700;letter-spacing:0.04em;font-family:'IBM Plex Mono',monospace;">{_rword}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">as of {last_date_c8.strftime('%Y-%m-%d')}</div>
        </div>
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Carry Value</div>
          <div style="color:{_rcolor};font-size:1.6rem;font-weight:700;font-family:'IBM Plex Mono',monospace;">{carry_pct_last:+.3f}%</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">{carry_sub_label}</div>
        </div>
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Active Signal</div>
          <div style="color:{_rcolor};font-size:1.6rem;font-weight:700;font-family:'IBM Plex Mono',monospace;">{_rsig}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">{carry_timing}</div>
        </div>
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Last Regime Flip</div>
          <div style="color:#D4CFC8;font-size:1.1rem;font-weight:600;font-family:'IBM Plex Mono',monospace;">{flip_str}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">Avg duration: {avg_back_dur}d back / {avg_cont_dur}d cont</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Section 3: Carry Signal History ───────────────────────────────────────
    st.divider()
    section_header("CARRY SIGNAL HISTORY")

    carry_pct_series = carry_raw * 100
    fig_csh = make_subplots(
        rows=2, cols=1, shared_xaxes=True, row_heights=[0.65, 0.35],
        vertical_spacing=0.04,
    )
    fig_csh.add_trace(go.Scatter(
        x=carry_pct_series.index, y=carry_pct_series.values,
        name="Carry Value (%)", mode="lines",
        line=dict(color=COLORS["primary"], width=1.5),
        hovertemplate="%{x|%b %d, %Y}<br>Carry: %{y:.4f}%<extra></extra>",
    ), row=1, col=1)
    fig_csh.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1.2, row=1, col=1)

    # Shaded regions for backwardation / contango
    _back_arr = carry_pct_series.where(_c8_back_mask, np.nan)
    _cont_arr = carry_pct_series.where(_c8_cont_mask, np.nan)
    fig_csh.add_trace(go.Scatter(
        x=_back_arr.index, y=_back_arr.values,
        name="Backwardation", mode="lines",
        line=dict(color="#5BAD72", width=0), fill="tozeroy",
        fillcolor="rgba(91,173,114,0.18)",
        hovertemplate="%{x|%b %d, %Y}<br>Back.: %{y:.4f}%<extra></extra>",
    ), row=1, col=1)
    fig_csh.add_trace(go.Scatter(
        x=_cont_arr.index, y=_cont_arr.values,
        name="Contango", mode="lines",
        line=dict(color="#B85450", width=0), fill="tozeroy",
        fillcolor="rgba(184,84,80,0.18)",
        hovertemplate="%{x|%b %d, %Y}<br>Cont.: %{y:.4f}%<extra></extra>",
    ), row=1, col=1)

    _sig_long_c8  = _c8_sig_bin.where(_c8_sig_bin > 0, 0.0)
    _sig_short_c8 = _c8_sig_bin.where(_c8_sig_bin < 0, 0.0)
    fig_csh.add_trace(go.Bar(
        x=_c8_idx, y=_sig_long_c8.values,
        name="Long (+1)", marker_color="#00E676", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Long<extra></extra>",
    ), row=2, col=1)
    fig_csh.add_trace(go.Bar(
        x=_c8_idx, y=_sig_short_c8.values,
        name="Short (-1)", marker_color="#FF1744", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Short<extra></extra>",
    ), row=2, col=1)

    fig_csh.update_layout(
        **CHART_LAYOUT, height=500, barmode="overlay",
        title=dict(text=f"{carry_sub_label} — Raw Carry Value & Binary Signal", font=dict(size=13)),
        hovermode="x unified", showlegend=True,
    )
    fig_csh.update_yaxes(title_text="Carry (%)", row=1, col=1)
    fig_csh.update_yaxes(title_text="Signal", tickvals=[-1, 0, 1],
                          ticktext=["Short", "Flat", "Long"], row=2, col=1)
    fig_csh.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_csh, use_container_width=True)
    st.caption("Top: raw carry value (continuous magnitude). Green fill = backwardation (long), Red fill = contango (short). "
               "Bottom: resulting binary signal. Unlike momentum, carry has an economic magnitude — deep backwardation periods are visually distinct.")

    # ── Section 4: Date Filter + Performance Metrics ──────────────────────────
    st.divider()
    pf8_c1, _ = st.columns([3, 1])
    with pf8_c1:
        carry_perf_dates = st.date_input(
            "Performance period  (signal uses full history — only metrics & charts below update)",
            value=(_c8_idx[0].date(), _c8_idx[-1].date()),
            min_value=_c8_idx[0].date(), max_value=_c8_idx[-1].date(),
            key="carry_perf_dates",
        )
    cp8_start = pd.Timestamp(carry_perf_dates[0]) if len(carry_perf_dates) >= 1 else _c8_idx[0]
    cp8_end   = pd.Timestamp(carry_perf_dates[1]) if len(carry_perf_dates) == 2 else _c8_idx[-1]
    cp8_mask  = (_c8_idx >= cp8_start) & (_c8_idx <= cp8_end)

    c8_gross_pnl_f = c8_gross_pnl[cp8_mask];  c8_net_pnl_f = c8_net_pnl[cp8_mask]
    c8_pos_f       = carry_pos[cp8_mask]
    c8_gross_ret_f = c8_gross_ret_all[cp8_mask]; c8_net_ret_f = c8_net_ret_all[cp8_mask]

    def _carry_perf(daily_pnl: pd.Series, daily_ret: pd.Series, position: pd.Series, label: str) -> dict:
        active  = daily_ret[position != 0].dropna()
        pnl_act = daily_pnl[position != 0].dropna()
        n = len(active)
        if n < 20:
            return {}
        ann_r  = float(active.mean() * 252 * 100)
        ann_sd = float(active.std()  * np.sqrt(252) * 100)
        sharpe = ann_r / ann_sd if ann_sd > 0 else np.nan
        down   = active[active < 0]
        srt_d  = float(down.std() * np.sqrt(252) * 100) if len(down) > 1 else np.nan
        sortino = ann_r / srt_d if srt_d and srt_d > 0 else np.nan
        cum_r   = daily_ret.fillna(0).cumsum() * 100
        mdd_pct = float((cum_r - cum_r.cummax()).min())
        calmar  = ann_r / abs(mdd_pct) if mdd_pct != 0 else np.nan
        wins, losses = pnl_act[pnl_act > 0], pnl_act[pnl_act < 0]
        return {
            "label": label, "n": n,
            "sharpe": sharpe, "sortino": sortino,
            "ann_ret_pct": ann_r, "ann_std_pct": ann_sd,
            "mdd_pct": mdd_pct, "calmar": calmar,
            "hit_rate": float((active > 0).mean()) * 100,
            "profit_factor": abs(wins.sum() / losses.sum()) if len(losses) > 0 else np.nan,
            "total_pnl_usdmt": float(pnl_act.sum()),
        }

    cm8_gross = _carry_perf(c8_gross_pnl_f, c8_gross_ret_f, c8_pos_f, "Gross (No TC)")
    cm8_net   = _carry_perf(c8_net_pnl_f,   c8_net_ret_f,   c8_pos_f, f"Net ({carry_tc_label})")

    st.divider()
    section_header("PERFORMANCE METRICS")
    st.caption(f"**Strategy:** {carry_vgroup} / {carry_sub_label}  |  **Entry:** {carry_timing}  |  **TC:** {carry_tc_label}")

    if cm8_gross and cm8_net:
        _c8_cols1 = st.columns(8)
        _cmcard(_c8_cols1[0], "Sharpe Gross",    cm8_gross.get("sharpe"),      ".2f")
        _cmcard(_c8_cols1[1], "Sharpe Net",      cm8_net.get("sharpe"),        ".2f")
        _cmcard(_c8_cols1[2], "Sortino Gross",   cm8_gross.get("sortino"),     ".2f")
        _cmcard(_c8_cols1[3], "Sortino Net",     cm8_net.get("sortino"),       ".2f")
        _cmcard(_c8_cols1[4], "Ann Ret% Gross",  cm8_gross.get("ann_ret_pct"), ".2f", "%")
        _cmcard(_c8_cols1[5], "Ann Ret% Net",    cm8_net.get("ann_ret_pct"),   ".2f", "%")
        _cmcard(_c8_cols1[6], "Ann Std Dev%",    cm8_gross.get("ann_std_pct"), ".2f", "%")
        _cmcard(_c8_cols1[7], "Active Days",     float(cm8_gross.get("n", 0)), ",.0f")

        _c8_cols2 = st.columns(8)
        _cmcard(_c8_cols2[0], "Max DD% Gross",   cm8_gross.get("mdd_pct"),        ".2f", "%")
        _cmcard(_c8_cols2[1], "Max DD% Net",     cm8_net.get("mdd_pct"),          ".2f", "%")
        _cmcard(_c8_cols2[2], "Calmar Gross",    cm8_gross.get("calmar"),         ".2f")
        _cmcard(_c8_cols2[3], "Calmar Net",      cm8_net.get("calmar"),           ".2f")
        _cmcard(_c8_cols2[4], "Hit Rate",        cm8_gross.get("hit_rate"),       ".2f", "%")
        _cmcard(_c8_cols2[5], "Profit Factor",   cm8_gross.get("profit_factor"),  ".2f")
        _cmcard(_c8_cols2[6], "PnL Gross $/MT",  cm8_gross.get("total_pnl_usdmt"), ",.2f")
        _cmcard(_c8_cols2[7], "PnL Net $/MT",    cm8_net.get("total_pnl_usdmt"),   ",.2f")
    else:
        st.warning("Insufficient active trading days to compute metrics.")

    # ── Section 5: Rolling Sharpe ──────────────────────────────────────────────
    st.divider()
    section_header("ROLLING SHARPE RATIO")
    crs8_c1, crs8_c2 = st.columns([3, 1])
    with crs8_c2:
        crs8_win = st.radio("Window", ["1 Year (252d)", "2 Years (504d)", "Both"],
                             index=2, key="carry_rs_window", horizontal=False)
    _cdr8 = c8_gross_ret_all.fillna(0)
    croll8_252 = (_cdr8.rolling(252).mean() / _cdr8.rolling(252).std() * np.sqrt(252))
    croll8_504 = (_cdr8.rolling(504).mean() / _cdr8.rolling(504).std() * np.sqrt(252))
    with crs8_c1:
        fig_crs8 = go.Figure()
        if crs8_win in ("1 Year (252d)", "Both"):
            fig_crs8.add_trace(go.Scatter(
                x=croll8_252.index, y=croll8_252.values, name="Rolling Sharpe (1yr)",
                mode="lines", line=dict(color=COLORS["primary"], width=1.6),
                hovertemplate="%{x|%b %Y}<br>Sharpe (1yr): %{y:.2f}<extra></extra>",
            ))
        if crs8_win in ("2 Years (504d)", "Both"):
            fig_crs8.add_trace(go.Scatter(
                x=croll8_504.index, y=croll8_504.values, name="Rolling Sharpe (2yr)",
                mode="lines", line=dict(color=COLORS["amber"], width=1.6, dash="dot"),
                hovertemplate="%{x|%b %Y}<br>Sharpe (2yr): %{y:.2f}<extra></extra>",
            ))
        fig_crs8.add_hline(y=0,   line_dash="dash", line_color="#475569", line_width=1)
        fig_crs8.add_hline(y=0.5, line_dash="dot",  line_color=COLORS["green"],
                            line_width=0.8, annotation_text="0.5", annotation_position="right")
        fig_crs8.update_layout(
            **CHART_LAYOUT, height=320,
            title=dict(text=f"{carry_sub_label} — Rolling Sharpe ({carry_timing})", font=dict(size=13)),
            yaxis_title="Annualised Sharpe", xaxis_title=None, hovermode="x unified",
        )
        fig_crs8.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_crs8, use_container_width=True)
    st.caption("Carry tends to be regime-dependent — performs strongly during sustained backwardation cycles (e.g., 2006-2008, 2021-2022). "
               "Positive rolling Sharpe validates the strategy over that window.")

    # ── Section 6: Regime Statistics ──────────────────────────────────────────
    st.divider()
    section_header("REGIME STATISTICS")

    pct_back_c8 = float(_c8_back_mask.mean()) * 100
    pct_cont_c8 = float(_c8_cont_mask.mean()) * 100

    _active_back = c8_gross_ret_all[_c8_back_mask & (carry_pos != 0)].dropna()
    _active_cont = c8_gross_ret_all[_c8_cont_mask & (carry_pos != 0)].dropna()
    ret_in_back = float(_active_back.mean() * 252 * 100) if len(_active_back) > 5 else np.nan
    ret_in_cont = float(_active_cont.mean() * 252 * 100) if len(_active_cont) > 5 else np.nan
    n_flips_c8 = int(_c8_flip_mask.sum()) - 1

    _rs8_cols = st.columns(6)
    _cmcard(_rs8_cols[0], "% Days Backwardation", pct_back_c8, ".1f", "%")
    _cmcard(_rs8_cols[1], "% Days Contango",      pct_cont_c8, ".1f", "%")
    _cmcard(_rs8_cols[2], "Ann Ret in Back.",      ret_in_back, ".1f", "%")
    _cmcard(_rs8_cols[3], "Ann Ret in Contango",   ret_in_cont, ".1f", "%")
    _cmcard(_rs8_cols[4], "Avg Back. Duration",    float(avg_back_dur), ".0f", "d")
    _cmcard(_rs8_cols[5], "Regime Flips (Total)",  float(n_flips_c8), ",.0f")
    st.caption("Ann Ret in Back./Contango = annualized gross return on active trading days when the carry signal is in that regime.")

    # ── Section 7: Cumulative PnL + Carry vs Carry Comparison ─────────────────
    st.divider()
    section_header("CUMULATIVE PnL (USD/MT)")

    _c8_cmp_keys = list(_CARRY_CMP_OPTIONS.keys())
    cc8_c1, cc8_c2 = st.columns(2)
    with cc8_c1:
        carry_cmp_a = st.selectbox("Compare: Strategy A", _c8_cmp_keys, index=0, key="carry_cmp_a")
    with cc8_c2:
        carry_cmp_b = st.selectbox("Compare: Strategy B", _c8_cmp_keys, index=0, key="carry_cmp_b")

    fig_c8cum = go.Figure()
    _C8_CMP_COLORS = [COLORS["primary"], COLORS["amber"]]
    for _c8_lbl, _c8_ci in [(carry_cmp_a, 0), (carry_cmp_b, 1)]:
        _c8_spec = _CARRY_CMP_OPTIONS.get(_c8_lbl)
        if _c8_spec is None:
            continue
        _c8_cpnl = _carry_cum_pnl(c8_curve_px, _cash_cu8, cf1c, _c8_spec)
        if _c8_cpnl.empty:
            continue
        fig_c8cum.add_trace(go.Scatter(
            x=_c8_cpnl.index, y=_c8_cpnl.values, name=_c8_lbl, mode="lines",
            line=dict(color=_C8_CMP_COLORS[_c8_ci], width=1.8,
                      dash="dot" if _c8_ci == 1 else "solid"),
            hovertemplate=f"%{{x|%b %d, %Y}}<br>{_c8_lbl}: $%{{y:,.1f}}/MT<extra></extra>",
        ))

    fig_c8cum.add_trace(go.Scatter(
        x=c8_cum_gross.index, y=c8_cum_gross.values,
        name=f"{carry_sub_label} Gross", mode="lines",
        line=dict(color="#64748B", width=1.2),
        hovertemplate="%{x|%b %d, %Y}<br>Gross: $%{y:,.1f}/MT<extra></extra>",
        visible="legendonly",
    ))
    if carry_tc_bps > 0:
        fig_c8cum.add_trace(go.Scatter(
            x=c8_cum_net.index, y=c8_cum_net.values,
            name=f"{carry_sub_label} Net ({carry_tc_label})", mode="lines",
            line=dict(color="#94A3B8", width=1.2, dash="dot"),
            hovertemplate="%{x|%b %d, %Y}<br>Net: $%{y:,.1f}/MT<extra></extra>",
            visible="legendonly",
        ))

    fig_c8cum.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
    fig_c8cum.update_layout(
        **CHART_LAYOUT, height=420,
        title=dict(text="Carry Strategy Comparison — Cumulative PnL (Gross, USD/MT)", font=dict(size=13)),
        yaxis_title="Cumulative PnL (USD/MT)", xaxis_title=None, hovermode="x unified",
    )
    fig_c8cum.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_c8cum, use_container_width=True)

    # ── Section 8: Annual PnL ─────────────────────────────────────────────────
    st.divider()
    section_header("ANNUAL PnL BREAKDOWN (Gross, USD/MT)")

    c8_annual_pnl = c8_gross_pnl.resample("YE").sum()
    c8_annual_pnl.index = c8_annual_pnl.index.year
    c8_bar_colors = [COLORS["green"] if v >= 0 else COLORS["red"] for v in c8_annual_pnl.values]
    fig_c8ann = go.Figure(go.Bar(
        x=c8_annual_pnl.index.astype(str), y=c8_annual_pnl.values,
        marker_color=c8_bar_colors,
        hovertemplate="%{x}<br>PnL: $%{y:,.1f}/MT<extra></extra>",
    ))
    fig_c8ann.update_layout(
        **CHART_LAYOUT, height=300,
        title=dict(text="Annual PnL (Gross)", font=dict(size=13)),
        yaxis_title="PnL (USD/MT)", xaxis_title=None, showlegend=False,
    )
    st.plotly_chart(fig_c8ann, use_container_width=True)

    # ── Section 9: Tenor Comparison (V3 only) ─────────────────────────────────
    if carry_vgroup == "V3 — Long Slope":
        st.divider()
        section_header("TENOR PAIR COMPARISON — V3 LONG SLOPE")
        st.caption("Sharpe ratio for all 10 tenor pairs. Shows which part of the curve carries the most predictive power.")

        _v3_pairs = [(3,15),(4,16),(5,17),(6,18),(7,19),(8,20),(9,21),(10,22),(11,23),(12,24)]
        _v3_labels = [f"F{j}-F{k}" for j, k in _v3_pairs]
        _v3_sh_lag, _v3_sh_same = [], []

        for _vj, _vk in _v3_pairs:
            for _vsd, _vlst in [(False, _v3_sh_lag), (True, _v3_sh_same)]:
                _vsp = {"variant": "v3", "j": _vj, "k": _vk, "same_day": _vsd}
                _vcr = _carry_raw_signal(c8_curve_px, _cash_cu8, _vsp)
                if _vcr.empty:
                    _vlst.append(np.nan); continue
                _vi = _vcr.index.intersection(cf1c.index)
                _vcr = _vcr.reindex(_vi).dropna(); _vi = _vcr.index
                _vf1c = cf1c.reindex(_vi)
                _vsig = np.sign(_vcr.values); _vT = len(_vsig)
                _vpos = np.empty(_vT)
                if _vsd:
                    _vpos[:] = np.where(np.isfinite(_vsig), _vsig, 0.0)
                else:
                    _vpos[0] = 0.0
                    _vpos[1:] = np.where(np.isfinite(_vsig[:-1]), _vsig[:-1], 0.0)
                _vpos_s = pd.Series(_vpos, index=_vi)
                _vgpnl = _vpos_s * _vf1c.diff()
                _vgret = (_vgpnl / _vf1c.shift(1)).replace([np.inf, -np.inf], np.nan)
                _vact = _vgret[_vpos_s != 0].dropna()
                if len(_vact) < 20:
                    _vlst.append(np.nan); continue
                _vann_r = float(_vact.mean() * 252 * 100)
                _vann_sd = float(_vact.std() * np.sqrt(252) * 100)
                _vlst.append(_vann_r / _vann_sd if _vann_sd > 0 else np.nan)

        def _safe_color(val, pos_col, neg_col):
            if val is None or (isinstance(val, float) and np.isnan(val)):
                return "#475569"
            return pos_col if val >= 0 else neg_col

        fig_tenor = go.Figure()
        fig_tenor.add_trace(go.Bar(
            y=_v3_labels, x=_v3_sh_lag, name="Lag-1", orientation="h",
            marker_color=[_safe_color(v, COLORS["green"], COLORS["red"]) for v in _v3_sh_lag],
            hovertemplate="%{y}<br>Sharpe Lag-1: %{x:.3f}<extra></extra>",
        ))
        fig_tenor.add_trace(go.Bar(
            y=_v3_labels, x=_v3_sh_same, name="Same-Day", orientation="h",
            marker_color=[_safe_color(v, COLORS["amber"], "#7A3030") for v in _v3_sh_same],
            opacity=0.75,
            hovertemplate="%{y}<br>Sharpe Same-Day: %{x:.3f}<extra></extra>",
        ))
        fig_tenor.add_vline(x=0, line_dash="dash", line_color="#475569", line_width=1)
        fig_tenor.update_layout(
            **CHART_LAYOUT, height=420, barmode="group",
            title=dict(text="V3 Long Slope — Sharpe by Tenor Pair (Gross, No TC)", font=dict(size=13)),
            xaxis_title="Sharpe Ratio", yaxis_title=None,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_tenor, use_container_width=True)
        st.caption("Short-to-medium tenors (F3-F15, F4-F16) typically carry more predictive power than long-end pairs (F12-F24), "
                   "where price noise dominates. A decaying Sharpe across tenor pairs is a structural finding.")

    # ── Section 10: Signal & Position Over Time ────────────────────────────────
    st.divider()
    section_header("SIGNAL & POSITION OVER TIME")
    st.caption("F1_continuous price (USD/MT) with carry-driven position overlay. "
               "Unlike the Carry History chart above, this shows *which price moves* the position captured.")

    fig_c8sig = make_subplots(
        rows=2, cols=1, shared_xaxes=True, row_heights=[0.65, 0.35],
        vertical_spacing=0.04,
    )
    fig_c8sig.add_trace(go.Scatter(
        x=cf1c_a.index, y=cf1c_a.values, name="F1 Continuous ($/MT)",
        line=dict(color=COLORS["primary"], width=1.5),
        hovertemplate="%{x|%b %d, %Y}<br>F1_cont: $%{y:,.1f}<extra></extra>",
    ), row=1, col=1)

    _c8p_long  = carry_pos.where(carry_pos > 0, 0.0)
    _c8p_short = carry_pos.where(carry_pos < 0, 0.0)
    fig_c8sig.add_trace(go.Bar(
        x=carry_pos.index, y=_c8p_long.values,
        name="Long (+1)", marker_color="#00E676", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Long<extra></extra>",
    ), row=2, col=1)
    fig_c8sig.add_trace(go.Bar(
        x=carry_pos.index, y=_c8p_short.values,
        name="Short (-1)", marker_color="#FF1744", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Short<extra></extra>",
    ), row=2, col=1)

    fig_c8sig.update_layout(
        **CHART_LAYOUT, height=480, barmode="overlay",
        title=dict(text=f"{carry_sub_label} — F1 Price & Position ({carry_timing})", font=dict(size=13)),
        hovermode="x unified", showlegend=True,
    )
    fig_c8sig.update_yaxes(title_text="F1_cont ($/MT)", row=1, col=1)
    fig_c8sig.update_yaxes(title_text="Position", tickvals=[-1, 0, 1],
                             ticktext=["Short", "Flat", "Long"], row=2, col=1)
    fig_c8sig.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_c8sig, use_container_width=True)

    # ── Recent Signal Flips ────────────────────────────────────────────────────
    st.divider()
    section_header("RECENT SIGNAL FLIPS (Last 20)")

    _c8_flips_mask2 = carry_pos.diff().abs() > 0
    _c8_flips_mask2.iloc[0] = carry_pos.iloc[0] != 0
    _c8_flip_dates2 = carry_pos[_c8_flips_mask2].tail(20)

    if not _c8_flip_dates2.empty:
        _c8_flip_df = pd.DataFrame({
            "Date":            _c8_flip_dates2.index.strftime("%Y-%m-%d"),
            "Position":        _c8_flip_dates2.values.astype(int),
            "Direction":       ["LONG" if v > 0 else "SHORT" for v in _c8_flip_dates2.values],
            "F1_raw ($/MT)":   cf1r.reindex(_c8_flip_dates2.index).round(1).values,
            "Carry Value (%)": (carry_raw.reindex(_c8_flip_dates2.index) * 100).round(4).values,
            "Gross PnL":       c8_gross_pnl.reindex(_c8_flip_dates2.index).round(2).values,
            "TC cost":         c8_tc_cost.reindex(_c8_flip_dates2.index).round(2).values,
            "Net PnL":         c8_net_pnl.reindex(_c8_flip_dates2.index).round(2).values,
        })
        st.dataframe(_c8_flip_df, use_container_width=True, hide_index=True)
    else:
        st.info("No signal flips found in data.")

    # ── Methodology Notes ──────────────────────────────────────────────────────
    st.divider()
    with st.expander("Methodology Notes", expanded=False):
        st.markdown("""
**Economic Intuition**
Commodity carry (basis) reflects expected convenience yield and storage costs.
- **Backwardation** (nearby > far): tight nearby supply / positive convenience yield → *Long*
- **Contango** (nearby < far): excess supply / storage costs dominant → *Short*

**Signal Variant Formulas**
- **V1 Roll Yield**: `(F1-F2)/F1`, `(F1-F3)/F1`, `(Cash-3M)/Cash`
  Measures the 1-period roll cost as a fraction of current price.
- **V2 Annualized**: V1 x 12 (for F1-F2) or x 6 (for F1-F3) — scaled to annual rate.
  Note: binary signal (sign) is identical to V1; annualization only affects magnitude.
- **V3 Long Slope**: `(Fj-Fk)/Fk` for j < k (e.g., F3-F15, F4-F16, ... F12-F24)
  Measures the slope of the forward curve at longer tenors; downward slope = backwardation at the long end = Long signal.
- **V4 Z-score**: 252-day rolling standardization of (F1-F2)/F1.
  Filters out permanent level shifts; signal fires when carry is unusually high/low relative to recent history.

**Position Timing**
- *Same-Day*: `position[t] = sign(carry[t])` — taken at same close as signal
- *Lag-1*: `position[t] = sign(carry[t-1])` — entered the following day

**Why Same-Day Dominates for Carry**
Carry is a level-based signal. On flip days (contango→backwardation), a large price move
occurs in the direction of the new regime. Same-Day captures this move; Lag-1 takes a
2x swing loss (wrong position the whole day, corrects only the next morning). Unlike
momentum signals (which flip gradually), carry flips are discrete and high-impact events.

**Transaction Costs**
`TC[t] = |delta_position[t]| x (bps/10000/2) x F1_cont[t]`
Flip (+1 to -1): delta=2, cost = 1 full round-trip x price.
Entry (0 to +/-1): delta=1, cost = 0.5 round-trip x price.

**PnL & Returns**
All strategies trade F1_continuous regardless of which tenor pair generates the signal.
`daily_ret[t] = position[t] x delta_F1_cont[t] / F1_cont[t-1]`
Sharpe, Sortino, Max DD, Calmar all computed from daily_ret in % terms.

**Reference**
Baz, J., Granger, N. M. (2015). Dissecting Investment Strategies in the Cross Section and Time Series. SSRN.
        """)
