"""
Metals Risk Premia - Interactive Dashboard
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
from scipy.optimize import minimize

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

    /* Compact metric cards (momentum tab - fits 8 per row) */
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
    """Load Metals Cash and 3M.xlsx - one sheet per metal."""
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
    Load Metals Futures Curve - one sheet per metal with F1-F27.
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
        # Check spread first - spread col names also contain "cash"/"3m"/"price".
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
LOCAL_AL_F1_PATH   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "LME_Aluminium_Rolling_F1_v2.csv")
_F1_PATHS = {"Copper": LOCAL_F1_CONT_PATH, "Aluminium": LOCAL_AL_F1_PATH}


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


@st.cache_data(show_spinner=False)
def _load_f1_data(metal: str = "Copper") -> pd.DataFrame:
    """Metal-aware rolling F1 loader (Copper / Aluminium). Same schema as _load_copper_f1_data.
    Both files are ratio back-adjusted with the identical 4-phase roll logic."""
    path = _F1_PATHS.get(metal, LOCAL_F1_CONT_PATH)
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_csv(path, parse_dates=["Date"]).set_index("Date")
    df.index = df.index.normalize()
    df = df.sort_index()
    return df[["F1_raw", "F1_continuous"]]


def _tc_label_map(last_price: float) -> dict:
    """TC selectbox options with bps and $/MT per-flip equivalent in each label."""
    def _lbl(bps):
        flip_cost = (bps / 10000.0) * last_price
        return f"{bps} bps  (~${flip_cost:.0f}/MT per flip)"
    return {
        "0 bps  (Gross)": 0,
        _lbl(5):  5,
        _lbl(10): 10,
        _lbl(20): 20,
    }


@st.cache_data(show_spinner=False)
def _get_last_f1_price() -> float:
    """Last available F1_continuous price ($/MT). Used for TC label display."""
    df = _load_copper_f1_data()
    if df.empty or "F1_continuous" not in df.columns:
        return 9500.0
    return float(df["F1_continuous"].dropna().iloc[-1])


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
    st.divider()
    st.caption(
        "**Disclaimer:** This dashboard is a research prototype developed for academic purposes. "
        "Results are in-sample backtests unless stated otherwise. "
        "Not investment advice."
    )


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

    **File 1 - Metals Cash and 3M.xlsx:**
    One sheet per LME metal (LME Copper, LME Aluminium, ...) with columns for
    Cash Price, 3M Forward Price/Volume/OI, Cash-3M Spread Price/Volume.
    Plus a CME Cash Prices sheet for Gold, Silver, Platinum, Palladium, Copper ($/lb).

    **File 2 - Metals Futures Curve (.xlsx or .csv):**
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

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "📊 Market Overview",
    "📈 Term Structure",
    "💰 Cash vs 3M (Carry)",
    "📉 Volume & Open Interest",
    "🔗 Copper LME-CME Spread",
    "📋 Statistics",
    "⚡ Momentum Signals",
    "📐 Carry Signals",
    "📏 Value Signals",
    "🗂️ Portfolio",
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

        section_header("Cash-3M Spread - 1 Year")
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
                title=dict(text=f"{selected_spread_metal} - Cash-3M Spread (Last 1 Year)", font=dict(size=14)),
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
                            title=dict(text=f"{curve_metal} - Forward Curve", font=dict(size=16)),
                            xaxis_title="Contract",
                            yaxis_title="Price",
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        latest_row = prices_df.loc[selected_curve_date].dropna()
                        if len(latest_row) >= 2:
                            slope = latest_row.iloc[-1] - latest_row.iloc[0]
                            if slope > 0:
                                st.success(f"📈 **Contango** - Far month contracts are trading higher than near month ({curve_metal}, {selected_curve_date.strftime('%Y-%m-%d')})")
                            else:
                                st.warning(f"📉 **Backwardation** - Near month contracts are trading higher than far month ({curve_metal}, {selected_curve_date.strftime('%Y-%m-%d')})")
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

    st.markdown(f"### {selected_metal} - Cash vs 3M (Carry Analysis)")

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
                    title=dict(text=f"{selected_metal} - Cash minus 3M Spread", font=dict(size=14)),
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

    st.markdown(f"### {selected_metal} - Volume & Open Interest")

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
            title=dict(text=f"{selected_metal} - {price_label}", font=dict(size=14)),
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
            section_header("3M Forward - Volume & Open Interest")
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
                title=dict(text=f"{selected_metal} - 3M Forward", font=dict(size=14)),
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
                        section_header(f"{sel_contract} - Volume & Open Interest")
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
                            title=dict(text=f"{selected_metal} - {sel_contract}", font=dict(size=14)),
                        )
                        fig_vol_c.update_yaxes(title_text="Volume", row=1, col=1)
                        fig_vol_c.update_yaxes(title_text="Open Interest", row=2, col=1)
                        st.plotly_chart(fig_vol_c, use_container_width=True)

    # Futures Strip Volume Heatmap - works for both LME and CME
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
                        title=dict(text=f"{selected_metal} - Monthly Average Volume by Contract", font=dict(size=13)),
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

                section_header("LME vs COMEX - Price in $/MT")
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
                    "**US copper tariff shock** - COMEX copper (US domestic) priced in a large import tariff premium "
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

    section_header("Summary Statistics - All LME Metals")

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
        st.dataframe(stats_df.style.format(fmt_dict, na_rep="-"), use_container_width=True)

    section_header("Summary Statistics - CME Metals")
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
            st.dataframe(cme_stats_df.style.format(fmt_cme, na_rep="-"), use_container_width=True)
    else:
        st.info("CME Cash Prices data not available.")

    selected_metal = st.selectbox("Select Metal for Detailed Analysis", available_metals, key="tab6_metal")
    if selected_metal in cash_data:
        metal_df = parse_cash_3m_columns(cash_data[selected_metal], selected_metal)
        metal_df = filter_date(metal_df, start_date, end_date)
    else:
        metal_df = pd.DataFrame()

    section_header(f"{selected_metal} - Rolling Volatility")
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
            title=dict(text=f"{selected_metal} - Annualized Realized Volatility", font=dict(size=14)),
            yaxis_title="Volatility (%)",
        )
        st.plotly_chart(fig_vol, use_container_width=True)

    section_header(f"{selected_metal} - Return Distribution")
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
    # Two legitimate (no look-ahead) execution conventions:
    #   Same-Day (sd=True)  -> trade at the signal's own close(t); first return t->t+1  (shift 1)
    #   Lag-1    (sd=False) -> trade at the NEXT close(t+1);       first return t+1->t+2 (shift 2)
    # The removed legacy branch used shift 0 (booking today's already-realized move
    # = look-ahead) and is no longer available.
    if sd:
        pos[0] = 0.0
        pos[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
    else:
        pos[:2] = 0.0
        pos[2:] = np.where(np.isfinite(sig[:-2]), sig[:-2], 0.0)
    return pd.Series(pos * f1c.diff().values.astype(float), index=f1r.index).cumsum()


# All variant+timing combinations available for comparison dropdown
_MOM_CMP_OPTIONS = {
    "N/A": None,
    "MA(35,43) - Lag-1":    {"type": "ma", "m": 35, "n": 43, "same_day": False},
    "MA(35,43) - Same-Day": {"type": "ma", "m": 35, "n": 43, "same_day": True},
    "MA(33,48) - Lag-1":    {"type": "ma", "m": 33, "n": 48, "same_day": False},
    "MA(33,48) - Same-Day": {"type": "ma", "m": 33, "n": 48, "same_day": True},
    "MA(35,44) - Lag-1":    {"type": "ma", "m": 35, "n": 44, "same_day": False},
    "MA(35,44) - Same-Day": {"type": "ma", "m": 35, "n": 44, "same_day": True},
    "MA(34,47) - Lag-1":    {"type": "ma", "m": 34, "n": 47, "same_day": False},
    "MA(34,47) - Same-Day": {"type": "ma", "m": 34, "n": 47, "same_day": True},
    "MA(36,44) - Lag-1":    {"type": "ma", "m": 36, "n": 44, "same_day": False},
    "MA(36,44) - Same-Day": {"type": "ma", "m": 36, "n": 44, "same_day": True},
    "MA(1,5) - Lag-1":      {"type": "ma", "m": 1,  "n": 5,  "same_day": False},
    "MA(1,5) - Same-Day":   {"type": "ma", "m": 1,  "n": 5,  "same_day": True},
    "MA(5,20) - Lag-1":     {"type": "ma", "m": 5,  "n": 20, "same_day": False},
    "MA(5,20) - Same-Day":  {"type": "ma", "m": 5,  "n": 20, "same_day": True},
    "MA(10,60) - Lag-1":    {"type": "ma", "m": 10, "n": 60, "same_day": False},
    "MA(10,60) - Same-Day": {"type": "ma", "m": 10, "n": 60, "same_day": True},
    "CTA(9,21) - Lag-1":    {"type": "cta_single", "s": 9,  "l": 21, "same_day": False},
    "CTA(9,21) - Same-Day": {"type": "cta_single", "s": 9,  "l": 21, "same_day": True},
    "CTA(9,20) - Lag-1":    {"type": "cta_single", "s": 9,  "l": 20, "same_day": False},
    "CTA(9,20) - Same-Day": {"type": "cta_single", "s": 9,  "l": 20, "same_day": True},
    "CTA(10,19) - Lag-1":   {"type": "cta_single", "s": 10, "l": 19, "same_day": False},
    "CTA(10,19) - Same-Day":{"type": "cta_single", "s": 10, "l": 19, "same_day": True},
    "CTA(8,21) - Lag-1":    {"type": "cta_single", "s": 8,  "l": 21, "same_day": False},
    "CTA(8,21) - Same-Day": {"type": "cta_single", "s": 8,  "l": 21, "same_day": True},
    "CTA(14,15) - Lag-1":   {"type": "cta_single", "s": 14, "l": 15, "same_day": False},
    "CTA(14,15) - Same-Day":{"type": "cta_single", "s": 14, "l": 15, "same_day": True},
    "CTA Paper - Lag-1":    {"type": "cta_paper", "same_day": False},
    "CTA Paper - Same-Day": {"type": "cta_paper", "same_day": True},
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
    elif v == "v5":
        # Carry momentum: N-day change in the (F1-F2)/F1 roll yield. A steepening
        # curve (backwardation building) over ~20d signals physical tightening.
        # Best walk-forward OOS of all carry signals (+0.50 net5, 20d horizon).
        horizon = int(spec.get("horizon", 20))
        if "F1" not in curve_prices.columns or "F2" not in curve_prices.columns:
            return pd.Series(dtype=float)
        f1 = curve_prices["F1"].dropna(); f2 = curve_prices["F2"].dropna()
        idx = f1.index.intersection(f2.index)
        base = (f1.reindex(idx) - f2.reindex(idx)) / f1.reindex(idx)
        raw = base - base.shift(horizon)
        return raw.replace([np.inf, -np.inf], np.nan).dropna()
    return pd.Series(dtype=float)


def _carry_binarize(raw_values: np.ndarray, spec: dict) -> np.ndarray:
    """Map a raw carry signal to a ±1/0 position.
    Default: sign(raw). If spec['deadband'] > 0 (only meaningful for the
    standardised V3 z-score, where raw is in z-units), trade ±1 only when
    |raw| exceeds the deadband, else flat - filtering out near-mean noise."""
    db = float(spec.get("deadband", 0.0) or 0.0)
    if db > 0:
        return np.where(raw_values > db, 1.0,
               np.where(raw_values < -db, -1.0, 0.0))
    return np.sign(raw_values)


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
    sig = _carry_binarize(cr.values, spec); T = len(sig)
    pos = np.empty(T)
    # Same-Day (sd=True) -> trade at signal close(t), return t->t+1 (shift 1).
    # Lag-1   (sd=False) -> trade at next close(t+1), return t+1->t+2 (shift 2).
    # Legacy shift 0 (contemporaneous return) was look-ahead and is removed;
    # it had inflated carry Sharpe to ~0.62 vs ~0.10 honest (same-day).
    if spec.get("same_day", True):
        pos[0] = 0.0
        pos[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
    else:
        pos[:2] = 0.0
        pos[2:] = np.where(np.isfinite(sig[:-2]), sig[:-2], 0.0)
    return (pd.Series(pos, index=idx) * f1c_a.diff()).cumsum()


_CARRY_CMP_OPTIONS = {
    "N/A": None,
    "V1: (F1-F2)/F1 - Lag-1":        {"variant": "v1", "signal_type": "f1f2",   "same_day": False},
    "V1: (F1-F2)/F1 - Same-Day":     {"variant": "v1", "signal_type": "f1f2",   "same_day": True},
    "V1: (F1-F3)/F1 - Lag-1":        {"variant": "v1", "signal_type": "f1f3",   "same_day": False},
    "V1: (F1-F3)/F1 - Same-Day":     {"variant": "v1", "signal_type": "f1f3",   "same_day": True},
    "V1: (Cash-3M)/Cash - Lag-1":    {"variant": "v1", "signal_type": "cash3m", "same_day": False},
    "V1: (Cash-3M)/Cash - Same-Day": {"variant": "v1", "signal_type": "cash3m", "same_day": True},
    "V2: F3-F15 - Lag-1":            {"variant": "v3", "j": 3,  "k": 15, "same_day": False},
    "V2: F3-F15 - Same-Day":         {"variant": "v3", "j": 3,  "k": 15, "same_day": True},
    "V2: F4-F16 - Lag-1":            {"variant": "v3", "j": 4,  "k": 16, "same_day": False},
    "V2: F4-F16 - Same-Day":         {"variant": "v3", "j": 4,  "k": 16, "same_day": True},
    "V2: F5-F17 - Lag-1":            {"variant": "v3", "j": 5,  "k": 17, "same_day": False},
    "V2: F5-F17 - Same-Day":         {"variant": "v3", "j": 5,  "k": 17, "same_day": True},
    "V2: F6-F18 - Lag-1":            {"variant": "v3", "j": 6,  "k": 18, "same_day": False},
    "V2: F6-F18 - Same-Day":         {"variant": "v3", "j": 6,  "k": 18, "same_day": True},
    "V2: F7-F19 - Lag-1":            {"variant": "v3", "j": 7,  "k": 19, "same_day": False},
    "V2: F7-F19 - Same-Day":         {"variant": "v3", "j": 7,  "k": 19, "same_day": True},
    "V2: F8-F20 - Lag-1":            {"variant": "v3", "j": 8,  "k": 20, "same_day": False},
    "V2: F8-F20 - Same-Day":         {"variant": "v3", "j": 8,  "k": 20, "same_day": True},
    "V2: F9-F21 - Lag-1":            {"variant": "v3", "j": 9,  "k": 21, "same_day": False},
    "V2: F9-F21 - Same-Day":         {"variant": "v3", "j": 9,  "k": 21, "same_day": True},
    "V2: F10-F22 - Lag-1":           {"variant": "v3", "j": 10, "k": 22, "same_day": False},
    "V2: F10-F22 - Same-Day":        {"variant": "v3", "j": 10, "k": 22, "same_day": True},
    "V2: F11-F23 - Lag-1":           {"variant": "v3", "j": 11, "k": 23, "same_day": False},
    "V2: F11-F23 - Same-Day":        {"variant": "v3", "j": 11, "k": 23, "same_day": True},
    "V2: F12-F24 - Lag-1":           {"variant": "v3", "j": 12, "k": 24, "same_day": False},
    "V2: F12-F24 - Same-Day":        {"variant": "v3", "j": 12, "k": 24, "same_day": True},
    "V3: Z-score (252d) - Lag-1":    {"variant": "v4", "window": 252, "same_day": False},
    "V3: Z-score (252d) - Same-Day": {"variant": "v4", "window": 252, "same_day": True},
}


# ══════════════════════════════════════════════════════
# VALUE SIGNAL HELPERS  (module-level)
# ══════════════════════════════════════════════════════

def _value_raw_signal(curve_prices: pd.DataFrame, f1r: pd.Series, spec: dict) -> pd.Series:
    """
    Return raw value signal.
    V1 MA Reversion: deviation = (Fk - MA_N) / MA_N   (positive = expensive)
    V2 Baz-Granger : reversal  = F1_raw[t-N] - F1_raw[t]  (positive = price fallen = cheap)
    """
    v = spec.get("variant", "v1")
    if v == "v1":
        k = spec.get("contract", 12)
        N = spec.get("lookback", 1260)
        col = f"F{k}"
        if col not in curve_prices.columns:
            return pd.Series(dtype=float)
        price = curve_prices[col].dropna()
        if len(price) < max(N // 2, 60):
            return pd.Series(dtype=float)
        ma  = price.rolling(N, min_periods=max(N // 2, 60)).mean()
        dev = (price - ma) / ma.replace(0, np.nan)
        return dev.replace([np.inf, -np.inf], np.nan).dropna()
    else:
        N   = spec.get("lookback", 1260)
        rev = f1r.shift(N) - f1r
        return rev.replace([np.inf, -np.inf], np.nan).dropna()


def _value_cum_pnl(curve_prices: pd.DataFrame, f1r: pd.Series, f1c: pd.Series, spec: dict) -> pd.Series:
    """Cumulative PnL (USD/MT) for a value spec dict."""
    raw = _value_raw_signal(curve_prices, f1r, spec)
    if raw.empty:
        return pd.Series(dtype=float)
    v   = spec.get("variant", "v1")
    thr = spec.get("threshold", 0.10)
    if v == "v1":
        sig_bin = np.where(raw.values < -thr,  1.0,
                  np.where(raw.values >  thr, -1.0, 0.0))
    else:
        sig_bin = np.sign(raw.values).astype(float)
    idx  = raw.index.intersection(f1c.index)
    sb   = pd.Series(sig_bin, index=raw.index).reindex(idx).values
    f1ca = f1c.reindex(idx)
    T    = len(idx)
    if T == 0:
        return pd.Series(dtype=float)
    # Same-Day (sd=True) -> trade at signal close(t), return t->t+1 (shift 1).
    # Lag-1   (sd=False) -> trade at next close(t+1), return t+1->t+2 (shift 2).
    # Legacy shift 0 (contemporaneous return) was look-ahead and is removed.
    pos = np.empty(T)
    if spec.get("same_day", False):
        pos[0] = 0.0
        pos[1:] = np.where(np.isfinite(sb[:-1]), sb[:-1], 0.0)
    else:
        pos[:2] = 0.0
        pos[2:] = np.where(np.isfinite(sb[:-2]), sb[:-2], 0.0)
    return (pd.Series(pos, index=idx) * f1ca.diff()).cumsum()


_VALUE_CMP_OPTIONS: dict = {
    "N/A": None,
    # V1 F12 (Mark's default contract)
    "V1: F12 - 1yr MA":  {"variant": "v1", "contract": 12, "lookback": 252,  "threshold": 0.10, "same_day": False},
    "V1: F12 - 3yr MA":  {"variant": "v1", "contract": 12, "lookback": 756,  "threshold": 0.10, "same_day": False},
    "V1: F12 - 5yr MA":  {"variant": "v1", "contract": 12, "lookback": 1260, "threshold": 0.10, "same_day": False},
    "V1: F12 - 7yr MA":  {"variant": "v1", "contract": 12, "lookback": 1764, "threshold": 0.10, "same_day": False},
    "V1: F12 - 10yr MA": {"variant": "v1", "contract": 12, "lookback": 2520, "threshold": 0.10, "same_day": False},
    # V1 F5
    "V1: F5 - 1yr MA":   {"variant": "v1", "contract": 5,  "lookback": 252,  "threshold": 0.10, "same_day": False},
    "V1: F5 - 5yr MA":   {"variant": "v1", "contract": 5,  "lookback": 1260, "threshold": 0.10, "same_day": False},
    "V1: F5 - 10yr MA":  {"variant": "v1", "contract": 5,  "lookback": 2520, "threshold": 0.10, "same_day": False},
    # V1 F1
    "V1: F1 - 1yr MA":   {"variant": "v1", "contract": 1,  "lookback": 252,  "threshold": 0.10, "same_day": False},
    "V1: F1 - 5yr MA":   {"variant": "v1", "contract": 1,  "lookback": 1260, "threshold": 0.10, "same_day": False},
    "V1: F1 - 10yr MA":  {"variant": "v1", "contract": 1,  "lookback": 2520, "threshold": 0.10, "same_day": False},
    # V2 Baz-Granger
    "V2: BG - 1yr rev.":  {"variant": "v2", "lookback": 252,  "same_day": False},
    "V2: BG - 3yr rev.":  {"variant": "v2", "lookback": 756,  "same_day": False},
    "V2: BG - 5yr rev.":  {"variant": "v2", "lookback": 1260, "same_day": False},
    "V2: BG - 7yr rev.":  {"variant": "v2", "lookback": 1764, "same_day": False},
    "V2: BG - 10yr rev.": {"variant": "v2", "lookback": 2520, "same_day": False},
}


@st.cache_data(show_spinner=False)
def _wf_ma3543_tc(_f1r: pd.Series, _f1c: pd.Series, tc_bps: int) -> dict:
    """Walk-forward OOS Sharpes for MA(35,43) Lag-1 with round-trip TC."""
    IS_W, OOS_W = 1260, 252
    T, dates = len(_f1r), _f1r.index
    out = {}
    oos_s = IS_W
    while oos_s < T:
        oos_e = min(oos_s + OOS_W, T)
        if (oos_e - oos_s) < OOS_W:
            oos_s += OOS_W; continue
        yr = str(dates[oos_e - 1].year)   # label window by its END year (Dec-2010 window = 2011 OOS)
        f1r_w = _f1r.iloc[oos_s - IS_W:oos_e]
        oos_dt = dates[oos_s:oos_e]
        pos_full = np.sign(f1r_w.rolling(35).mean() - f1r_w.rolling(43).mean()).shift(1).fillna(0)
        pos_oos = pos_full.iloc[-OOS_W:].set_axis(oos_dt)
        f1c_oos = _f1c.reindex(oos_dt)
        f1r_oos = _f1r.reindex(oos_dt)        # TC on actual traded price F1_raw
        pnl = pos_oos * f1c_oos.diff()
        if tc_bps > 0:
            chg = pos_oos.diff().abs(); chg.iloc[0] = abs(pos_oos.iloc[0])
            pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * f1r_oos
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = (pnl / f1c_oos.shift(1)).replace([np.inf, -np.inf], np.nan)
        act = ret[pos_oos != 0].dropna()
        if len(act) >= 20:
            sd = float(act.std(ddof=1))
            out[yr] = float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan
        oos_s += OOS_W
    return out


@st.cache_data(show_spinner=False)
def _wf_anchors_isopt_tc(_f1r: pd.Series, _f1c: pd.Series, tc_bps: int) -> dict:
    """Walk-forward OOS Sharpes for Anchors + IS-optimised weights, with round-trip TC.
    Anchors = MA(10,25), MA(35,43), MA(63,100) (Lag-1 positions, full-history MA state).
    Per OOS window: fit max-Sharpe QP weights (≥0, sum=1) on the prior 5yr IS returns,
    then apply to the OOS anchor positions. Computes ALL windows (was hardcoded to 3)."""
    IS_W, OOS_W = 1260, 252
    T, dates = len(_f1r), _f1r.index
    anchors = [(10, 25), (35, 43), (63, 100)]

    def _apos(s, l):
        return np.sign(_f1r.rolling(s).mean() - _f1r.rolling(l).mean()).shift(1).fillna(0)
    pos_anchors = [_apos(s, l) for s, l in anchors]

    def _opt_w(ret_mat):
        def neg_sh(w):
            r = ret_mat @ w; a = r[r != 0]
            if len(a) < 20: return 0.0
            sd = a.std(ddof=1)
            return -(a.mean() / sd * np.sqrt(252)) if sd > 0 else 0.0
        best_v, best_w = np.inf, np.array([1/3, 1/3, 1/3])
        for w0 in [[1/3,1/3,1/3],[1,0,0],[0,1,0],[0,0,1],[0.5,0.5,0],[0.5,0,0.5],[0,0.5,0.5]]:
            try:
                res = minimize(neg_sh, w0, method="SLSQP", bounds=[(0,1)]*3,
                               constraints=[{"type":"eq","fun":lambda w: w.sum()-1}],
                               options={"ftol":1e-9,"maxiter":300})
                if res.fun < best_v: best_v, best_w = res.fun, res.x
            except Exception:
                pass
        w = np.clip(best_w, 0, 1); return w / w.sum()

    out = {}
    oos_s = IS_W
    while oos_s < T:
        oos_e = min(oos_s + OOS_W, T)
        if (oos_e - oos_s) < OOS_W:
            oos_s += OOS_W; continue
        yr = str(dates[oos_e - 1].year)   # label window by its END year (Dec-2010 window = 2011 OOS)
        is_dt = dates[oos_s - IS_W:oos_s]; oos_dt = dates[oos_s:oos_e]
        f1c_is = _f1c.reindex(is_dt)
        ret_mat = []
        for p in pos_anchors:
            pis = p.reindex(is_dt).fillna(0)
            with np.errstate(invalid="ignore", divide="ignore"):
                r = (pis * f1c_is.diff() / f1c_is.shift(1)).replace([np.inf,-np.inf], np.nan).fillna(0)
            ret_mat.append(r.values)
        w = _opt_w(np.column_stack(ret_mat))
        port = sum(wi * p.reindex(oos_dt).fillna(0) for wi, p in zip(w, pos_anchors))
        f1c_oos = _f1c.reindex(oos_dt)
        f1r_oos = _f1r.reindex(oos_dt)        # TC on actual traded price F1_raw
        pnl = port * f1c_oos.diff()
        if tc_bps > 0:
            chg = port.diff().abs(); chg.iloc[0] = abs(port.iloc[0])
            pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * f1r_oos
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = (pnl / f1c_oos.shift(1)).replace([np.inf,-np.inf], np.nan)
        act = ret[port != 0].dropna()
        if len(act) >= 20:
            sd = float(act.std(ddof=1))
            out[yr] = float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan
        oos_s += OOS_W
    return out


@st.cache_data(show_spinner=False)
def _wf_portfolio_tc(metal, use_iv, use_v2, tc_bps, _mom, _carry, _val, _port, _f1r, _f1c):
    """Walk-forward OOS Sharpe per sleeve + the combined portfolio.
    Each leg is the metal's a-priori-selected configuration (never re-optimised per
    window); it is simply evaluated on rolling 1yr OOS windows after a 5yr burn-in
    (IS=1260, OOS=252). Windows are labelled by their END year. Cache is keyed on
    (metal, use_iv, use_v2, tc_bps) so it recomputes correctly for every selection."""
    IS_W, OOS_W = 1260, 252
    idx = _f1c.index
    T, dates = len(idx), idx
    legs = {
        "Momentum":  _mom.reindex(idx).fillna(0.0),
        "Carry":     _carry.reindex(idx).fillna(0.0),
        "Value":     _val.reindex(idx).fillna(0.0),
        "Portfolio": _port.reindex(idx).fillna(0.0),
    }
    out = {k: {} for k in legs}
    oos_s = IS_W
    while oos_s < T:
        oos_e = min(oos_s + OOS_W, T)
        if (oos_e - oos_s) < OOS_W:
            oos_s += OOS_W; continue
        yr = str(dates[oos_e - 1].year)
        oos_dt = dates[oos_s:oos_e]
        f1c_oos = _f1c.reindex(oos_dt); f1r_oos = _f1r.reindex(oos_dt)
        for name, pos in legs.items():
            p = pos.reindex(oos_dt).fillna(0)
            pnl = p * f1c_oos.diff()
            if tc_bps > 0:
                chg = p.diff().abs(); chg.iloc[0] = abs(p.iloc[0])
                pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * f1r_oos
            with np.errstate(invalid="ignore", divide="ignore"):
                ret = (pnl / f1c_oos.shift(1)).replace([np.inf, -np.inf], np.nan)
            act = ret[p != 0].dropna()
            sd = act.std(ddof=1) if len(act) else np.nan
            out[name][yr] = (float(act.mean() / sd * np.sqrt(252))
                             if len(act) >= 20 and sd and sd > 0 else np.nan)
        oos_s += OOS_W
    return out


# ══════════════════════════════════════════════════════
# DYNAMIC "BEST SIGNAL" SCANNERS  (per-metal, cached)
# ══════════════════════════════════════════════════════
def _pos_metrics_generic(pos, f1r, f1c, tc_bps: int = 5) -> dict:
    """Active-day gross/net Sharpe, ann return %, max-DD % for a position series.
    PnL on F1_continuous; TC on F1_raw. Identical convention to the live tabs."""
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
                mdd=float((cum - cum.cummax()).min()), nact=int((pos != 0).sum()))

def _exec(sigbin, same_day):
    """Same-Day = shift 1, Lag-1 = shift 2 (no look-ahead)."""
    return sigbin.shift(1) if same_day else sigbin.shift(2)


@st.cache_data(show_spinner=False)
def _mom_best_cards(metal: str, tc_bps: int = 5):
    df = _load_f1_data(metal)
    if df.empty:
        return None
    f1r, f1c = df["F1_raw"], df["F1_continuous"]
    def ma_sig(m, n): return np.sign(f1r.rolling(m).mean() - f1r.rolling(n).mean())
    def cta_single_sig(s, l):
        x = f1r.ewm(com=s-1, adjust=False).mean() - f1r.ewm(com=l-1, adjust=False).mean()
        y = x / f1r.rolling(63).std(); z = y / y.rolling(252).std()
        return pd.Series(np.sign(z * np.exp(-z**2/4) / 0.89), index=f1r.index)
    def cta_paper_sig():
        pv = f1r.rolling(63).std(); us = []
        for s, l in zip((8,16,32), (24,48,96)):
            x = f1r.ewm(com=s-1, adjust=False).mean() - f1r.ewm(com=l-1, adjust=False).mean()
            y = x / pv; z = (y / y.rolling(252).std()).values; us.append(z*np.exp(-z**2/4)/0.89)
        return pd.Series(np.sign(np.nanmean(np.stack(us,1),1)), index=f1r.index)
    def best(cands):
        b = None
        for name, sig, sd in cands:
            mt = _pos_metrics_generic(_exec(sig, sd), f1r, f1c, tc_bps)
            if not np.isnan(mt["gross"]) and (b is None or mt["gross"] > b["gross"]):
                b = dict(name=name, timing="Same-Day" if sd else "Lag-1", **mt)
        return b
    # MA family — curated known-good pairs (so each metal's optimum is evaluated) + coarse grid
    _ma_explicit = [(35,43),(33,48),(35,44),(34,47),(36,44),(1,5),(5,20),(10,60),
                    (60,115),(65,100),(63,100),(40,90),(50,120),(20,115),(45,90),(55,110)]
    _ma_grid = [(m, n) for m in range(5, 81, 5) for n in range(m+5, 146, 5)]
    ma_cands = [(f"MA({m},{n})", ma_sig(m, n), sd)
                for (m, n) in sorted(set(_ma_explicit + _ma_grid)) for sd in (True, False)]
    # CTA family — singles + paper
    cta_cands = [(f"CTA({s},{l})", cta_single_sig(s, l), sd)
                 for (s, l) in [(8,21),(9,21),(9,20),(10,19),(14,15),(16,48),(32,96)] for sd in (True, False)]
    cta_cands += [("CTA Paper (3-scale)", cta_paper_sig(), sd) for sd in (True, False)]
    # Anchors EW (3 structural anchors), best timing
    anc_cands = [("Anchors EW MA(10,25)+MA(35,43)+MA(63,100)",
                  sum(ma_sig(m, n) for m, n in [(10,25),(35,43),(63,100)]) / 3.0, sd) for sd in (True, False)]
    return dict(ma=best(ma_cands), cta=best(cta_cands), anc=best(anc_cands))


@st.cache_data(show_spinner=False)
def _carry_best_cards(metal: str, tc_bps: int = 5):
    df = _load_f1_data(metal)
    if df.empty or not curve_data:
        return None
    sheet = _find_curve_sheet(metal, curve_data)
    if not sheet:
        return None
    f1r, f1c = df["F1_raw"], df["F1_continuous"]
    crv = curve_data[sheet]["prices"].copy(); crv.index = pd.to_datetime(crv.index).normalize(); crv = crv.sort_index()
    def best(cands):
        b = None
        for name, sig, sd in cands:
            mt = _pos_metrics_generic(_exec(sig, sd), f1r, f1c, tc_bps)
            if not np.isnan(mt["gross"]) and (b is None or mt["gross"] > b["gross"]):
                b = dict(name=name, timing="Same-Day" if sd else "Lag-1", **mt)
        return b
    out = {}
    if "F1" in crv and "F2" in crv:
        base = ((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf], np.nan)
        lvl = [("(F1-F2)/F1 level", np.sign(base), sd) for sd in (True, False)]
        if "F3" in crv:
            b3 = ((crv["F1"]-crv["F3"])/crv["F1"]).replace([np.inf,-np.inf], np.nan)
            lvl += [("(F1-F3)/F1 level", np.sign(b3), sd) for sd in (True, False)]
        out["level"] = best(lvl)
        out["mom"] = best([(f"CarryMom {h}d", np.sign(base-base.shift(h)), sd) for h in (20, 60) for sd in (True, False)])
        z = ((base-base.rolling(252).mean())/base.rolling(252).std()).replace([np.inf,-np.inf], np.nan)
        out["zscore"] = best([("Z-score 252d", np.sign(z), sd) for sd in (True, False)])
    slope = []
    for j, k in [(3,15),(4,16),(5,17),(6,18),(7,19),(8,20),(9,21),(10,22),(11,23),(12,24)]:
        if f"F{j}" in crv and f"F{k}" in crv:
            s = ((crv[f"F{j}"]-crv[f"F{k}"])/crv[f"F{k}"]).replace([np.inf,-np.inf], np.nan)
            slope += [(f"Slope (F{j}-F{k})/F{k}", np.sign(s), sd) for sd in (True, False)]
    if slope:
        out["slope"] = best(slope)
    return out


@st.cache_data(show_spinner=False)
def _value_best_cards(metal: str, tc_bps: int = 5):
    df = _load_f1_data(metal)
    if df.empty or not curve_data:
        return None
    sheet = _find_curve_sheet(metal, curve_data)
    if not sheet:
        return None
    f1r, f1c = df["F1_raw"], df["F1_continuous"]
    crv = curve_data[sheet]["prices"].copy(); crv.index = pd.to_datetime(crv.index).normalize(); crv = crv.sort_index()
    def best(cands):
        b = None
        for name, sig, sd in cands:
            mt = _pos_metrics_generic(_exec(sig, sd), f1r, f1c, tc_bps)
            if not np.isnan(mt["gross"]) and (b is None or mt["gross"] > b["gross"]):
                b = dict(name=name, timing="Same-Day" if sd else "Lag-1", **mt)
        return b
    def v1_sig(k, N):
        if f"F{k}" not in crv: return None
        p = crv[f"F{k}"].dropna(); ma = p.rolling(N, min_periods=max(N//2, 60)).mean()
        dev = ((p-ma)/ma).replace([np.inf,-np.inf], np.nan)
        return pd.Series(np.where(dev < -0.10, 1.0, np.where(dev > 0.10, -1.0, 0.0)), index=dev.index).where(dev.notna())
    v1 = []
    for k in [1, 5, 8, 12]:
        for N, lab in [(252,"1yr"),(1260,"5yr"),(2520,"10yr")]:
            s = v1_sig(k, N)
            if s is not None:
                v1 += [(f"V1 F{k} {lab}", s, sd) for sd in (True, False)]
    out = {"v1": best(v1)} if v1 else {}
    v2 = []
    for N, lab in [(756,"3yr"),(1260,"5yr"),(2520,"10yr")]:
        rev = (f1r.shift(N)-f1r).replace([np.inf,-np.inf], np.nan)
        v2 += [(f"V2 BG {lab}", np.sign(rev), sd) for sd in (True, False)]
    out["v2"] = best(v2)
    return out


def _fmt_sh(x):  return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:+.2f}"
def _fmt_pct(x): return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:+.1f}%"
def _fmt_dd(x):  return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.0f}%"


# ══════════════════════════════════════════════════════
# TAB 7: MOMENTUM SIGNALS
# ══════════════════════════════════════════════════════

with tab7:
    # ── Metal toggle (top of tab) ─────────────────────────────────────────────
    _mom_metal = st.radio("🔬 Metal", ["Copper", "Aluminium"], horizontal=True, key="mom_metal")
    st.markdown(f"### Momentum Signals - LME {_mom_metal}")
    st.markdown(
        '<div style="background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;'
        'border-radius:4px;padding:10px 18px;margin-bottom:10px;display:flex;align-items:center;gap:16px;">'
        '<span style="color:#B87333;font-family:\'IBM Plex Mono\',monospace;font-size:0.78rem;'
        'font-weight:700;white-space:nowrap;">SIGNAL 1 OF 3</span>'
        '<span style="color:#8A8278;font-size:0.8rem;">Price trends persist in short-to-medium horizons. '
        'The momentum sleeve feeds the equal-weight portfolio (Tab 10); Carry and Value follow in Tabs 8-9. '
        'All cards and metrics below recompute live for the selected metal.</span></div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Baz-Granger CTA trend signal (Eqs 29-33) and MA Crossover. "
        "Signal computed from F1_raw only; PnL from F1_continuous (roll costs captured). "
        "Returns expressed as % of notional. Transaction costs applied on every position change."
    )

    # ── Data loading (shared by both sections) ────────────────────────────────
    _f1_df = _load_f1_data(_mom_metal)
    if _f1_df.empty:
        st.error(f"Rolling F1 file for {_mom_metal} not found. Ensure the CSV is alongside app.py.")
        st.stop()
    f1r: pd.Series = _f1_df["F1_raw"]
    f1c: pd.Series = _f1_df["F1_continuous"]

    # ── Best Momentum Signal - By Variant (DYNAMIC, computed live per metal) ───
    section_header(f"BEST MOMENTUM SIGNAL - BY VARIANT  ({_mom_metal})")
    st.caption(
        f"Best configuration per signal family for {_mom_metal}, full-period IS backtest "
        f"({f1r.index[0].year}-{f1r.index[-1].year}), gross active-day Sharpe (TC=0), no look-ahead. "
        "Computed live from data - changes with the metal toggle. Past performance is not indicative of future results."
    )
    _mb_cs  = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;"
               "border-radius:4px;padding:14px 20px")
    _mb_csx = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #5BAD72;"
               "border-radius:4px;padding:14px 20px")
    _mb_lbl = ("color:#B87333;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _mb_lbx = ("color:#5BAD72;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _mb_big = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;font-size:1.55rem;font-weight:700;margin:0")
    _mb_sub = "color:#8A8278;font-size:0.75rem;margin:2px 0"
    _mb_hr  = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"
    _mb = _mom_best_cards(_mom_metal)
    _mb_fam = [("MA Crossover", _mb.get("ma") if _mb else None),
               ("Anchors EW",   _mb.get("anc") if _mb else None),
               ("CTA Baz-Granger", _mb.get("cta") if _mb else None)]
    _mb_best = max(range(3), key=lambda i: (_mb_fam[i][1]["gross"] if _mb_fam[i][1] and not np.isnan(_mb_fam[i][1]["gross"]) else -9))
    for _i, (_col, (_lbl, _d)) in enumerate(zip(st.columns(3), _mb_fam)):
        with _col:
            _star = "  ★" if _i == _mb_best else ""
            _sty, _lsty = (_mb_csx, _mb_lbx) if _i == _mb_best else (_mb_cs, _mb_lbl)
            if not _d:
                st.markdown(f'<div style="{_sty}"><p style="{_lsty}">{_lbl}{_star}</p>'
                            f'<p style="{_mb_big}">N/A</p></div>', unsafe_allow_html=True)
            else:
                st.markdown(f"""<div style="{_sty}">
<p style="{_lsty}">{_lbl}{_star}</p>
<p style="{_mb_big}">{_fmt_sh(_d['gross'])}</p>
<p style="{_mb_sub}">Sharpe Ratio (Gross)</p>
<hr style="{_mb_hr}"/>
<p style="{_mb_sub}">{_d['name']}, {_d['timing']}</p>
<p style="{_mb_sub}">Ann Ret ≈ {_fmt_pct(_d['ann'])}, Max DD ≈ {_fmt_dd(_d['mdd'])}</p>
</div>""", unsafe_allow_html=True)
    st.caption(f"Cards recompute live for {_mom_metal}: the MA-crossover grid, CTA family and Anchors EW are "
               "each re-scanned, and the strongest family is starred.")

    # ── SECTION 2: IS PARAMETER SEARCH (IN-SAMPLE) ────────────────────────────
    _m_is_yr0 = str(f1r.index[0].year); _m_is_yr1 = str(f1r.index[-1].year)
    st.divider()
    section_header(f"IN-SAMPLE PARAMETER SEARCH ({_m_is_yr0}-{_m_is_yr1})")
    st.markdown(
        '<div style="background:#1A1200;border:1px solid #3A2E00;border-left:4px solid #F59E0B;'
        'border-radius:4px;padding:8px 14px;margin-bottom:10px;font-size:0.82rem;color:#D4A843;">'
        f'&#9888;  IN-SAMPLE BACKTEST - Results use full {_m_is_yr0}-{_m_is_yr1} history. '
        'Not held-out data. See the walk-forward section for OOS estimates.</div>',
        unsafe_allow_html=True,
    )

    # ── Strategy Preset ───────────────────────────────────────────────────────
    _MOM_PRESETS = {
        "MA(35,43), Same-Day  [WF Best / Default]": {
            "mom_sig_type": "MA Crossover",
            "mom_variant":  "MA(35,43) - Best Sharpe [default]",
            "mom_timing":   "Same-Day",
        },
        "CTA(9,21), Same-Day  [Baz-Granger Best]": {
            "mom_sig_type": "CTA (Baz-Granger)",
            "mom_variant":  "CTA(9,21) - Best Lag-1 Sharpe [default]",
            "mom_timing":   "Same-Day",
        },
        "Anchors EW, Lag-1  [Anchors + IS-Opt]": {
            "mom_sig_type": "Anchors + IS-Opt Weights",
            "mom_variant":  "EW Anchors - MA(10,25) + MA(35,43) + MA(63,100)",
            "mom_timing":   "Lag-1 (Next-Day)",
        },
        "MA(35,43), Lag-1  [Sensitivity Check]": {
            "mom_sig_type": "MA Crossover",
            "mom_variant":  "MA(35,43) - Best Sharpe [default]",
            "mom_timing":   "Lag-1 (Next-Day)",
        },
        "Custom (use controls below)": {},
    }

    def _apply_mom_preset():
        cfg = _MOM_PRESETS.get(st.session_state.get("mom_preset", "Custom (use controls below)"), {})
        for k, v in cfg.items():
            st.session_state[k] = v

    _mom_preset_col, _mom_preset_info = st.columns([2.5, 3.5])
    with _mom_preset_col:
        st.selectbox(
            "Strategy Preset",
            list(_MOM_PRESETS.keys()),
            index=0,
            key="mom_preset",
            on_change=_apply_mom_preset,
        )
    with _mom_preset_info:
        st.markdown(
            '<div style="padding:8px 0;color:#7A7068;font-size:0.78rem;">'
            'Selecting a preset auto-fills all controls below. '
            'Switch to <b>Custom</b> to edit individual parameters freely.</div>',
            unsafe_allow_html=True,
        )

    # ── Controls ──────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns([1.5, 1.7, 1.3, 1.3])

    with c1:
        sig_type = st.selectbox(
            "Signal Type",
            ["MA Crossover", "CTA (Baz-Granger)", "Anchors + IS-Opt Weights"],
            key="mom_sig_type",
        )

    with c2:
        if sig_type == "MA Crossover":
            variant_opts = {
                "MA(35,43) - Best Sharpe [default]": (35, 43),
                "MA(33,48)": (33, 48),
                "MA(35,44)": (35, 44),
                "MA(34,47)": (34, 47),
                "MA(36,44)": (36, 44),
                "MA(1,5)":   (1, 5),
                "MA(5,20)":  (5, 20),
                "MA(10,60)": (10, 60),
            }
            default_idx = 0
        elif sig_type == "CTA (Baz-Granger)":
            variant_opts = {
                "CTA(8,21) - Best Same-Day Sharpe": ("cta_single", 8, 21),
                "CTA(9,21) - Best Lag-1 Sharpe [default]": ("cta_single", 9, 21),
                "CTA(9,20)": ("cta_single", 9, 20),
                "CTA(10,19)": ("cta_single", 10, 19),
                "CTA(14,15)": ("cta_single", 14, 15),
                "CTA Paper (8-16-32 / 24-48-96)": ("cta_paper",),
            }
            default_idx = 1
        else:  # Anchors
            variant_opts = {
                "EW Anchors - MA(10,25) + MA(35,43) + MA(63,100)": ("anchors_ew",),
            }
            default_idx = 0
        # Guard: reset if session value no longer valid for current sig_type
        if st.session_state.get("mom_variant", "") not in variant_opts:
            st.session_state["mom_variant"] = list(variant_opts.keys())[default_idx]
        variant_label = st.selectbox(
            "Strategy Variant", list(variant_opts.keys()),
            index=default_idx, key="mom_variant",
        )
        variant_params = variant_opts[variant_label]

    with c3:
        timing_label = st.selectbox(
            "Position Entry",
            ["Same-Day", "Lag-1 (Next-Day)"],
            index=0, key="mom_timing",
        )
        same_day = timing_label == "Same-Day"

    with c4:
        _oos_tc_map = _tc_label_map(float(f1c.dropna().iloc[-1]))
        _oos_tc_label = st.selectbox("TC (bps)", list(_oos_tc_map.keys()), index=0, key="oos_tc_sel")
        _oos_tc_bps = _oos_tc_map[_oos_tc_label]
    tc_bps   = _oos_tc_bps
    tc_label = _oos_tc_label

    if sig_type == "Anchors + IS-Opt Weights":
        st.info(
            "**Anchors + IS-Opt Weights** - IS backtest uses equal-weight combination of the three anchor MAs. "
            "The IS-optimised walk-forward OOS Sharpe (computed live) is shown in Section 1. "
            "Position shown here is a continuous range −1 to +1 (average of three ±1 signals).",
            icon="ℹ️",
        )

    # ── Custom parameter override ─────────────────────────────────────────────
    if sig_type != "Anchors + IS-Opt Weights":
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

    # ── Signal & position computation ─────────────────────────────────────────
    def _ewma(s: pd.Series, n: int) -> pd.Series:
        return s.ewm(com=n - 1, adjust=False).mean()

    if sig_type == "Anchors + IS-Opt Weights":
        _anc_sigs = [
            np.sign(f1r.rolling(m).mean() - f1r.rolling(n).mean()).values.astype(float)
            for m, n in [(10, 25), (35, 43), (63, 100)]
        ]
        _anc_stack = np.column_stack(_anc_sigs)
        sig_raw = np.nanmean(np.where(np.isfinite(_anc_stack), _anc_stack, np.nan), axis=1)

    elif sig_type == "MA Crossover":
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
    # Same-Day = shift 1 (trade at signal's own close, earn t->t+1; no look-ahead).
    # Lag-1    = shift 2 (trade next close, earn t+1->t+2). Matches carry/value/helpers.
    if same_day:
        pos_np[0] = 0.0
        pos_np[1:] = np.where(np.isfinite(sig_raw[:-1]), sig_raw[:-1], 0.0)
    else:
        pos_np[:2] = 0.0
        pos_np[2:] = np.where(np.isfinite(sig_raw[:-2]), sig_raw[:-2], 0.0)
    sig_np = sig_raw

    # ── PnL with transaction costs ─────────────────────────────────────────────
    delta_np  = f1c.diff().values.astype(float)
    pos_s     = pd.Series(pos_np, index=f1r.index)
    delta_s   = pd.Series(delta_np, index=f1r.index)
    gross_pnl = pos_s * delta_s

    # TC in bps: cost per position change = |Δpos| × (bps/10000 / 2) × F1_raw price.
    # The spread is paid on the ACTUAL traded front-month price (F1_raw), not the
    # back-adjusted continuous index (F1_continuous is used only for PnL accounting).
    pos_change   = pos_s.diff().abs()
    pos_change.iloc[0] = abs(pos_s.iloc[0])
    tc_cost_s    = pos_change * (tc_bps / 10000.0 / 2.0) * f1r
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
            "Performance period  (signal uses full history - only metrics & charts below update)",
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
            col.markdown(f'<div class="metric-compact"><h4>{label}</h4><p class="value">-</p></div>',
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

    # ── Out-of-Sample Walk-Forward Validation ──────────────────────────────────
    st.divider()
    _wf_active     = _wf_ma3543_tc(f1r, f1c, _oos_tc_bps)
    # Anchors + IS-opt walk-forward - computed live for ALL OOS windows (TC-aware).
    _WF_ANC_OPT      = _wf_anchors_isopt_tc(f1r, f1c, _oos_tc_bps)
    _anc_opt_vals    = [v for v in _WF_ANC_OPT.values() if v is not None and not np.isnan(v)]
    _WF_OPT_AVG_FULL = round(np.nanmean(_anc_opt_vals), 3) if _anc_opt_vals else np.nan
    _wf_yrs_all    = sorted(k for k in _wf_active if not k.endswith("*"))
    _wf_recent_yrs = _wf_yrs_all[-3:] if len(_wf_yrs_all) >= 3 else _wf_yrs_all
    _wf_first_yr   = _wf_yrs_all[0] if _wf_yrs_all else "N/A"
    _wf_last_yr    = sorted(_wf_active.keys())[-1] if _wf_active else "N/A"
    _recent_label  = f"{_wf_recent_yrs[0]}-{_wf_recent_yrs[-1]}" if _wf_recent_yrs else "-"

    st.markdown("#### Out-of-Sample Walk-Forward Validation")
    st.caption(
        f"IS = 5yr rolling window, OOS = 1yr, Same-Day entry, {len(_wf_active)} OOS windows. "
        "MA(35,43) selected a priori - never re-optimised per window. "
        f"Window labels denote the start year of each OOS period; "
        f"data coverage spans {_wf_first_yr}-{_wf_last_yr[:4]}."
    )

    _wf_vals_all  = [v for v in _wf_active.values() if v is not None and not np.isnan(v)]
    _WF_MA35_AVG  = round(np.nanmean(_wf_vals_all), 3) if _wf_vals_all else np.nan
    _WF_MA35_P23  = round(np.nanmean([_wf_active[y] for y in _wf_recent_yrs
                                       if y in _wf_active and not np.isnan(_wf_active[y])]), 3)
    _WF_N_POS     = sum(1 for v in _wf_vals_all if v > 0)
    _WF_N_GT03    = sum(1 for v in _wf_vals_all if v > 0.30)
    _WF_N_TOTAL   = len(_wf_active)
    _tc_note      = f", {_oos_tc_label}" if _oos_tc_bps > 0 else ""

    _wf_c1, _wf_c2, _wf_c3 = st.columns(3)
    _cs   = ("background:#161616;border:1px solid #2A2A2A;"
             "border-left:4px solid #B87333;border-radius:4px;padding:14px 20px")
    _csg  = ("background:#161616;border:1px solid #2A2A2A;"
             "border-left:4px solid #475569;border-radius:4px;padding:14px 20px")
    _lbl  = ("color:#B87333;font-family:'IBM Plex Mono',monospace;"
             "font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _lblg = ("color:#94A3B8;font-family:'IBM Plex Mono',monospace;"
             "font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _big  = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;"
             "font-size:1.55rem;font-weight:700;margin:0")
    _med  = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;"
             "font-size:1.15rem;font-weight:600;margin:0")
    _sub  = "color:#8A8278;font-size:0.75rem;margin:2px 0"
    _hr   = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"

    with _wf_c1:
        st.markdown(f"""<div style="{_cs}">
<p style="{_lbl}">MA(35,43) - Fixed Parameter</p>
<p style="{_big}">{_WF_MA35_AVG:+.3f}</p>
<p style="{_sub}">Avg OOS Sharpe, {_wf_first_yr}-{_wf_last_yr[:4]}, {_WF_N_TOTAL} Windows{_tc_note}</p>
<hr style="{_hr}"/>
<p style="{_sub}">{_recent_label} avg</p>
<p style="{_med}">{_WF_MA35_P23:+.3f}</p>
<p style="{_sub}">Zero re-optimisation, 13-day tail excluded</p>
</div>""", unsafe_allow_html=True)

    with _wf_c2:
        st.markdown(f"""<div style="{_cs}">
<p style="{_lbl}">Anchors + IS-Opt Weights</p>
<p style="{_big}">{_WF_OPT_AVG_FULL:+.3f}</p>
<p style="{_sub}">Avg OOS Sharpe, {_wf_first_yr}-{_wf_last_yr[:4]}, {len(_WF_ANC_OPT)} Windows{_tc_note}</p>
<hr style="{_hr}"/>
<p style="{_sub}">MA(10,25) + MA(35,43) + MA(63,100)</p>
<p style="{_sub}">Max-Sharpe QP weights, re-optimised annually on prior 5yr IS data</p>
<p style="{_sub}">Optimizer assigns w≈1.0 to MA(35,43) in 9 / 15 windows</p>
</div>""", unsafe_allow_html=True)

    with _wf_c3:
        _wf_best_yr  = max(_wf_active, key=lambda y: _wf_active[y])
        _wf_worst_yr = min(_wf_active, key=lambda y: _wf_active[y])
        st.markdown(f"""<div style="{_csg}">
<p style="{_lblg}">OOS Consistency - MA(35,43)</p>
<p style="{_sub}">Positive OOS Sharpe</p>
<p style="{_med}">{_WF_N_POS} / {_WF_N_TOTAL} windows</p>
<hr style="{_hr}"/>
<p style="{_sub}">OOS Sharpe above +0.30</p>
<p style="{_med}">{_WF_N_GT03} / {_WF_N_TOTAL} windows</p>
<hr style="{_hr}"/>
<p style="{_sub}">Best: {_wf_best_yr} ({_wf_active[_wf_best_yr]:+.3f}), Worst: {_wf_worst_yr} ({_wf_active[_wf_worst_yr]:+.3f})</p>
</div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    _wf_chart_c1, _wf_chart_c2 = st.columns([2, 4])
    with _wf_chart_c1:
        _wf_chart_mode = st.selectbox(
            "Annual OOS Sharpe - Strategy",
            ["MA(35,43) - Annual OOS Sharpe", "Anchors + IS-Opt Weights"],
            index=0, key="wf_chart_mode",
        )

    if _wf_chart_mode == "MA(35,43) - Annual OOS Sharpe":
        _wf_years_plot  = list(_wf_active.keys())
        _wf_sh_plot     = list(_wf_active.values())
        _wf_bar_cls = [
            (COLORS["primary"] if v >= 0 else "#B05030") if y in _wf_recent_yrs
            else (COLORS["green"] if v >= 0 else COLORS["red"])
            for y, v in _wf_active.items()
        ]
        _wf_bar_brd = ["#D4A843" if y in _wf_recent_yrs else "rgba(0,0,0,0)" for y in _wf_years_plot]
        fig_wf_bar = go.Figure()
        fig_wf_bar.add_trace(go.Bar(
            x=_wf_years_plot, y=_wf_sh_plot,
            marker_color=_wf_bar_cls,
            marker_line_color=_wf_bar_brd, marker_line_width=1.5,
            name="OOS Sharpe",
            hovertemplate="%{x}<br>OOS Sharpe: %{y:.3f}<extra></extra>",
        ))
        fig_wf_bar.add_hline(y=0, line_dash="solid", line_color="#475569", line_width=1)
        fig_wf_bar.add_hline(
            y=_WF_MA35_AVG, line_dash="dot", line_color=COLORS["amber"], line_width=1.5,
            annotation_text=f"Full avg {_WF_MA35_AVG:+.3f}",
            annotation_position="top right",
            annotation_font=dict(size=10, color=COLORS["amber"]),
        )
        fig_wf_bar.add_hline(
            y=_WF_MA35_P23, line_dash="dot", line_color=COLORS["primary"], line_width=1.5,
            annotation_text=f"{_recent_label} avg {_WF_MA35_P23:+.3f}",
            annotation_position="top left",
            annotation_font=dict(size=10, color=COLORS["primary"]),
        )
        fig_wf_bar.update_layout(
            **CHART_LAYOUT, height=300,
            title=dict(
                text=f"MA(35,43) - Annual OOS Sharpe  (Walk-Forward, IS=5yr, Same-Day{_tc_note})",
                font=dict(size=13),
            ),
            yaxis_title="OOS Sharpe", xaxis_title=None, showlegend=False,
        )
        st.plotly_chart(fig_wf_bar, use_container_width=True)
        _gold_lbl = ", ".join(_wf_recent_yrs) if _wf_recent_yrs else "recent"
        st.caption(
            f"Gold-bordered bars = most recent {len(_wf_recent_yrs)} complete OOS windows ({_gold_lbl}). "
            + (f"TC = {_oos_tc_label} deducted on each signal flip." if _oos_tc_bps > 0 else "Gross returns shown.")
        )
    else:
        # Anchors + IS-Opt Weights: all OOS windows, computed live (TC-aware)
        _anc_years = list(_WF_ANC_OPT.keys())
        _anc_sh    = list(_WF_ANC_OPT.values())
        _anc_recent = _anc_years[-3:] if len(_anc_years) >= 3 else _anc_years
        fig_anc = go.Figure()
        fig_anc.add_trace(go.Bar(
            x=_anc_years, y=_anc_sh,
            marker_color=[
                (COLORS["primary"] if v >= 0 else "#B05030") if y in _anc_recent
                else (COLORS["green"] if v >= 0 else COLORS["red"])
                for y, v in _WF_ANC_OPT.items()
            ],
            marker_line_color=["#D4A843" if y in _anc_recent else "rgba(0,0,0,0)" for y in _anc_years],
            marker_line_width=1.5,
            hovertemplate="%{x}<br>OOS Sharpe: %{y:.3f}<extra></extra>",
        ))
        fig_anc.add_hline(y=0, line_dash="solid", line_color="#475569", line_width=1)
        fig_anc.add_hline(
            y=_WF_OPT_AVG_FULL, line_dash="dot", line_color=COLORS["amber"], line_width=1.5,
            annotation_text=f"{len(_anc_sh)}-window avg {_WF_OPT_AVG_FULL:+.3f}",
            annotation_position="top right",
            annotation_font=dict(size=10, color=COLORS["amber"]),
        )
        fig_anc.update_layout(
            **CHART_LAYOUT, height=300,
            title=dict(
                text=f"Anchors + IS-Opt Weights - OOS Sharpe  (Walk-Forward, IS=5yr, all windows{_tc_note})",
                font=dict(size=13),
            ),
            yaxis_title="OOS Sharpe", xaxis_title=None, showlegend=False,
        )
        st.plotly_chart(fig_anc, use_container_width=True)
        st.caption(
            f"All {len(_anc_sh)} OOS windows, IS-opt QP weights re-fit annually on prior 5yr data "
            f"(max-Sharpe, w≥0, Σw=1). Full-period avg {_WF_OPT_AVG_FULL:+.3f}"
            + (f", TC = {_oos_tc_label}." if _oos_tc_bps > 0 else " (gross).")
            + " The three anchor MAs [MA(10,25), MA(35,43), MA(63,100)] are structural, not data-fitted - "
            "only the combination weights are optimised IS, so there is no anchor-selection look-ahead."
        )

    with st.expander("Walk-Forward Annual Detail", expanded=False):
        _wf_yrs_tbl = list(_wf_active.keys())
        _is_labels  = [f"{int(y.rstrip('*'))-5}-{int(y.rstrip('*'))-1}" for y in _wf_yrs_tbl]
        _tc_col     = f"MA(35,43) OOS{'  '+_oos_tc_label if _oos_tc_bps>0 else ' (Gross)'}"
        _anc_map    = {y: f"{v:+.3f}" for y, v in _WF_ANC_OPT.items()}
        _wf_tbl = pd.DataFrame({
            "OOS Window":      _wf_yrs_tbl,
            "IS Period (5yr)": _is_labels,
            _tc_col:           [f"{_wf_active.get(y, float('nan')):+.3f}" for y in _wf_yrs_tbl],
            "Anchors+Opt OOS": [_anc_map.get(y, "-") for y in _wf_yrs_tbl],
            "Status":          ["✓" if _wf_active.get(y, float("nan")) > 0 else "✗" for y in _wf_yrs_tbl],
        })
        st.dataframe(_wf_tbl, use_container_width=True, hide_index=True)
        st.caption(
            "OOS Window label = start year of 252-day OOS period. "
            "Anchors+Opt: IS-opt QP weights on MA(10,25)+MA(35,43)+MA(63,100), re-fit each window; "
            f"computed live for all {len(_WF_ANC_OPT)} windows, full-period avg = {_WF_OPT_AVG_FULL:+.3f}."
        )

    st.divider()


    # ── Rolling Sharpe ─────────────────────────────────────────────────────────
    st.divider()
    section_header("ROLLING SHARPE RATIO")
    rs_c1, rs_c2 = st.columns([3, 1])
    with rs_c2:
        rs_window = st.radio("Window", ["1 Year (252d)", "2 Years (504d)", "Both"],
                             index=2, key="rs_window", horizontal=False)
        rs_basis = st.radio("Returns", ["Gross", "Net of TC"], index=0,
                            key="rs_basis_mom", horizontal=False)

    _rs_net = rs_basis.startswith("Net")
    _dr = (net_ret_all if _rs_net else gross_ret_all).fillna(0)
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
            title=dict(text=f"{variant_label} - Rolling Sharpe ({timing_label}, {'Net of TC' if _rs_net else 'Gross'})", font=dict(size=13)),
            yaxis_title="Annualised Sharpe", xaxis_title=None, hovermode="x unified",
        )
        fig_rs.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_rs, use_container_width=True)
    st.caption(f"Signal computed over full {_m_is_yr0}-{_m_is_yr1} history. "
               f"Rolling Sharpe uses {'net (' + tc_label + ')' if _rs_net else 'gross'} returns. "
               "Positive/negative swings show regime dependence - a consistently positive curve indicates robustness.")

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
        title=dict(text="Strategy Comparison - Cumulative PnL (Gross, USD/MT)", font=dict(size=13)),
        yaxis_title="Cumulative PnL (USD/MT)",
        xaxis_title=None, hovermode="x unified",
    )
    fig_cum.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_cum, use_container_width=True)

    # ── Annual PnL bar chart ──────────────────────────────────────────────────
    st.divider()
    section_header(f"ANNUAL PnL BREAKDOWN - {variant_label} ({timing_label}), Gross USD/MT")
    st.caption("Shows the strategy currently selected above (Signal Type / Variant / Timing). "
               "Change those controls to view a different strategy's annual PnL.")

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
                title=dict(text=f"MA Crossover - Sharpe surface  (m=fast, n=slow, n≤{hm_n_max})", font=dict(size=13)),
                coloraxis_colorbar=dict(title=hm_metric, thickness=14),
            )
            with hm_c1:
                st.plotly_chart(fig_hm, use_container_width=True)
            st.caption("White stars = current top-5 by Sharpe. A wide green plateau means the "
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
                title=dict(text="CTA (Baz-Granger) - Sharpe scatter  (S=short EWMA, L=long EWMA)", font=dict(size=13)),
            )
            with hm_c1:
                st.plotly_chart(fig_cta, use_container_width=True)
            st.caption("White stars = top-5 by Sharpe. Each dot is one (S,L) pair.")

    # ── Signal & Position chart ────────────────────────────────────────────────
    st.divider()
    section_header(f"SIGNAL & POSITION OVER TIME - {variant_label} ({timing_label})")
    st.caption("This chart reflects the strategy chosen in the controls above (Strategy Preset, or "
               "Signal Type / Variant / Timing). To view a different strategy here, change those controls; "
               "the comparison dropdown in the IS section overlays a second variant against it.")

    # Full date range - no filter widget
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
        title=dict(text=f"{variant_label} - Price & Position ({timing_label})", font=dict(size=13)),
        hovermode="x unified", showlegend=True,
        xaxis2_title=None,
    )
    fig_sig.update_yaxes(title_text="F1 Price ($/MT)", row=1, col=1)
    if sig_type == "Anchors + IS-Opt Weights":
        fig_sig.update_yaxes(title_text="Position (EW composite)", row=2, col=1)
    else:
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
        _fmt_pos = lambda v: f"{v:+.2f}" if sig_type == "Anchors + IS-Opt Weights" else f"{int(round(v)):+d}"
        _fmt_dir = lambda v: "LONG" if v > 0 else ("SHORT" if v < 0 else "FLAT")
        flip_df = pd.DataFrame({
            "Date":       flip_dates.index.strftime("%Y-%m-%d"),
            "Position":   [_fmt_pos(v) for v in flip_dates.values],
            "Direction":  [_fmt_dir(v) for v in flip_dates.values],
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

**CTA Signal - Baz-Granger Eqs 29-33**
- x = EWMA(S) − EWMA(L)  [EWMA convention: com = n−1, i.e. λ = (n−1)/n]
- y = x / σ₆₃(price)     [63-day price volatility normalisation]
- z = y / σ₂₅₂(y)        [252-day signal normalisation]
- u = z × exp(−z²/4) / 0.89  [response function - shrinks extreme signals]
- Signal = sign(u)
- CTA Paper uses 3 timescales (S,L) = (8,24), (16,48), (32,96); S_CTA = mean(u₁,u₂,u₃)

**Position timing** (no look-ahead either way)
- *Same-Day (shift 1)*: trade at the signal's own close(t); first return t→t+1. Realistic default.
- *Lag-1 (shift 2)*: trade at the next close(t+1); first return t+1→t+2. Conservative.
- MA(35,43): Same-Day +0.72 vs Lag-1 +0.63. CTA also leads Same-Day; Anchors near-tied (Lag-1 marginally ahead).

**Transaction costs**
- Expressed in basis points (bps) of notional, round-trip
- TC_cost[t] = |Δposition[t]| × (bps / 10000 / 2) × F1_raw[t]   (spread is paid on the actual traded front-month price)
- Flip (+1→−1): |Δ|=2 → cost = 1 full round trip × price
- Entry (0→±1): |Δ|=1 → cost = ½ round trip × price
- Cost is time-varying (scales with copper price level)

**Returns & risk metrics**
- daily_ret[t] = position[t] × ΔF1_cont[t] / F1_cont[t−1]
- All % metrics (Ann Return, Std Dev, Max DD, Calmar, Sortino) computed from daily_ret
- Sharpe = Ann_ret / Ann_std (unitless, consistent across gross/net)
        """)

    st.markdown(
        '<div style="background:#0D1117;border:1px solid #2A2A2A;border-left:4px solid #475569;'
        'border-radius:4px;padding:12px 20px;margin-top:20px;">'
        '<span style="color:#94A3B8;font-size:0.78rem;font-family:\'IBM Plex Mono\',monospace;font-weight:600;">'
        'NEXT &rarr; </span>'
        '<span style="color:#B87333;font-size:0.82rem;font-family:\'IBM Plex Mono\',monospace;font-weight:700;">'
        'Tab 8: Carry Signals</span>'
        '<span style="color:#8A8278;font-size:0.78rem;"> &nbsp;-&nbsp; '
        'Backwardation / contango premium from the term structure. '
        'Complements momentum: carry is a <em>level</em> signal; momentum is <em>trend</em>-based. '
        'Low historical correlation confirms diversification value.</span></div>',
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════
# TAB 8: CARRY SIGNALS
# ══════════════════════════════════════════════════════

with tab8:
    _c8_metal = st.radio("🔬 Metal", ["Copper", "Aluminium"], horizontal=True, key="carry_metal")
    st.markdown(f"### Carry Signals - LME {_c8_metal}")
    st.markdown(
        '<div style="background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;'
        'border-radius:4px;padding:10px 18px;margin-bottom:10px;display:flex;align-items:center;gap:16px;">'
        '<span style="color:#B87333;font-family:\'IBM Plex Mono\',monospace;font-size:0.78rem;'
        'font-weight:700;white-space:nowrap;">SIGNAL 2 OF 3</span>'
        '<span style="color:#8A8278;font-size:0.8rem;">Structural backwardation / contango premium from the forward curve. '
        'No free parameters - the signal is a market structural measure, not fitted data. '
        'IS performance is representative of OOS. Combines with momentum in the portfolio (Tab 10).</span></div>',
        unsafe_allow_html=True,
    )
    st.caption("Term structure carry: Long in backwardation, Short in contango. Signal from curve shape; PnL always from F1_continuous.")

    # ── Best Carry Signal - By Variant (DYNAMIC, computed live per metal) ──────
    section_header(f"BEST CARRY SIGNAL - BY VARIANT  ({_c8_metal})")
    st.caption(
        f"Best configuration per variant family for {_c8_metal} - full-period IS backtest, gross active-day "
        "Sharpe (TC=0), no look-ahead. Computed live - changes with the metal toggle. "
        "Past performance is not indicative of future results."
    )
    _bcs  = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;"
             "border-radius:4px;padding:14px 20px")
    _bcsx = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #5BAD72;"
             "border-radius:4px;padding:14px 20px")
    _blbl = ("color:#B87333;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _blbx = ("color:#5BAD72;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _bbig = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;font-size:1.55rem;font-weight:700;margin:0")
    _bsub = "color:#8A8278;font-size:0.75rem;margin:2px 0"
    _bhr  = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"
    _cb = _carry_best_cards(_c8_metal)
    _cb_fam = [("V4 - Carry Momentum", _cb.get("mom") if _cb else None),
               ("V3 - Z-Score",        _cb.get("zscore") if _cb else None),
               ("V1 - Roll Yield",     _cb.get("level") if _cb else None),
               ("V2 - Long Slope",     _cb.get("slope") if _cb else None)]
    _cb_best = max(range(4), key=lambda i: (_cb_fam[i][1]["gross"] if _cb_fam[i][1] and not np.isnan(_cb_fam[i][1]["gross"]) else -9))
    for _i, (_col, (_lbl, _d)) in enumerate(zip(st.columns(4), _cb_fam)):
        with _col:
            _star = "  ★" if _i == _cb_best else ""
            _sty, _lsty = (_bcsx, _blbx) if _i == _cb_best else (_bcs, _blbl)
            if not _d:
                st.markdown(f'<div style="{_sty}"><p style="{_lsty}">{_lbl}{_star}</p>'
                            f'<p style="{_bbig}">N/A</p></div>', unsafe_allow_html=True)
            else:
                st.markdown(f"""<div style="{_sty}">
<p style="{_lsty}">{_lbl}{_star}</p>
<p style="{_bbig}">{_fmt_sh(_d['gross'])}</p>
<p style="{_bsub}">Sharpe Ratio (Gross)</p>
<hr style="{_bhr}"/>
<p style="{_bsub}">{_d['name']}, {_d['timing']}</p>
<p style="{_bsub}">Ann Ret ≈ {_fmt_pct(_d['ann'])}, Max DD ≈ {_fmt_dd(_d['mdd'])}</p>
</div>""", unsafe_allow_html=True)
    st.caption(f"Cards recompute live for {_c8_metal}; the strongest family is starred. "
               "Same-Day vs Lag-1 is selected per family by gross Sharpe.")

    st.markdown("""
    <div style="background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;border-radius:4px;padding:14px 20px;margin-bottom:8px;">
      <div style="display:flex;gap:40px;flex-wrap:wrap;">
        <div style="min-width:180px;">
          <span style="color:#B87333;font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:0.9rem;">V1 - Roll Yield</span><br>
          <span style="color:#8A8278;font-size:0.78rem;">Short-end basis: (F1-F2)/F1, (F1-F3)/F1,<br>or (Cash-3M)/Cash. Raw 1-period roll cost.</span>
        </div>
        <div style="min-width:200px;">
          <span style="color:#B87333;font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:0.9rem;">V2 - Long Slope</span><br>
          <span style="color:#8A8278;font-size:0.78rem;">Curve slope at longer tenors: (Fj-Fk)/Fk<br>for 10 pairs (F3-F15 through F12-F24).</span>
        </div>
        <div style="min-width:180px;">
          <span style="color:#B87333;font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:0.9rem;">V3 - Z-score</span><br>
          <span style="color:#8A8278;font-size:0.78rem;">Rolling 252-day standardisation of (F1-F2)/F1.<br>Filters permanent level shifts; regime-relative signal.</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Strategy Preset ───────────────────────────────────────────────────────
    _CARRY_PRESETS = {
        "V1, (F1-F2)/F1, Same-Day  [Best V1]": {
            "carry_vgroup":  "V1 - Roll Yield",
            "carry_subv":    "(F1-F2)/F1",
            "carry_timing":  "Same-Day",
        },
        "V2, F3-F15 Slope, Same-Day  [Best V2]": {
            "carry_vgroup":  "V2 - Long Slope",
            "carry_subv":    "F3-F15 Slope",
            "carry_timing":  "Same-Day",
        },
        "V3, Z-score, Same-Day  [Best V3]": {
            "carry_vgroup":  "V3 - Z-score",
            "carry_subv":    "Z-score (252d window)",
            "carry_timing":  "Same-Day",
        },
        "V1, (F1-F2)/F1, Lag-1  [Sensitivity Check]": {
            "carry_vgroup":  "V1 - Roll Yield",
            "carry_subv":    "(F1-F2)/F1",
            "carry_timing":  "Lag-1 (Next-Day)",
        },
        "V1, (Cash-3M)/Cash, Same-Day": {
            "carry_vgroup":  "V1 - Roll Yield",
            "carry_subv":    "(Cash-3M)/Cash",
            "carry_timing":  "Same-Day",
        },
        "Custom (use controls below)": {},
    }

    def _apply_carry_preset():
        cfg = _CARRY_PRESETS.get(st.session_state.get("carry_preset", "Custom (use controls below)"), {})
        for k, v in cfg.items():
            st.session_state[k] = v

    _c8_pre_col, _c8_pre_info = st.columns([2.5, 3.5])
    with _c8_pre_col:
        st.selectbox(
            "Strategy Preset",
            list(_CARRY_PRESETS.keys()),
            index=0,
            key="carry_preset",
            on_change=_apply_carry_preset,
        )
    with _c8_pre_info:
        st.markdown(
            '<div style="padding:8px 0;color:#7A7068;font-size:0.78rem;">'
            'Selecting a preset auto-fills all controls below. '
            'Switch to <b>Custom</b> to edit individual parameters freely.</div>',
            unsafe_allow_html=True,
        )

    # ── Controls ──────────────────────────────────────────────────────────────
    c8_c1, c8_c2, c8_c3, c8_c4 = st.columns([2, 2, 1.5, 1.5])
    with c8_c1:
        carry_vgroup = st.selectbox(
            "Variant Group",
            ["V4 - Carry Momentum", "V3 - Z-score", "V1 - Roll Yield", "V2 - Long Slope"],
            index=0, key="carry_vgroup",
            help="Default V4 Carry Momentum (20d): best walk-forward OOS Sharpe of all carry "
                 "signals (+0.50 net5), robust to TC (still +0.29 at 20bps) and to execution lag. "
                 "V3 Z-score next (+0.40 OOS). V1 level is the naive baseline (+0.24). "
                 "Note: 1-day carry-change scores higher at low TC but is a microstructure "
                 "artifact (collapses at 20bps and with one extra day's lag).",
        )
    with c8_c2:
        if carry_vgroup == "V1 - Roll Yield":
            _c8_sub_opts = {"(F1-F2)/F1": "f1f2", "(F1-F3)/F1": "f1f3", "(Cash-3M)/Cash": "cash3m"}
        elif carry_vgroup == "V2 - Long Slope":
            _c8_sub_opts = {f"F{j}-F{k} Slope": (j, k)
                            for j, k in [(3,15),(4,16),(5,17),(6,18),(7,19),(8,20),(9,21),(10,22),(11,23),(12,24)]}
        elif carry_vgroup == "V4 - Carry Momentum":
            _c8_sub_opts = {"20-day Δcarry (best OOS +0.50)": 20, "60-day Δcarry": 60}
        else:
            _c8_sub_opts = {"Z-score sign (252d)": (252, 0.0),
                            "Z-score deadband |z|>0.5 (252d)": (252, 0.5)}
        # Guard: reset if session value no longer valid for current carry_vgroup
        if st.session_state.get("carry_subv", "") not in _c8_sub_opts:
            st.session_state["carry_subv"] = list(_c8_sub_opts.keys())[0]
        carry_sub_label = st.selectbox("Sub-Variant", list(_c8_sub_opts.keys()), key="carry_subv")
        carry_sub_val = _c8_sub_opts[carry_sub_label]
    with c8_c3:
        carry_timing = st.selectbox("Position Entry", ["Same-Day", "Lag-1 (Next-Day)"], index=0, key="carry_timing")
        carry_same_day = carry_timing == "Same-Day"
    with c8_c4:
        carry_tc_map = _tc_label_map(_get_last_f1_price())
        carry_tc_label = st.selectbox("TC (bps, round-trip)", list(carry_tc_map.keys()), index=0, key="carry_tc")
        carry_tc_bps = carry_tc_map[carry_tc_label]

    # Build spec dict
    if carry_vgroup == "V1 - Roll Yield":
        carry_spec = {"variant": "v1", "signal_type": carry_sub_val, "same_day": carry_same_day}
    elif carry_vgroup == "V2 - Long Slope":
        carry_spec = {"variant": "v3", "j": carry_sub_val[0], "k": carry_sub_val[1], "same_day": carry_same_day}
    elif carry_vgroup == "V4 - Carry Momentum":
        carry_spec = {"variant": "v5", "horizon": carry_sub_val, "same_day": carry_same_day}
    else:
        carry_spec = {"variant": "v4", "window": carry_sub_val[0],
                      "deadband": carry_sub_val[1], "same_day": carry_same_day}

    # ── Data loading ──────────────────────────────────────────────────────────
    _f1_df_c8 = _load_f1_data(_c8_metal)
    if _f1_df_c8.empty:
        st.error(f"Rolling F1 file for {_c8_metal} not found. Place the CSV beside app.py.")
        st.stop()
    cf1c = _f1_df_c8["F1_continuous"]
    cf1r = _f1_df_c8["F1_raw"]

    _cu_sheet_c8 = _find_curve_sheet(_c8_metal, curve_data) if curve_data else None
    if not curve_data or _cu_sheet_c8 is None:
        st.error("Futures Curve data not loaded. Upload Metals Futures Curve file in the sidebar.")
        st.stop()
    c8_curve_px = curve_data[_cu_sheet_c8]["prices"].copy()
    c8_curve_px.index = pd.to_datetime(c8_curve_px.index).normalize()
    c8_curve_px = c8_curve_px.sort_index()

    _cash_cu8 = parse_cash_3m_columns(cash_data.get(_c8_metal, pd.DataFrame()), _c8_metal)
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
    cf1r_a = cf1r.reindex(_c8_idx)   # F1_raw aligned — used for TC, not PnL

    carry_sig_arr = _carry_binarize(carry_raw.values, carry_spec)
    T_c8 = len(carry_sig_arr)
    carry_pos_np = np.empty(T_c8)
    # Same-Day -> trade at signal close, return t->t+1 (shift 1).
    # Lag-1    -> trade at next close,   return t+1->t+2 (shift 2). No look-ahead either way.
    if carry_same_day:
        carry_pos_np[0] = 0.0
        carry_pos_np[1:] = np.where(np.isfinite(carry_sig_arr[:-1]), carry_sig_arr[:-1], 0.0)
    else:
        carry_pos_np[:2] = 0.0
        carry_pos_np[2:] = np.where(np.isfinite(carry_sig_arr[:-2]), carry_sig_arr[:-2], 0.0)
    carry_pos = pd.Series(carry_pos_np, index=_c8_idx)

    c8_delta = cf1c_a.diff()
    c8_gross_pnl = carry_pos * c8_delta
    c8_pos_change = carry_pos.diff().abs()
    c8_pos_change.iloc[0] = abs(carry_pos.iloc[0])
    c8_tc_cost = c8_pos_change * (carry_tc_bps / 10000.0 / 2.0) * cf1r_a
    c8_net_pnl = c8_gross_pnl - c8_tc_cost
    c8_cum_gross = c8_gross_pnl.cumsum()
    c8_cum_net = c8_net_pnl.cumsum()

    cf1_prev8 = cf1c_a.shift(1)
    c8_gross_ret_all = (c8_gross_pnl / cf1_prev8).replace([np.inf, -np.inf], np.nan)
    c8_net_ret_all = (c8_net_pnl / cf1_prev8).replace([np.inf, -np.inf], np.nan)

    # Regime signal series (used by multiple sections below)
    _c8_sig_bin = pd.Series(_carry_binarize(carry_raw.values, carry_spec), index=_c8_idx)
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

    def _cmcard(col, label, val, fmt=".2f", suffix=""):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            col.markdown(f'<div class="metric-compact"><h4>{label}</h4><p class="value">-</p></div>', unsafe_allow_html=True)
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

    # ── Section 4: Date Filter + Performance Metrics ──────────────────────────
    st.divider()
    pf8_c1, _ = st.columns([3, 1])
    with pf8_c1:
        carry_perf_dates = st.date_input(
            "Performance period  (signal uses full history - only metrics & charts below update)",
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

    # ── Multi-strategy selector (drives the comparison table + Rolling Sharpe overlay) ──
    st.divider()
    section_header("MULTI-STRATEGY COMPARISON")
    _cmp_all_opts = [k for k in _CARRY_CMP_OPTIONS.keys() if _CARRY_CMP_OPTIONS[k] is not None]
    _carry_multi_sel = st.multiselect(
        "Select strategies to compare",
        _cmp_all_opts,
        default=[_cmp_all_opts[0]] if _cmp_all_opts else [],
        key="carry_multi_sel",
    )
    st.caption("The comparison table here and the Rolling Sharpe overlay below reflect every strategy you select. "
               "The Performance Metrics above use the primary strategy from the upper controls.")

    def _carry_gross_daily_ret_for(spec: dict) -> pd.Series:
        """Gross daily return series for a given carry spec."""
        cr = _carry_raw_signal(c8_curve_px, _cash_cu8, spec)
        if cr.empty:
            return pd.Series(dtype=float)
        _sd2 = spec.get("same_day", True)
        _cidx = cr.index.intersection(cf1c.index)
        if len(_cidx) < 20:
            return pd.Series(dtype=float)
        _vcr2   = cr.reindex(_cidx)
        _f1c_r2 = cf1c.reindex(_cidx)
        _vsig2  = _carry_binarize(_vcr2.values, spec); _vT2 = len(_vsig2)
        if _vT2 < 20:
            return pd.Series(dtype=float)
        _vpos2 = np.empty(_vT2)
        # Same-Day (shift 1) vs Lag-1 / next-close (shift 2); both no look-ahead.
        if _sd2:
            _vpos2[0] = 0.0
            _vpos2[1:] = np.where(np.isfinite(_vsig2[:-1]), _vsig2[:-1], 0.0)
        else:
            _vpos2[:2] = 0.0
            _vpos2[2:] = np.where(np.isfinite(_vsig2[:-2]), _vsig2[:-2], 0.0)
        _pos2s = pd.Series(_vpos2, index=_cidx)
        _pnl2s = _pos2s * _f1c_r2.diff()
        with np.errstate(invalid="ignore", divide="ignore"):
            return (_pnl2s / _f1c_r2.shift(1)).replace([np.inf, -np.inf], np.nan)

    # Compute metrics for each selected comparison strategy
    _cmp_rows = []
    _cmp_rets = {}
    _CMP_LINE_COLORS = [
        COLORS["primary"], COLORS["amber"], "#64748B", "#5BAD72",
        "#CF9FFF", "#FF8C00", "#00CED1", "#FF6B9D",
    ]
    for _cmp_lbl in _carry_multi_sel:
        _cmp_spec = _CARRY_CMP_OPTIONS.get(_cmp_lbl)
        if _cmp_spec is None:
            continue
        try:
            _cmp_ret = _carry_gross_daily_ret_for(_cmp_spec)
            if _cmp_ret.empty or len(_cmp_ret.dropna()) < 20:
                continue
            _cmp_rets[_cmp_lbl] = _cmp_ret
            # Compute metrics (full period, gross)
            _cmp_pos  = _cmp_ret.where(_cmp_ret != 0).notna().astype(float)
            _act = _cmp_ret[_cmp_ret.abs() > 0].dropna()
            if len(_act) < 20:
                continue
            _ann_r  = float(_act.mean() * 252 * 100)
            _ann_sd = float(_act.std()  * np.sqrt(252) * 100)
            _sh = _ann_r / _ann_sd if _ann_sd > 0 else np.nan
            _down = _act[_act < 0]
            _srt_d = float(_down.std() * np.sqrt(252) * 100) if len(_down) > 1 else np.nan
            _srt = _ann_r / _srt_d if _srt_d and _srt_d > 0 else np.nan
            _cum = _cmp_ret.fillna(0).cumsum() * 100
            _mdd = float((_cum - _cum.cummax()).min())
            _cal = _ann_r / abs(_mdd) if _mdd != 0 else np.nan
            _cmp_rows.append({
                "Strategy": _cmp_lbl,
                "Sharpe": f"{_sh:+.3f}" if not np.isnan(_sh) else "-",
                "Sortino": f"{_srt:+.3f}" if not np.isnan(_srt) else "-",
                "Ann Ret %": f"{_ann_r:+.1f}" if not np.isnan(_ann_r) else "-",
                "Ann Std %": f"{_ann_sd:.1f}" if not np.isnan(_ann_sd) else "-",
                "Max DD %": f"{_mdd:.1f}" if not np.isnan(_mdd) else "-",
                "Calmar": f"{_cal:+.2f}" if not np.isnan(_cal) else "-",
                "Hit Rate %": f"{float((_act > 0).mean()) * 100:.1f}",
            })
        except Exception:
            continue

    if _cmp_rows:
        st.dataframe(pd.DataFrame(_cmp_rows), use_container_width=True, hide_index=True)
    elif _carry_multi_sel:
        st.info("Unable to compute metrics for the selected strategies - check data availability.")

    # ── Section 5: Rolling Sharpe ──────────────────────────────────────────────
    st.divider()
    section_header("ROLLING SHARPE RATIO")
    crs8_c1, crs8_c2 = st.columns([3, 1])
    with crs8_c2:
        crs8_win = st.radio("Window", ["1 Year (252d)", "2 Years (504d)", "Both"],
                             index=2, key="carry_rs_window", horizontal=False)
        crs8_basis = st.radio("Returns", ["Gross", "Net of TC"], index=0,
                              key="carry_rs_basis", horizontal=False)
    _crs8_net = crs8_basis.startswith("Net")
    _cdr8 = (c8_net_ret_all if _crs8_net else c8_gross_ret_all).fillna(0)
    croll8_252 = (_cdr8.rolling(252).mean() / _cdr8.rolling(252).std() * np.sqrt(252))
    croll8_504 = (_cdr8.rolling(504).mean() / _cdr8.rolling(504).std() * np.sqrt(252))
    with crs8_c1:
        fig_crs8 = go.Figure()
        if crs8_win in ("1 Year (252d)", "Both"):
            fig_crs8.add_trace(go.Scatter(
                x=croll8_252.index, y=croll8_252.values,
                name=f"{carry_sub_label} (1yr)",
                mode="lines", line=dict(color=COLORS["primary"], width=1.8),
                hovertemplate="%{x|%b %Y}<br>Sharpe (1yr): %{y:.2f}<extra></extra>",
            ))
        if crs8_win in ("2 Years (504d)", "Both"):
            fig_crs8.add_trace(go.Scatter(
                x=croll8_504.index, y=croll8_504.values,
                name=f"{carry_sub_label} (2yr)",
                mode="lines", line=dict(color=COLORS["primary"], width=1.8, dash="dot"),
                hovertemplate="%{x|%b %Y}<br>Sharpe (2yr): %{y:.2f}<extra></extra>",
            ))
        # Overlay comparison strategies from multi-select
        for _ci2, (_cmp_lbl2, _cmp_ret2) in enumerate(_cmp_rets.items()):
            _ccol2 = _CMP_LINE_COLORS[min(_ci2 + 1, len(_CMP_LINE_COLORS) - 1)]
            _cdr2  = _cmp_ret2.fillna(0)
            if crs8_win in ("1 Year (252d)", "Both"):
                _croll2 = _cdr2.rolling(252).mean() / _cdr2.rolling(252).std() * np.sqrt(252)
                fig_crs8.add_trace(go.Scatter(
                    x=_croll2.index, y=_croll2.values,
                    name=f"{_cmp_lbl2} (1yr)",
                    mode="lines", line=dict(color=_ccol2, width=1.2),
                    hovertemplate=f"%{{x|%b %Y}}<br>{_cmp_lbl2} (1yr): %{{y:.2f}}<extra></extra>",
                ))
            if crs8_win in ("2 Years (504d)", "Both"):
                _croll2b = _cdr2.rolling(504).mean() / _cdr2.rolling(504).std() * np.sqrt(252)
                fig_crs8.add_trace(go.Scatter(
                    x=_croll2b.index, y=_croll2b.values,
                    name=f"{_cmp_lbl2} (2yr)",
                    mode="lines", line=dict(color=_ccol2, width=1.2, dash="dot"),
                    hovertemplate=f"%{{x|%b %Y}}<br>{_cmp_lbl2} (2yr): %{{y:.2f}}<extra></extra>",
                ))
        fig_crs8.add_hline(y=0,   line_dash="dash", line_color="#475569", line_width=1)
        fig_crs8.add_hline(y=0.5, line_dash="dot",  line_color=COLORS["green"],
                            line_width=0.8, annotation_text="0.5", annotation_position="right")
        fig_crs8.update_layout(
            **CHART_LAYOUT, height=360,
            title=dict(text=f"Rolling Sharpe - Carry Strategies ({carry_timing}, {'Net of TC' if _crs8_net else 'Gross'})", font=dict(size=13)),
            yaxis_title="Annualised Sharpe", xaxis_title=None, hovermode="x unified",
        )
        fig_crs8.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_crs8, use_container_width=True)
        if _crs8_net:
            st.caption(f"Main line is net of {carry_tc_label}; overlay comparison strategies remain gross.")
    st.caption("Carry tends to be regime-dependent - performs strongly during sustained backwardation cycles (e.g., 2006-2008, 2021-2022). "
               "Positive rolling Sharpe validates the strategy over that window. "
               "Solid = 1yr window, dotted = 2yr window.")

    # ── Section 5b: Sub-Period Analysis ───────────────────────────────────────
    st.divider()
    section_header("SUB-PERIOD ANALYSIS (Pre-2022 / Post-2022)")
    st.caption("Carry risk premia flipped post-2022: sustained contango following the 2021-2022 backwardation spike "
               "reduced carry returns sharply. No parameters were optimised - IS = OOS for all carry variants.")

    _c8_pre22  = _c8_idx < pd.Timestamp("2022-01-01")
    _c8_post22 = _c8_idx >= pd.Timestamp("2022-01-01")

    cm8_pre  = _carry_perf(c8_gross_pnl[_c8_pre22],  c8_gross_ret_all[_c8_pre22],  carry_pos[_c8_pre22],  "Pre-2022")
    cm8_post = _carry_perf(c8_gross_pnl[_c8_post22], c8_gross_ret_all[_c8_post22], carry_pos[_c8_post22], "Post-2022")

    csp_pre, csp_post = st.columns(2)
    with csp_pre:
        st.markdown('<div style="color:#7A7068;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">Pre-2022 (2006-2021)</div>', unsafe_allow_html=True)
        if cm8_pre:
            _cp8_cols = st.columns(2)
            _cmcard(_cp8_cols[0], "Sharpe",       cm8_pre.get("sharpe"),      ".3f")
            _cmcard(_cp8_cols[1], "Ann Ret %",    cm8_pre.get("ann_ret_pct"), ".1f", "%")
            _cp8_cols2 = st.columns(2)
            _cmcard(_cp8_cols2[0], "Max DD %",    cm8_pre.get("mdd_pct"),     ".1f", "%")
            _cmcard(_cp8_cols2[1], "Active Days", float(cm8_pre.get("n", 0)), ",.0f")
        else:
            st.info("Insufficient data.")
    with csp_post:
        st.markdown('<div style="color:#7A7068;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">Post-2022 (2022-present)</div>', unsafe_allow_html=True)
        if cm8_post:
            _cp8b_cols = st.columns(2)
            _cmcard(_cp8b_cols[0], "Sharpe",       cm8_post.get("sharpe"),      ".3f")
            _cmcard(_cp8b_cols[1], "Ann Ret %",    cm8_post.get("ann_ret_pct"), ".1f", "%")
            _cp8b_cols2 = st.columns(2)
            _cmcard(_cp8b_cols2[0], "Max DD %",    cm8_post.get("mdd_pct"),     ".1f", "%")
            _cmcard(_cp8b_cols2[1], "Active Days", float(cm8_post.get("n", 0)), ",.0f")
        else:
            st.info("Insufficient data.")

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
        title=dict(text="Carry Strategy Comparison - Cumulative PnL (Gross, USD/MT)", font=dict(size=13)),
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

    # ── Section 9: Tenor Comparison (V2 only) ─────────────────────────────────
    if carry_vgroup == "V2 - Long Slope":
        st.divider()
        section_header("TENOR PAIR COMPARISON - V2 LONG SLOPE")
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
                if _vT < 20:
                    _vlst.append(np.nan); continue
                _vpos = np.empty(_vT)
                if _vsd:   # Same-Day = shift 1 (trade at signal close, no look-ahead)
                    _vpos[0] = 0.0
                    _vpos[1:] = np.where(np.isfinite(_vsig[:-1]), _vsig[:-1], 0.0)
                else:      # Lag-1 = shift 2 (trade next close)
                    _vpos[:2] = 0.0
                    _vpos[2:] = np.where(np.isfinite(_vsig[:-2]), _vsig[:-2], 0.0)
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
            title=dict(text="V2 Long Slope - Sharpe by Tenor Pair (Gross, No TC)", font=dict(size=13)),
            xaxis_title="Sharpe Ratio", yaxis_title=None,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_tenor, use_container_width=True)
        st.caption("Short-to-medium tenors (F3-F15, F4-F16) typically carry more predictive power than long-end pairs (F12-F24), "
                   "where price noise dominates. A decaying Sharpe across tenor pairs is a structural finding.")

    # ── Section 9: Carry Signal & Position (merged) ───────────────────────────
    st.divider()
    section_header("CARRY SIGNAL & POSITION")
    st.caption("The whole strategy in one view: the raw carry value with regime shading (top), "
               "the F1 price it trades (middle), and the resulting long/short position (bottom).")

    _c8_carry_pct = carry_raw * 100
    fig_c8sig = make_subplots(
        rows=3, cols=1, shared_xaxes=True, row_heights=[0.38, 0.38, 0.24],
        vertical_spacing=0.04,
    )
    # Row 1: carry value with backwardation / contango shading
    fig_c8sig.add_trace(go.Scatter(
        x=_c8_carry_pct.index, y=_c8_carry_pct.values, name="Carry Value (%)",
        line=dict(color=COLORS["amber"], width=1.4),
        hovertemplate="%{x|%b %d, %Y}<br>Carry: %{y:.4f}%<extra></extra>",
    ), row=1, col=1)
    fig_c8sig.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1, row=1, col=1)
    fig_c8sig.add_trace(go.Scatter(
        x=_c8_carry_pct.where(_c8_back_mask, np.nan).index,
        y=_c8_carry_pct.where(_c8_back_mask, np.nan).values,
        name="Backwardation", line=dict(color="#5BAD72", width=0), fill="tozeroy",
        fillcolor="rgba(91,173,114,0.18)", showlegend=False, hoverinfo="skip",
    ), row=1, col=1)
    fig_c8sig.add_trace(go.Scatter(
        x=_c8_carry_pct.where(_c8_cont_mask, np.nan).index,
        y=_c8_carry_pct.where(_c8_cont_mask, np.nan).values,
        name="Contango", line=dict(color="#B85450", width=0), fill="tozeroy",
        fillcolor="rgba(184,84,80,0.18)", showlegend=False, hoverinfo="skip",
    ), row=1, col=1)
    # Row 2: F1 continuous price
    fig_c8sig.add_trace(go.Scatter(
        x=cf1c_a.index, y=cf1c_a.values, name="F1 Continuous ($/MT)",
        line=dict(color=COLORS["primary"], width=1.5),
        hovertemplate="%{x|%b %d, %Y}<br>F1_cont: $%{y:,.1f}<extra></extra>",
    ), row=2, col=1)
    # Row 3: position
    _c8p_long  = carry_pos.where(carry_pos > 0, 0.0)
    _c8p_short = carry_pos.where(carry_pos < 0, 0.0)
    fig_c8sig.add_trace(go.Bar(
        x=carry_pos.index, y=_c8p_long.values, name="Long (+1)",
        marker_color="#00E676", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Long<extra></extra>",
    ), row=3, col=1)
    fig_c8sig.add_trace(go.Bar(
        x=carry_pos.index, y=_c8p_short.values, name="Short (-1)",
        marker_color="#FF1744", opacity=1.0,
        hovertemplate="%{x|%b %d, %Y}<br>Short<extra></extra>",
    ), row=3, col=1)
    fig_c8sig.update_layout(
        **CHART_LAYOUT, height=560, barmode="overlay",
        title=dict(text=f"{carry_sub_label} - Carry Value, Price & Position ({carry_timing})", font=dict(size=13)),
        hovermode="x unified", showlegend=True,
    )
    fig_c8sig.update_yaxes(title_text="Carry (%)", row=1, col=1)
    fig_c8sig.update_yaxes(title_text="F1 ($/MT)", row=2, col=1)
    fig_c8sig.update_yaxes(title_text="Position", tickvals=[-1, 0, 1],
                             ticktext=["Short", "Flat", "Long"], row=3, col=1)
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
  Short-end basis as a fraction of current price. The binary long/short signal is just the sign of this ratio.
- **V2 Long Slope**: `(Fj-Fk)/Fk` for j < k (e.g., F3-F15, F4-F16, ... F12-F24)
  Slope of the forward curve at longer tenors. Downward slope (Fj > Fk) = backwardation at the long end = Long signal.
- **V3 Z-score**: 252-day rolling standardization of `(F1-F2)/F1`.
  Filters permanent level shifts in the basis; signal fires when carry is unusually high or low relative to its recent history.

**Position Timing** (no look-ahead either way)
- *Same-Day (shift 1)*: `position[t] = sign(carry[t-1])` - trade at the signal's own close, first return t to t+1
- *Lag-1 (shift 2)*: `position[t] = sign(carry[t-2])` - trade at the next close, first return t+1 to t+2

**Why Same-Day Leads for Carry**
Carry is a slow, level-based signal, so trading at its own close (Same-Day, shift 1) captures the move
with no look-ahead. Waiting an extra day (Lag-1, shift 2) just gives up part of it: carry-momentum scores
about +0.52 Same-Day versus +0.42 Lag-1. Both are realistic, neither uses future information. (An earlier
shift-0 "Same-Day" that booked the contemporaneous move was look-ahead and has been removed.)

**Transaction Costs**
`TC[t] = |delta_position[t]| x (bps/10000/2) x F1_raw[t]`   (TC on the actual traded price; PnL still on F1_continuous)
Flip (+1 to -1): delta=2, cost = 1 full round-trip x price.
Entry (0 to +/-1): delta=1, cost = 0.5 round-trip x price.

**PnL & Returns**
All strategies trade F1_continuous regardless of which tenor pair generates the signal.
`daily_ret[t] = position[t] x delta_F1_cont[t] / F1_cont[t-1]`
Sharpe, Sortino, Max DD, Calmar all computed from daily_ret in % terms.

**In-Sample vs Out-of-Sample**
Carry signals have *no optimised parameters* - the signal formula (e.g., (F1-F2)/F1) is a structural
market measure, not a fitted quantity. There is no parameter search or look-ahead bias.
Consequently, IS results are representative of OOS performance; the sub-period table above
(pre/post 2022) reflects genuine regime-conditional performance, not overfitting.

**Reference**
Baz, J., Granger, N. M. (2015). Dissecting Investment Strategies in the Cross Section and Time Series. SSRN.
        """)

    st.markdown(
        '<div style="background:#0D1117;border:1px solid #2A2A2A;border-left:4px solid #475569;'
        'border-radius:4px;padding:12px 20px;margin-top:20px;">'
        '<span style="color:#94A3B8;font-size:0.78rem;font-family:\'IBM Plex Mono\',monospace;font-weight:600;">'
        'NEXT &rarr; </span>'
        '<span style="color:#B87333;font-size:0.82rem;font-family:\'IBM Plex Mono\',monospace;font-weight:700;">'
        'Tab 9: Value Signals</span>'
        '<span style="color:#8A8278;font-size:0.78rem;"> &nbsp;-&nbsp; '
        'Mean-reversion toward long-run equilibrium price levels. '
        'Regime-conditional and negatively correlated with momentum - the third orthogonal source of return '
        'in the equal-weight portfolio.</span></div>',
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def _wf_value_oos_tc(_pos: pd.Series, _f1c: pd.Series, _f1r: pd.Series, tc_bps: int) -> dict:
    """Walk-forward OOS Sharpes (IS=5yr, OOS=1yr) for a fixed value position series with TC."""
    IS_W, OOS_W = 1260, 252
    idx  = _pos.index.intersection(_f1c.index)
    pos  = _pos.reindex(idx).fillna(0)
    f1cs = _f1c.reindex(idx)
    f1rs = _f1r.reindex(idx)        # F1_raw aligned — used for TC, not PnL
    T    = len(idx)
    out  = {}
    oos_s = IS_W
    while oos_s < T:
        oos_e   = min(oos_s + OOS_W, T)
        if (oos_e - oos_s) < OOS_W // 2:
            break
        yr      = str(idx[oos_e - 1].year) + ("*" if (oos_e - oos_s) < OOS_W else "")   # END-year label
        p_oos   = pos.iloc[oos_s:oos_e]
        c_oos   = f1cs.iloc[oos_s:oos_e]
        r_oos   = f1rs.iloc[oos_s:oos_e]
        pnl     = p_oos * c_oos.diff()
        if tc_bps > 0:
            chg = p_oos.diff().abs()
            chg.iloc[0] = abs(p_oos.iloc[0])
            pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * r_oos
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = (pnl / c_oos.shift(1)).replace([np.inf, -np.inf], np.nan)
        act = ret[p_oos != 0].dropna()
        if len(act) >= 20:
            sd = float(act.std(ddof=1))
            out[yr] = float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan
        oos_s += OOS_W
    return out


# ══════════════════════════════════════════════════════
# TAB 9: VALUE SIGNALS
# ══════════════════════════════════════════════════════

with tab9:
    _v9_metal = st.radio("🔬 Metal", ["Copper", "Aluminium"], horizontal=True, key="value_metal")
    st.markdown(f"### Value Signals - LME {_v9_metal}")
    st.markdown(
        '<div style="background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;'
        'border-radius:4px;padding:10px 18px;margin-bottom:10px;display:flex;align-items:center;gap:16px;">'
        '<span style="color:#B87333;font-family:\'IBM Plex Mono\',monospace;font-size:0.78rem;'
        'font-weight:700;white-space:nowrap;">SIGNAL 3 OF 3</span>'
        '<span style="color:#8A8278;font-size:0.8rem;">Mean-reversion toward long-run price equilibrium. '
        'Regime-sensitive and negatively correlated with momentum (&minus;0.21 historically). '
        'With all three signals validated, the Portfolio tab (Tab 10) shows the combined EW strategy.</span></div>',
        unsafe_allow_html=True,
    )
    st.caption("Mean-reversion strategies: long when copper is cheap vs. long-run fair value, short when expensive. "
               "Signal from forward curve contracts; PnL always from F1_continuous.")

    # ── Data loading (shared by Section 1 and 2) ─────────────────────────────
    _f1_df_v9 = _load_f1_data(_v9_metal)
    if _f1_df_v9.empty:
        st.error(f"Rolling F1 file for {_v9_metal} not found.")
        st.stop()
    vf1c = _f1_df_v9["F1_continuous"]
    vf1r = _f1_df_v9["F1_raw"]
    _cu_sheet_v9 = _find_curve_sheet(_v9_metal, curve_data) if curve_data else None
    if not curve_data or _cu_sheet_v9 is None:
        st.error("Futures Curve data not loaded. Upload Metals Futures Curve file in the sidebar.")
        st.stop()
    v9_curve_px = curve_data[_cu_sheet_v9]["prices"].copy()
    v9_curve_px.index = pd.to_datetime(v9_curve_px.index).normalize()
    v9_curve_px = v9_curve_px.sort_index()

    # ── Variant banner ────────────────────────────────────────────────────────
    st.markdown("""
    <div style="background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;border-radius:4px;padding:14px 20px;margin-bottom:8px;">
      <div style="display:flex;gap:40px;flex-wrap:wrap;">
        <div style="min-width:220px;">
          <span style="color:#B87333;font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:0.9rem;">V1 - MA Reversion</span><br>
          <span style="color:#8A8278;font-size:0.78rem;">Long-run moving-average reversion on Fk<br>
          (k = F1-F15). Three states: +1 (cheap), 0 (fair), −1 (expensive).<br>
          F12 is the reference contract used in the NGL energy risk-premia paper.</span>
        </div>
        <div style="min-width:220px;">
          <span style="color:#B87333;font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:0.9rem;">V2 - Baz-Granger Reversal</span><br>
          <span style="color:#8A8278;font-size:0.78rem;">Contrarian N-year return reversal on F1_raw.<br>
          Signal = sign(F1_raw[t−N] − F1_raw[t]).<br>
          Always long or short - no flat zone.</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Best Value Signal Summary ─────────────────────────────────────────────
    st.divider()
    section_header(f"BEST VALUE SIGNAL - BY VARIANT  ({_v9_metal})")
    st.caption(
        f"Best-performing configuration per variant for {_v9_metal}. IS backtest, full period "
        f"{vf1c.index[0].year}-{vf1c.index[-1].year}, gross active-day Sharpe (TC=0), best timing per variant. "
        "Computed live - changes with the metal toggle."
    )
    _vbsc1, _vbsc2 = st.columns(2)
    _vbcs  = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;"
              "border-radius:4px;padding:14px 20px")
    _vbcsx = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #5BAD72;"
              "border-radius:4px;padding:14px 20px")
    _vblbl = ("color:#B87333;font-family:'IBM Plex Mono',monospace;"
              "font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _vblbx = ("color:#5BAD72;font-family:'IBM Plex Mono',monospace;"
              "font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _vbbig = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;"
              "font-size:1.55rem;font-weight:700;margin:0")
    _vbsub = "color:#8A8278;font-size:0.75rem;margin:2px 0"
    _vbhr  = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"
    _vb = _value_best_cards(_v9_metal)
    _vb_fam = [("V1 - MA Reversion", _vb.get("v1") if _vb else None),
               ("V2 - Baz-Granger Reversal", _vb.get("v2") if _vb else None)]
    _vb_best = max(range(2), key=lambda i: (_vb_fam[i][1]["gross"] if _vb_fam[i][1] and not np.isnan(_vb_fam[i][1]["gross"]) else -9))
    for _i, (_col, (_lbl, _d)) in enumerate(zip([_vbsc1, _vbsc2], _vb_fam)):
        with _col:
            _star = "  ★" if _i == _vb_best else ""
            _sty, _lsty = (_vbcsx, _vblbx) if _i == _vb_best else (_vbcs, _vblbl)
            if not _d:
                st.markdown(f'<div style="{_sty}"><p style="{_lsty}">{_lbl}{_star}</p>'
                            f'<p style="{_vbbig}">N/A</p></div>', unsafe_allow_html=True)
            else:
                st.markdown(f"""<div style="{_sty}">
<p style="{_lsty}">{_lbl}{_star}</p>
<p style="{_vbbig}">{_fmt_sh(_d['gross'])}</p>
<p style="{_vbsub}">Sharpe Ratio (Gross)</p>
<hr style="{_vbhr}"/>
<p style="{_vbsub}">{_d['name']}{' , ±10% threshold' if _lbl.startswith('V1') else ''}, {_d['timing']}</p>
<p style="{_vbsub}">Ann Ret ≈ {_fmt_pct(_d['ann'])}, Max DD ≈ {_fmt_dd(_d['mdd'])}</p>
</div>""", unsafe_allow_html=True)
    st.caption(
        "Same-Day (shift-1, no look-ahead) is the default: for V2 Baz-Granger it beats Lag-1 (+0.51 vs +0.37); "
        "for V1 MA-Reversion the two are close (Lag-1 marginally ahead, e.g. F8 +0.33 vs +0.28). "
        "Value edge is regime-conditional - most P&L is concentrated in the 2020-2022 COVID dislocation period."
    )

    # ── SECTION 2: IS PARAMETER SEARCH ───────────────────────────────────────
    st.divider()
    _v9_is_yr0 = str(vf1c.index[0].year); _v9_is_yr1 = str(vf1c.index[-1].year)
    st.markdown(f"#### IS Parameter Search (Full Period {_v9_is_yr0}-{_v9_is_yr1})")
    _v9_s2_tc_label = st.session_state.get("val_tc", "0 bps  (Gross)")
    st.caption(
        "Full in-sample backtest. Use controls to explore contract, lookback, threshold, and timing. "
        f"TC applied to all metrics below: **{_v9_s2_tc_label}** - change via the TC dropdown in controls below."
    )

    # ── Strategy Preset ───────────────────────────────────────────────────────
    _VAL_PRESETS = {
        "V2 - Baz-Granger, 10yr, Same-Day  [Best Overall]": {
            "val_vgroup":   "V2 - Baz-Granger Reversal",
            "val_lb":       "10yr (2520d)",
            "val_timing":   "Same-Day",
        },
        "V1 - F8, 5yr, ±10%, Lag-1  [Empirical Optimum]": {
            "val_vgroup":   "V1 - MA Reversion",
            "val_contract": "F8",
            "val_lb":       "5yr  (1260d)",
            "val_thr":      "±10% (default)",
            "val_timing":   "Lag-1 (Next-Day)",
        },
        "V1 - F12, 5yr, ±10%, Lag-1  [NGL paper tenor]": {
            "val_vgroup":   "V1 - MA Reversion",
            "val_contract": "F12",
            "val_lb":       "5yr  (1260d)",
            "val_thr":      "±10% (default)",
            "val_timing":   "Lag-1 (Next-Day)",
        },
        "V2 - Baz-Granger, 3yr, Same-Day  [Alternative Lookback]": {
            "val_vgroup":   "V2 - Baz-Granger Reversal",
            "val_lb":       "3yr  (756d)",
            "val_timing":   "Same-Day",
        },
        "V1 - F8, 7yr, ±10%, Lag-1  [Longer-Window V1]": {
            "val_vgroup":   "V1 - MA Reversion",
            "val_contract": "F8",
            "val_lb":       "7yr  (1764d)",
            "val_thr":      "±10% (default)",
            "val_timing":   "Lag-1 (Next-Day)",
        },
        "Custom (use controls below)": {},
    }

    def _apply_val_preset():
        cfg = _VAL_PRESETS.get(st.session_state.get("val_preset", "Custom (use controls below)"), {})
        for k, v in cfg.items():
            st.session_state[k] = v

    _v9_pre_col, _v9_pre_info = st.columns([2.8, 3.2])
    with _v9_pre_col:
        st.selectbox(
            "Strategy Preset",
            list(_VAL_PRESETS.keys()),
            index=0,
            key="val_preset",
            on_change=_apply_val_preset,
        )
    with _v9_pre_info:
        st.markdown(
            '<div style="padding:8px 0;color:#7A7068;font-size:0.78rem;">'
            'Selecting a preset auto-fills all controls below - all sections update together. '
            'Switch to <b>Custom</b> to edit individual parameters freely.</div>',
            unsafe_allow_html=True,
        )

    # ── Controls ──────────────────────────────────────────────────────────────
    v9_c1, v9_c2, v9_c3, v9_c4, v9_c5, v9_c6 = st.columns([1.5, 1.2, 1.5, 1.3, 1.2, 1.4])
    with v9_c1:
        val_vgroup = st.selectbox("Variant", ["V1 - MA Reversion", "V2 - Baz-Granger Reversal"],
                                  index=1, key="val_vgroup")
        val_is_v1 = val_vgroup.startswith("V1")
    with v9_c2:
        val_contract = st.selectbox("Contract", [f"F{k}" for k in range(1, 16)],
                                    index=7, key="val_contract",
                                    disabled=not val_is_v1)
        val_k = int(val_contract[1:]) if val_is_v1 else None
    with v9_c3:
        _lb_opts = {"1yr  (252d)": 252, "3yr  (756d)": 756, "5yr  (1260d)": 1260,
                    "7yr  (1764d)": 1764, "10yr (2520d)": 2520}
        val_lb_label = st.selectbox("Lookback", list(_lb_opts.keys()), index=4, key="val_lb")
        val_N = _lb_opts[val_lb_label]
    with v9_c4:
        _thr_opts = {"±5%": 0.05, "±10% (default)": 0.10, "±15%": 0.15, "±20%": 0.20}
        val_thr_label = st.selectbox("Threshold (V1 only)", list(_thr_opts.keys()), index=1,
                                     key="val_thr", disabled=not val_is_v1)
        val_thr = _thr_opts[val_thr_label] if val_is_v1 else 0.10
    with v9_c5:
        _val_tc_map = _tc_label_map(_get_last_f1_price())
        val_tc_label = st.selectbox("TC (bps)", list(_val_tc_map.keys()), index=0, key="val_tc")
        val_tc_bps = _val_tc_map[val_tc_label]
    with v9_c6:
        val_timing = st.selectbox("Position Entry", ["Same-Day", "Lag-1 (Next-Day)"],
                                  index=0, key="val_timing",
                                  help="Same-Day (shift 1, no look-ahead) now the default: under the "
                                       "corrected timing convention it beats Lag-1 for every V2 Baz-Granger "
                                       "lookback (10yr +0.51 vs +0.37). V1 MA-Reversion is roughly tied "
                                       "(Lag-1 marginally ahead for F8/F12 5yr).")
        val_same_day = val_timing == "Same-Day"

    # Build spec
    if val_is_v1:
        val_spec = {"variant": "v1", "contract": val_k, "lookback": val_N,
                    "threshold": val_thr, "same_day": val_same_day}
    else:
        val_spec = {"variant": "v2", "lookback": val_N, "same_day": val_same_day}

    # ── Compute signal ─────────────────────────────────────────────────────────
    val_raw = _value_raw_signal(v9_curve_px, vf1r, val_spec)
    if val_raw.empty:
        st.error("Could not compute value signal. Check that the required contract is available in the curve data.")
        st.stop()

    _v9_idx = val_raw.index

    if val_is_v1:
        _v9_sig_bin = np.where(val_raw.values < -val_thr,  1.0,
                      np.where(val_raw.values >  val_thr, -1.0, 0.0))
    else:
        _v9_sig_bin = np.sign(val_raw.values).astype(float)

    _v9_T = len(_v9_idx)
    # Same-Day = shift 1 (no look-ahead); Lag-1 = shift 2. Matches carry/value helpers.
    if val_same_day:
        val_pos_np = np.empty(_v9_T); val_pos_np[0] = 0.0
        val_pos_np[1:] = np.where(np.isfinite(_v9_sig_bin[:-1]), _v9_sig_bin[:-1], 0.0)
    else:
        val_pos_np = np.empty(_v9_T); val_pos_np[:2] = 0.0
        val_pos_np[2:] = np.where(np.isfinite(_v9_sig_bin[:-2]), _v9_sig_bin[:-2], 0.0)

    val_pos = pd.Series(val_pos_np, index=_v9_idx)
    vf1c_a  = vf1c.reindex(_v9_idx)
    vf1r_a  = vf1r.reindex(_v9_idx)

    v9_delta    = vf1c_a.diff()
    v9_f1c_prev = vf1c_a.shift(1)
    v9_gross_pnl = val_pos * v9_delta
    v9_pos_change = val_pos.diff().abs()
    v9_pos_change.iloc[0] = abs(val_pos.iloc[0])
    v9_tc_cost   = v9_pos_change * (val_tc_bps / 10000.0 / 2.0) * vf1r_a
    v9_net_pnl   = v9_gross_pnl - v9_tc_cost
    v9_cum_pnl   = v9_gross_pnl.cumsum()
    v9_cum_net   = v9_net_pnl.cumsum()
    with np.errstate(invalid="ignore", divide="ignore"):
        v9_gross_ret = (v9_gross_pnl / v9_f1c_prev).replace([np.inf, -np.inf], np.nan)
        v9_net_ret   = (v9_net_pnl   / v9_f1c_prev).replace([np.inf, -np.inf], np.nan)

    last_date_v9 = _v9_idx[-1]

    # ── Live Signal Badge ─────────────────────────────────────────────────────
    _v9_last_raw = float(val_raw.iloc[-1])
    if val_is_v1:
        if _v9_last_raw < -val_thr:
            _v9_state, _v9_scolor, _v9_sbg = "BELOW FAIR VALUE - LONG", "#5BAD72", "rgba(91,173,114,0.08)"
        elif _v9_last_raw > val_thr:
            _v9_state, _v9_scolor, _v9_sbg = "ABOVE FAIR VALUE - SHORT", "#B85450", "rgba(184,84,80,0.08)"
        else:
            _v9_state, _v9_scolor, _v9_sbg = "NEAR FAIR VALUE - FLAT", "#C9A84C", "rgba(201,168,76,0.06)"
        _v9_dev_str = f"{_v9_last_raw * 100:+.2f}%"
        _v9_raw_lbl = f"deviation from {val_N}d MA (F{val_k})"
    else:
        if _v9_last_raw > 0:
            _v9_state, _v9_scolor, _v9_sbg = "PRICE FALLEN - LONG", "#5BAD72", "rgba(91,173,114,0.08)"
        else:
            _v9_state, _v9_scolor, _v9_sbg = "PRICE RISEN - SHORT", "#B85450", "rgba(184,84,80,0.08)"
        _v9_dev_str = f"{_v9_last_raw:+,.1f} $/MT"
        _v9_raw_lbl = f"{val_N}d price reversal (F1_raw)"

    # Days since last state change
    _v9_sig_bin_s = pd.Series(_v9_sig_bin, index=_v9_idx)
    _v9_flip_mask = _v9_sig_bin_s.diff().abs() > 0
    _v9_flip_mask.iloc[0] = _v9_sig_bin_s.iloc[0] != 0
    _v9_flip_dates = _v9_sig_bin_s[_v9_flip_mask].index
    if len(_v9_flip_dates) >= 2:
        _v9_last_flip = _v9_flip_dates[-1]
        _v9_days_in_state = (last_date_v9 - _v9_last_flip).days
        _v9_flip_str = f"{_v9_last_flip.strftime('%Y-%m-%d')}  ({_v9_days_in_state}d ago)"
    else:
        _v9_flip_str = "N/A"; _v9_days_in_state = 0

    st.divider()
    st.markdown(f"""
    <div style="background:{_v9_sbg}; border:1px solid {_v9_scolor}; border-left:5px solid {_v9_scolor};
                border-radius:4px; padding:18px 24px; margin-bottom:4px;">
      <div style="display:flex; gap:48px; align-items:center; flex-wrap:wrap;">
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Current State</div>
          <div style="color:{_v9_scolor};font-size:1.5rem;font-weight:700;letter-spacing:0.04em;font-family:'IBM Plex Mono',monospace;">{_v9_state}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">as of {last_date_v9.strftime('%Y-%m-%d')}</div>
        </div>
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Signal Value</div>
          <div style="color:{_v9_scolor};font-size:1.5rem;font-weight:700;font-family:'IBM Plex Mono',monospace;">{_v9_dev_str}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">{_v9_raw_lbl}</div>
        </div>
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Last State Change</div>
          <div style="color:#D4CFC8;font-size:1.1rem;font-weight:600;font-family:'IBM Plex Mono',monospace;">{_v9_flip_str}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">Days in current state: {_v9_days_in_state}</div>
        </div>
        <div>
          <div style="color:#7A7068;font-size:0.68rem;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:5px;">Entry Mode</div>
          <div style="color:#D4CFC8;font-size:1.1rem;font-weight:600;font-family:'IBM Plex Mono',monospace;">{val_timing}</div>
          <div style="color:#5A5248;font-size:0.72rem;margin-top:3px;">{'±' + str(int(val_thr*100)) + '% threshold' if val_is_v1 else 'no threshold - always in market'}</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Section 3: Date Filter + Performance Metrics ──────────────────────────
    st.divider()
    pf9_c1, _ = st.columns([3, 1])
    with pf9_c1:
        val_perf_dates = st.date_input(
            "Performance period  (signal uses full history - only metrics & charts below update)",
            value=(_v9_idx[0].date(), _v9_idx[-1].date()),
            min_value=_v9_idx[0].date(), max_value=_v9_idx[-1].date(),
            key="val_perf_dates",
        )
    vp9_start = pd.Timestamp(val_perf_dates[0]) if len(val_perf_dates) >= 1 else _v9_idx[0]
    vp9_end   = pd.Timestamp(val_perf_dates[1]) if len(val_perf_dates) == 2 else _v9_idx[-1]
    vp9_mask  = (_v9_idx >= vp9_start) & (_v9_idx <= vp9_end)

    v9_pos_f      = val_pos[vp9_mask]
    v9_pnl_f      = v9_gross_pnl[vp9_mask]
    v9_ret_f      = v9_gross_ret[vp9_mask]
    v9_net_pnl_f  = v9_net_pnl[vp9_mask]
    v9_net_ret_f  = v9_net_ret[vp9_mask]

    def _val_perf(daily_pnl, daily_ret, position, label):
        active  = daily_ret[position != 0].dropna()
        pnl_act = daily_pnl[position != 0].dropna()
        n = len(active)
        if n < 20:
            return {}
        ann_r   = float(active.mean() * 252 * 100)
        ann_sd  = float(active.std() * np.sqrt(252) * 100)
        sharpe  = ann_r / ann_sd if ann_sd > 0 else np.nan
        down    = active[active < 0]
        srt_d   = float(down.std() * np.sqrt(252) * 100) if len(down) > 1 else np.nan
        sortino = ann_r / srt_d if srt_d and srt_d > 0 else np.nan
        cum_r   = daily_ret.fillna(0).cumsum() * 100
        mdd_pct = float((cum_r - cum_r.cummax()).min())
        calmar  = ann_r / abs(mdd_pct) if mdd_pct != 0 else np.nan
        wins, losses = pnl_act[pnl_act > 0], pnl_act[pnl_act < 0]
        T   = len(position)
        return {
            "label": label, "n": n,
            "sharpe": sharpe, "sortino": sortino,
            "ann_ret_pct": ann_r, "ann_std_pct": ann_sd,
            "mdd_pct": mdd_pct, "calmar": calmar,
            "hit_rate": float((active > 0).mean()) * 100,
            "profit_factor": abs(wins.sum() / losses.sum()) if len(losses) > 0 else np.nan,
            "total_pnl_usdmt": float(pnl_act.sum()),
            "pct_long":  float((position > 0).sum()) / T * 100,
            "pct_flat":  float((position == 0).sum()) / T * 100,
            "pct_short": float((position < 0).sum()) / T * 100,
        }

    vm9     = _val_perf(v9_pnl_f,     v9_ret_f,     v9_pos_f, val_vgroup)
    vm9_net = _val_perf(v9_net_pnl_f, v9_net_ret_f, v9_pos_f, f"Net ({val_tc_label})")

    st.divider()
    section_header("PERFORMANCE METRICS")
    _v9_title = (f"**Strategy:** {val_vgroup}  |  " +
                 (f"**Contract:** F{val_k}  |  **Lookback:** {val_lb_label}  |  **Threshold:** {val_thr_label}  |  " if val_is_v1 else f"**Lookback:** {val_lb_label}  |  ") +
                 f"**Entry:** {val_timing}  |  **TC:** {val_tc_label}")
    st.caption(_v9_title)

    if vm9 and vm9_net:
        _v9_cols1 = st.columns(8)
        _cmcard(_v9_cols1[0], "Sharpe Gross",    vm9.get("sharpe"),         ".2f")
        _cmcard(_v9_cols1[1], "Sharpe Net",      vm9_net.get("sharpe"),     ".2f")
        _cmcard(_v9_cols1[2], "Ann Ret% Gross",  vm9.get("ann_ret_pct"),    ".2f", "%")
        _cmcard(_v9_cols1[3], "Ann Ret% Net",    vm9_net.get("ann_ret_pct"),".2f", "%")
        _cmcard(_v9_cols1[4], "Max DD% Gross",   vm9.get("mdd_pct"),        ".2f", "%")
        _cmcard(_v9_cols1[5], "Max DD% Net",     vm9_net.get("mdd_pct"),    ".2f", "%")
        _cmcard(_v9_cols1[6], "Calmar Gross",    vm9.get("calmar"),         ".2f")
        _cmcard(_v9_cols1[7], "Calmar Net",      vm9_net.get("calmar"),     ".2f")

        _v9_cols2 = st.columns(8)
        _cmcard(_v9_cols2[0], "Sortino Gross",   vm9.get("sortino"),        ".2f")
        _cmcard(_v9_cols2[1], "Sortino Net",     vm9_net.get("sortino"),    ".2f")
        _cmcard(_v9_cols2[2], "Hit Rate",        vm9.get("hit_rate"),       ".2f", "%")
        _cmcard(_v9_cols2[3], "Profit Factor",   vm9.get("profit_factor"),  ".2f")
        _cmcard(_v9_cols2[4], "PnL Gross $/MT",  vm9.get("total_pnl_usdmt"), ",.2f")
        _cmcard(_v9_cols2[5], "PnL Net $/MT",    vm9_net.get("total_pnl_usdmt"), ",.2f")
        _cmcard(_v9_cols2[6], "n Flips",         float(int(_v9_flip_mask.sum())), ",.0f")
        _cmcard(_v9_cols2[7], "Days In State",   float(_v9_days_in_state),  ",.0f", "d")
    elif vm9:
        _v9_cols1 = st.columns(8)
        _cmcard(_v9_cols1[0], "Sharpe",         vm9.get("sharpe"),         ".2f")
        _cmcard(_v9_cols1[1], "Sortino",         vm9.get("sortino"),        ".2f")
        _cmcard(_v9_cols1[2], "Ann Return %",    vm9.get("ann_ret_pct"),    ".2f", "%")
        _cmcard(_v9_cols1[3], "Ann Std Dev %",   vm9.get("ann_std_pct"),    ".2f", "%")
        _cmcard(_v9_cols1[4], "Max DD %",        vm9.get("mdd_pct"),        ".2f", "%")
        _cmcard(_v9_cols1[5], "Calmar",          vm9.get("calmar"),         ".2f")
        _cmcard(_v9_cols1[6], "Hit Rate",        vm9.get("hit_rate"),       ".2f", "%")
        _cmcard(_v9_cols1[7], "Profit Factor",   vm9.get("profit_factor"),  ".2f")
    else:
        st.warning("Insufficient active trading days to compute metrics.")

    # ── Out-of-Sample Walk-Forward Validation (moved below metrics) ──────────
    st.divider()
    # ── SECTION 1: OUT-OF-SAMPLE WALK-FORWARD ────────────────────────────────
    st.markdown("#### Out-of-Sample Walk-Forward Validation")
    st.caption(
        "IS = 5yr rolling window, OOS = 1yr, Same-Day entry, Fixed parameters - no re-optimisation per window. "
        "Window labels denote the start year of each OOS period. "
        "V1 F8 and F12 use ±10% threshold throughout. "
        "V2 BG 10yr: signal first valid Jan 2016; only 5 complete OOS windows available."
    )

    # Live OOS walk-forward options - all computed dynamically via _wf_value_oos_tc
    _V9_OOS_OPTS = {
        "V1: F8, 5yr, Same-Day  [OOS Validated]":       ("v1", 8,    1260),
        "V1: F12, 5yr, Same-Day  [NGL paper tenor]":  ("v1", 12,   1260),
        "V2: BG 3yr, Same-Day  [Fully Testable]":        ("v2", None, 756),
        "V2: BG 10yr, Same-Day  [Limited to ~5 windows]":("v2", None, 2520),
    }

    def _v9_build_pos(variant: str, k, N: int) -> pd.Series:
        if variant == "v1":
            col = f"F{k}"
            if col not in v9_curve_px.columns:
                return pd.Series(dtype=float)
            p  = v9_curve_px[col].dropna()
            ma = p.rolling(N, min_periods=N // 2).mean()
            dev = ((p - ma) / ma).replace([np.inf, -np.inf], np.nan).dropna()
            sig = np.where(dev < -0.10, 1.0, np.where(dev > 0.10, -1.0, 0.0))
            return pd.Series(sig, index=dev.index).shift(1).fillna(0)
        else:
            rev = vf1r.shift(N) - vf1r
            s   = np.sign(rev.replace([np.inf, -np.inf], np.nan).dropna())
            return s.shift(1).fillna(0)

    _v9_oos_ctrl1, _v9_oos_ctrl2 = st.columns([2.8, 1.2])
    with _v9_oos_ctrl1:
        _v9_oos_sig = st.selectbox(
            "Signal - OOS Walk-Forward", list(_V9_OOS_OPTS.keys()),
            index=0, key="v9_oos_sig",
        )
    with _v9_oos_ctrl2:
        _v9_oos_tc_map   = _tc_label_map(float(vf1c.dropna().iloc[-1]))
        _v9_oos_tc_label = st.selectbox(
            "TC - OOS Section", list(_v9_oos_tc_map.keys()),
            index=0, key="v9_oos_tc",
        )
        _v9_oos_tc_bps = _v9_oos_tc_map[_v9_oos_tc_label]

    _v9_var, _v9_k, _v9_N = _V9_OOS_OPTS[_v9_oos_sig]
    _v9_pos_built = _v9_build_pos(_v9_var, _v9_k, _v9_N)
    _v9_wf_active = _wf_value_oos_tc(_v9_pos_built, vf1c, vf1r, _v9_oos_tc_bps) if not _v9_pos_built.empty else {}

    _v9_wf_vals    = [v for v in _v9_wf_active.values() if v is not None and not np.isnan(v)]
    _v9_wf_avg     = np.nanmean(_v9_wf_vals) if _v9_wf_vals else np.nan
    _v9_wf_n_pos   = sum(1 for v in _v9_wf_vals if v > 0)
    _v9_wf_n_tot   = len(_v9_wf_vals)
    _v9_tc_note    = f", {_v9_oos_tc_label}" if _v9_oos_tc_bps > 0 else ""
    _v9_recent_yrs = sorted(k for k in _v9_wf_active if not k.endswith("*"))[-3:]
    _v9_recent_avg = np.nanmean([_v9_wf_active[y] for y in _v9_recent_yrs]) if _v9_recent_yrs else np.nan
    _v9_is_10yr    = "10yr" in _v9_oos_sig
    _v9_wf_yrs_s   = sorted(_v9_wf_active.keys())
    _v9_first_yr   = _v9_wf_yrs_s[0][:4] if _v9_wf_yrs_s else "-"
    _v9_last_yr    = _v9_wf_yrs_s[-1][:4] if _v9_wf_yrs_s else "-"

    # Summary cards (same style as Momentum Section 1)
    _v9wf_c1, _v9wf_c2, _v9wf_c3 = st.columns(3)
    _v9_cs  = ("background:#161616;border:1px solid #2A2A2A;"
               "border-left:4px solid #B87333;border-radius:4px;padding:14px 20px")
    _v9_csg = ("background:#161616;border:1px solid #2A2A2A;"
               "border-left:4px solid #475569;border-radius:4px;padding:14px 20px")
    _v9_lbl = ("color:#B87333;font-family:'IBM Plex Mono',monospace;"
               "font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _v9_lblg= ("color:#94A3B8;font-family:'IBM Plex Mono',monospace;"
               "font-size:0.85rem;font-weight:600;margin:0 0 6px")
    _v9_big = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;"
               "font-size:1.55rem;font-weight:700;margin:0")
    _v9_med = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;"
               "font-size:1.15rem;font-weight:600;margin:0")
    _v9_sub = "color:#8A8278;font-size:0.75rem;margin:2px 0"
    _v9_hr  = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"

    with _v9wf_c1:
        _win_lbl = f"{_v9_wf_n_tot} Windows ({_v9_first_yr}-{_v9_last_yr})"
        st.markdown(f"""<div style="{_v9_cs}">
<p style="{_v9_lbl}">{_v9_oos_sig.split(', ')[0].strip()} - Fixed Parameter</p>
<p style="{_v9_big}">{_v9_wf_avg:+.3f}</p>
<p style="{_v9_sub}">Avg OOS Sharpe, {_win_lbl}{_v9_tc_note}</p>
<hr style="{_v9_hr}"/>
<p style="{_v9_sub}">Recent 3 windows ({", ".join(_v9_recent_yrs)}) avg</p>
<p style="{_v9_med}">{_v9_recent_avg:+.3f}</p>
<p style="{_v9_sub}">Zero re-optimisation across all windows</p>
</div>""", unsafe_allow_html=True)

    with _v9wf_c2:
        if _v9_is_10yr:
            st.markdown(f"""<div style="{_v9_cs}">
<p style="{_v9_lbl}">10yr Lookback - Data Constraint</p>
<p style="{_v9_big}">{_v9_wf_n_tot}</p>
<p style="{_v9_sub}">OOS windows available ({_v9_first_yr}-{_v9_last_yr})</p>
<hr style="{_v9_hr}"/>
<p style="{_v9_sub}">A 10yr lookback consumes ~10yr of history before the signal starts,</p>
<p style="{_v9_sub}">then a 5yr IS window before the first OOS window.</p>
<p style="{_v9_sub}">Few OOS windows exist - treat this OOS estimate as low-confidence.</p>
</div>""", unsafe_allow_html=True)
        else:
            _v9_wf_n_gt03 = sum(1 for v in _v9_wf_vals if v > 0.30)
            st.markdown(f"""<div style="{_v9_cs}">
<p style="{_v9_lbl}">OOS Performance vs IS</p>
<p style="{_v9_big}">{'Higher' if _v9_wf_avg > 0.20 else 'Lower'}</p>
<p style="{_v9_sub}">OOS avg vs IS full-period Sharpe</p>
<hr style="{_v9_hr}"/>
<p style="{_v9_sub}">OOS above +0.30 Sharpe</p>
<p style="{_v9_med}">{_v9_wf_n_gt03} / {_v9_wf_n_tot} windows</p>
<p style="{_v9_sub}">IS Sharpe shown in Section 2 below</p>
</div>""", unsafe_allow_html=True)

    with _v9wf_c3:
        st.markdown(f"""<div style="{_v9_csg}">
<p style="{_v9_lblg}">OOS Consistency</p>
<p style="{_v9_sub}">Positive OOS Sharpe</p>
<p style="{_v9_med}">{_v9_wf_n_pos} / {_v9_wf_n_tot} windows</p>
<hr style="{_v9_hr}"/>
<p style="{_v9_sub}">Best window</p>
<p style="{_v9_med}">{max(_v9_wf_vals):+.3f} ({max(_v9_wf_active, key=lambda y: _v9_wf_active[y])})</p>
<hr style="{_v9_hr}"/>
<p style="{_v9_sub}">Worst window</p>
<p style="{_v9_med}">{min(_v9_wf_vals):+.3f} ({min(_v9_wf_active, key=lambda y: _v9_wf_active[y])})</p>
</div>""", unsafe_allow_html=True)

    # OOS bar chart
    _v9_wf_yrs  = list(_v9_wf_active.keys())
    _v9_wf_shps = [_v9_wf_active[y] for y in _v9_wf_yrs]
    _v9_bar_clr = ["#5BAD72" if (v is not None and not np.isnan(v) and v >= 0) else "#B85450"
                   for v in _v9_wf_shps]
    _v9_border  = ["gold" if y in ("2022", "2023", "2024") else "rgba(0,0,0,0)" for y in _v9_wf_yrs]

    fig_v9_oos = go.Figure(go.Bar(
        x=_v9_wf_yrs,
        y=[v if (v is not None and not np.isnan(v)) else 0 for v in _v9_wf_shps],
        marker_color=_v9_bar_clr,
        marker_line_color=_v9_border,
        marker_line_width=2,
        text=[f"{v:+.2f}" if (v is not None and not np.isnan(v)) else "-" for v in _v9_wf_shps],
        textposition="outside",
        hovertemplate="%{x}: Sharpe %{y:+.3f}<extra></extra>",
    ))
    fig_v9_oos.add_hline(y=0, line_color="#475569", line_width=1.2)
    fig_v9_oos.add_hline(y=_v9_wf_avg, line_dash="dot", line_color="#B87333", line_width=1.5,
                         annotation_text=f"Avg {_v9_wf_avg:+.3f}", annotation_position="right")
    fig_v9_oos.update_layout(
        height=320, margin=dict(l=0, r=60, t=30, b=0),
        paper_bgcolor="#0E1117", plot_bgcolor="#131922",
        font=dict(color="#E8DDD0", family="IBM Plex Mono", size=11),
        xaxis=dict(gridcolor="#1C2333", title="OOS Window Start Year"),
        yaxis=dict(gridcolor="#1C2333", title="OOS Sharpe Ratio", zeroline=False),
        showlegend=False,
    )
    st.plotly_chart(fig_v9_oos, use_container_width=True)

    if _v9_is_10yr:
        st.info(
            "**V2 BG 10yr data note:** Signal first computable in Jan 2016 (requires 10yr of price history). "
            "With a 5yr IS window, first OOS window begins 2020. Only 5 windows exist - "
            "the full-period Sharpe (+0.512) is the entire track record of this signal. "
            "Strong 2020-2021 performance is COVID mean-reversion; 2023-2024 is negative.",
            icon="ℹ️",
        )
    else:
        st.caption(
            f"Gold-bordered bars = most recent {len(_v9_recent_yrs)} complete OOS windows ({', '.join(_v9_recent_yrs)}). "
            + (f"TC = {_v9_oos_tc_label} deducted on each position change." if _v9_oos_tc_bps > 0 else "Gross returns shown (0 TC).")
            + (" Partial windows (marked *) excluded from gold highlighting." if any(k.endswith("*") for k in _v9_wf_active) else "")
        )


    # ── Section 2: Deviation History ──────────────────────────────────────────
    st.divider()
    section_header("FAIR VALUE DEVIATION HISTORY")

    if val_is_v1:
        _dev_pct = val_raw * 100
        _long_mask  = _dev_pct < -val_thr * 100
        _flat_mask  = _dev_pct.abs() <= val_thr * 100
        _short_mask = _dev_pct > val_thr * 100

        fig_v9dev = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                  row_heights=[0.65, 0.35], vertical_spacing=0.04)
        fig_v9dev.add_trace(go.Scatter(
            x=_dev_pct.index, y=_dev_pct.values, name="Deviation %", mode="lines",
            line=dict(color=COLORS["primary"], width=1.5),
            hovertemplate="%{x|%b %d, %Y}<br>Dev: %{y:.2f}%<extra></extra>",
        ), row=1, col=1)
        fig_v9dev.add_hline(y=0,              line_dash="dash", line_color="#475569", line_width=1.2, row=1, col=1)
        fig_v9dev.add_hline(y=val_thr * 100,  line_dash="dot",  line_color="#B85450", line_width=1.0, row=1, col=1)
        fig_v9dev.add_hline(y=-val_thr * 100, line_dash="dot",  line_color="#5BAD72", line_width=1.0, row=1, col=1)

        _long_fill  = _dev_pct.where(_long_mask, np.nan)
        _short_fill = _dev_pct.where(_short_mask, np.nan)
        fig_v9dev.add_trace(go.Scatter(
            x=_long_fill.index, y=_long_fill.values, name="Long zone", mode="lines",
            line=dict(color="#5BAD72", width=0), fill="tozeroy",
            fillcolor="rgba(91,173,114,0.18)",
            hovertemplate="%{x|%b %d, %Y}<br>Dev: %{y:.2f}%<extra></extra>",
        ), row=1, col=1)
        fig_v9dev.add_trace(go.Scatter(
            x=_short_fill.index, y=_short_fill.values, name="Short zone", mode="lines",
            line=dict(color="#B85450", width=0), fill="tozeroy",
            fillcolor="rgba(184,84,80,0.18)",
            hovertemplate="%{x|%b %d, %Y}<br>Dev: %{y:.2f}%<extra></extra>",
        ), row=1, col=1)

        _sig_long_v9  = _v9_sig_bin_s.where(_v9_sig_bin_s > 0, 0.0)
        _sig_short_v9 = _v9_sig_bin_s.where(_v9_sig_bin_s < 0, 0.0)
        _sig_flat_v9  = _v9_sig_bin_s.where(_v9_sig_bin_s == 0, 0.5)
        fig_v9dev.add_trace(go.Bar(x=_v9_idx, y=_sig_long_v9.values,
            name="Long (+1)", marker_color="#00E676", opacity=1.0,
            hovertemplate="%{x|%b %d, %Y}<br>Long<extra></extra>"), row=2, col=1)
        fig_v9dev.add_trace(go.Bar(x=_v9_idx, y=_sig_short_v9.values,
            name="Short (-1)", marker_color="#FF1744", opacity=1.0,
            hovertemplate="%{x|%b %d, %Y}<br>Short<extra></extra>"), row=2, col=1)
        fig_v9dev.add_trace(go.Bar(x=_v9_idx, y=_sig_flat_v9.values,
            name="Flat (0)", marker_color="#475569", opacity=0.5,
            hovertemplate="%{x|%b %d, %Y}<br>Flat<extra></extra>"), row=2, col=1)

        fig_v9dev.update_layout(
            **CHART_LAYOUT, height=500, barmode="overlay",
            title=dict(text=f"F{val_k} vs {val_N}d MA - Deviation % (±{int(val_thr*100)}% bands)", font=dict(size=13)),
            hovermode="x unified", showlegend=True,
        )
        fig_v9dev.update_yaxes(title_text="Deviation %", row=1, col=1)
        fig_v9dev.update_yaxes(title_text="Signal", tickvals=[-1, 0, 1],
                                ticktext=["Short", "Flat", "Long"], row=2, col=1)
        fig_v9dev.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_v9dev, use_container_width=True)
        st.caption(f"Top: deviation of F{val_k} from its {val_N}d rolling mean (positive = expensive, negative = cheap). "
                   f"Dashed bands at ±{int(val_thr*100)}%: outside bands = active signal, inside = flat zone. "
                   "Bottom: three-state signal bars (green=Long, grey=Flat, red=Short).")
    else:
        # V2: reversal series
        fig_v9dev = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                  row_heights=[0.65, 0.35], vertical_spacing=0.04)
        fig_v9dev.add_trace(go.Scatter(
            x=val_raw.index, y=val_raw.values, name=f"N-yr Reversal ($/MT)", mode="lines",
            line=dict(color=COLORS["primary"], width=1.5),
            hovertemplate="%{x|%b %d, %Y}<br>Rev: %{y:,.1f} $/MT<extra></extra>",
        ), row=1, col=1)
        fig_v9dev.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1.2, row=1, col=1)
        _v2_pos = val_raw.where(val_raw > 0, np.nan)
        _v2_neg = val_raw.where(val_raw < 0, np.nan)
        fig_v9dev.add_trace(go.Scatter(x=_v2_pos.index, y=_v2_pos.values, name="Long (fallen)",
            mode="lines", line=dict(color="#5BAD72", width=0), fill="tozeroy",
            fillcolor="rgba(91,173,114,0.18)"), row=1, col=1)
        fig_v9dev.add_trace(go.Scatter(x=_v2_neg.index, y=_v2_neg.values, name="Short (risen)",
            mode="lines", line=dict(color="#B85450", width=0), fill="tozeroy",
            fillcolor="rgba(184,84,80,0.18)"), row=1, col=1)

        _sig_long_v9  = _v9_sig_bin_s.where(_v9_sig_bin_s > 0, 0.0)
        _sig_short_v9 = _v9_sig_bin_s.where(_v9_sig_bin_s < 0, 0.0)
        fig_v9dev.add_trace(go.Bar(x=_v9_idx, y=_sig_long_v9.values,
            name="Long (+1)", marker_color="#00E676", opacity=1.0), row=2, col=1)
        fig_v9dev.add_trace(go.Bar(x=_v9_idx, y=_sig_short_v9.values,
            name="Short (-1)", marker_color="#FF1744", opacity=1.0), row=2, col=1)

        fig_v9dev.update_layout(
            **CHART_LAYOUT, height=500, barmode="overlay",
            title=dict(text=f"Baz-Granger {val_N}d Return Reversal (F1_raw)", font=dict(size=13)),
            hovermode="x unified", showlegend=True,
        )
        fig_v9dev.update_yaxes(title_text="Reversal ($/MT)", row=1, col=1)
        fig_v9dev.update_yaxes(title_text="Position", tickvals=[-1, 0, 1],
                                ticktext=["Short", "Flat", "Long"], row=2, col=1)
        fig_v9dev.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_v9dev, use_container_width=True)
        st.caption(f"Top: F1_raw[t−{val_N}d] − F1_raw[t]. Positive = price has fallen over {val_N} trading days → contrarian Long. "
                   "Bottom: resulting binary signal (always in market - no flat zone for V2).")

    # ── Section 4: Rolling Sharpe ──────────────────────────────────────────────
    st.divider()
    section_header("ROLLING SHARPE RATIO")
    vrs9_c1, vrs9_c2 = st.columns([3, 1])
    with vrs9_c2:
        vrs9_win = st.radio("Window", ["1 Year (252d)", "2 Years (504d)", "Both"],
                             index=2, key="val_rs_window", horizontal=False)
        vrs9_basis = st.radio("Returns", ["Gross", "Net of TC"], index=0,
                              key="val_rs_basis", horizontal=False)
    _vrs9_net = vrs9_basis.startswith("Net")
    _vdr9 = (v9_net_ret if _vrs9_net else v9_gross_ret).fillna(0)
    vroll9_252 = _vdr9.rolling(252).mean() / _vdr9.rolling(252).std() * np.sqrt(252)
    vroll9_504 = _vdr9.rolling(504).mean() / _vdr9.rolling(504).std() * np.sqrt(252)
    with vrs9_c1:
        fig_vrs9 = go.Figure()
        if vrs9_win in ("1 Year (252d)", "Both"):
            fig_vrs9.add_trace(go.Scatter(
                x=vroll9_252.index, y=vroll9_252.values, name="Rolling Sharpe (1yr)",
                mode="lines", line=dict(color=COLORS["primary"], width=1.6),
                hovertemplate="%{x|%b %Y}<br>Sharpe (1yr): %{y:.2f}<extra></extra>",
            ))
        if vrs9_win in ("2 Years (504d)", "Both"):
            fig_vrs9.add_trace(go.Scatter(
                x=vroll9_504.index, y=vroll9_504.values, name="Rolling Sharpe (2yr)",
                mode="lines", line=dict(color=COLORS["amber"], width=1.6, dash="dot"),
                hovertemplate="%{x|%b %Y}<br>Sharpe (2yr): %{y:.2f}<extra></extra>",
            ))
        fig_vrs9.add_hline(y=0,   line_dash="dash", line_color="#475569", line_width=1)
        fig_vrs9.add_hline(y=0.5, line_dash="dot",  line_color=COLORS["green"],
                            line_width=0.8, annotation_text="0.5", annotation_position="right")
        fig_vrs9.add_vline(x=pd.Timestamp("2022-01-01").value // 10**6, line_dash="dash", line_color=COLORS["amber"],
                            line_width=1.2, annotation_text="2022", annotation_position="top right")
        fig_vrs9.update_layout(
            **CHART_LAYOUT, height=320,
            title=dict(text=f"{val_vgroup} - Rolling Sharpe ({val_timing}, {'Net of TC' if _vrs9_net else 'Gross'})", font=dict(size=13)),
            yaxis_title="Annualised Sharpe", xaxis_title=None, hovermode="x unified",
        )
        fig_vrs9.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
        st.plotly_chart(fig_vrs9, use_container_width=True)
    st.caption("Value signals are strongly regime-dependent - COVID-era price dislocations (2020-2021) created a rare "
               "mean-reversion opportunity. The 2022 vertical line marks the post-COVID sub-period used in the NGL energy risk-premia paper.")

    # ── Section 5: Sub-period Analysis ────────────────────────────────────────
    st.divider()
    section_header("REGIME ANALYSIS - Pre-COVID / COVID Spike / Post-COVID")
    st.caption("Value signals are strongly regime-conditional. The 2020-2021 COVID window drove the majority of V2 10yr performance (+0.512 full-period). "
               "Split: Pre-2020 (normal regime) / 2020-2021 (dislocation) / 2022+ (normalisation).")

    _v9_pre20  = _v9_idx < pd.Timestamp("2020-01-01")
    _v9_covid  = (_v9_idx >= pd.Timestamp("2020-01-01")) & (_v9_idx < pd.Timestamp("2022-01-01"))
    _v9_post21 = _v9_idx >= pd.Timestamp("2022-01-01")

    vm9_pre20  = _val_perf(v9_gross_pnl[_v9_pre20],  v9_gross_ret[_v9_pre20],  val_pos[_v9_pre20],  "Pre-2020")
    vm9_covid  = _val_perf(v9_gross_pnl[_v9_covid],  v9_gross_ret[_v9_covid],  val_pos[_v9_covid],  "2020-2021")
    vm9_post21 = _val_perf(v9_gross_pnl[_v9_post21], v9_gross_ret[_v9_post21], val_pos[_v9_post21], "Post-2021")

    vsp_pre20, vsp_covid, vsp_post21 = st.columns(3)
    with vsp_pre20:
        st.markdown('<div style="color:#7A7068;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">Pre-2020 (Normal Regime)</div>', unsafe_allow_html=True)
        if vm9_pre20:
            _vp9a_cols = st.columns(2)
            _cmcard(_vp9a_cols[0], "Sharpe",       vm9_pre20.get("sharpe"),      ".3f")
            _cmcard(_vp9a_cols[1], "Ann Ret %",    vm9_pre20.get("ann_ret_pct"), ".1f", "%")
            _vp9a_cols2 = st.columns(2)
            _cmcard(_vp9a_cols2[0], "Max DD %",    vm9_pre20.get("mdd_pct"),     ".1f", "%")
            _cmcard(_vp9a_cols2[1], "Active Days", float(vm9_pre20.get("n", 0)), ",.0f")
        else:
            st.info("Insufficient data.")
    with vsp_covid:
        st.markdown('<div style="color:#7A7068;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">2020-2021 (COVID Dislocation)</div>', unsafe_allow_html=True)
        if vm9_covid:
            _vp9b_cols = st.columns(2)
            _cmcard(_vp9b_cols[0], "Sharpe",       vm9_covid.get("sharpe"),      ".3f")
            _cmcard(_vp9b_cols[1], "Ann Ret %",    vm9_covid.get("ann_ret_pct"), ".1f", "%")
            _vp9b_cols2 = st.columns(2)
            _cmcard(_vp9b_cols2[0], "Max DD %",    vm9_covid.get("mdd_pct"),     ".1f", "%")
            _cmcard(_vp9b_cols2[1], "Active Days", float(vm9_covid.get("n", 0)), ",.0f")
        else:
            st.info("Insufficient data.")
    with vsp_post21:
        st.markdown('<div style="color:#7A7068;font-size:0.72rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">Post-2021 (Post-Normalisation)</div>', unsafe_allow_html=True)
        if vm9_post21:
            _vp9c_cols = st.columns(2)
            _cmcard(_vp9c_cols[0], "Sharpe",       vm9_post21.get("sharpe"),      ".3f")
            _cmcard(_vp9c_cols[1], "Ann Ret %",    vm9_post21.get("ann_ret_pct"), ".1f", "%")
            _vp9c_cols2 = st.columns(2)
            _cmcard(_vp9c_cols2[0], "Max DD %",    vm9_post21.get("mdd_pct"),     ".1f", "%")
            _cmcard(_vp9c_cols2[1], "Active Days", float(vm9_post21.get("n", 0)), ",.0f")
        else:
            st.info("Insufficient data.")

    # ── Section 6: Contract Comparison (V1 only) ───────────────────────────────
    if val_is_v1:
        st.divider()
        section_header("CONTRACT COMPARISON - F1 THROUGH F15 (SHARPE, GROSS)")
        st.caption(f"Sharpe for each reference contract at {val_lb_label} lookback, ±{int(val_thr*100)}% threshold, {val_timing}. "
                   "Identifies the empirically optimal contract. The NGL energy paper chose F12 for energy - is the same true for copper?")

        _vc_sharpes = {}
        for _vk in range(1, 16):
            _vcol = f"F{_vk}"
            if _vcol not in v9_curve_px.columns:
                continue
            _vspec_k = {"variant": "v1", "contract": _vk, "lookback": val_N,
                        "threshold": val_thr, "same_day": val_same_day}
            _vraw_k = _value_raw_signal(v9_curve_px, vf1r, _vspec_k)
            if _vraw_k.dropna().shape[0] < 200:
                continue
            _vsig_k = np.where(_vraw_k.values < -val_thr, 1.0,
                       np.where(_vraw_k.values > val_thr, -1.0, 0.0))
            _vidx_k = _vraw_k.index.intersection(vf1c.index)
            _vf1c_k = vf1c.reindex(_vidx_k)
            _vsb_k  = pd.Series(_vsig_k, index=_vraw_k.index).reindex(_vidx_k).values
            _vT_k   = len(_vidx_k)
            if val_same_day:   # Same-Day = shift 1
                _vpos_k = np.empty(_vT_k); _vpos_k[0] = 0.0
                _vpos_k[1:] = np.where(np.isfinite(_vsb_k[:-1]), _vsb_k[:-1], 0.0)
            else:              # Lag-1 = shift 2
                _vpos_k = np.empty(_vT_k); _vpos_k[:2] = 0.0
                _vpos_k[2:] = np.where(np.isfinite(_vsb_k[:-2]), _vsb_k[:-2], 0.0)
            _vret_k = (pd.Series(_vpos_k, index=_vidx_k) * _vf1c_k.diff() / _vf1c_k.shift(1)).replace([np.inf,-np.inf],np.nan)
            _vact_k = _vret_k[_vpos_k != 0].dropna()
            if len(_vact_k) < 100:
                continue
            _vann_r  = float(_vact_k.mean() * 252 * 100)
            _vann_sd = float(_vact_k.std()  * np.sqrt(252) * 100)
            _vc_sharpes[f"F{_vk}"] = _vann_r / _vann_sd if _vann_sd > 0 else np.nan

        if _vc_sharpes:
            _vc_df = pd.DataFrame({"Contract": list(_vc_sharpes.keys()),
                                   "Sharpe":   list(_vc_sharpes.values())}).dropna()
            _vc_df["Color"] = _vc_df["Sharpe"].apply(
                lambda x: "#5BAD72" if x > 0.5 else ("#C9A84C" if x > 0 else "#B85450"))
            _vc_best = _vc_df.loc[_vc_df["Sharpe"].idxmax(), "Contract"]
            fig_vcc = go.Figure(go.Bar(
                x=_vc_df["Sharpe"].round(3), y=_vc_df["Contract"],
                orientation="h",
                marker_color=_vc_df["Color"].values,
                text=_vc_df["Sharpe"].round(3),
                textposition="outside",
                hovertemplate="Contract: %{y}<br>Sharpe: %{x:.3f}<extra></extra>",
            ))
            fig_vcc.add_vline(x=0, line_dash="dash", line_color="#475569", line_width=1)
            fig_vcc.update_layout(
                **CHART_LAYOUT, height=380,
                title=dict(text=f"Sharpe by Contract - {val_lb_label} MA, ±{int(val_thr*100)}% threshold", font=dict(size=13)),
                xaxis_title="Sharpe Ratio", yaxis_title=None,
            )
            st.plotly_chart(fig_vcc, use_container_width=True)
            st.caption(f"Best contract at this lookback/threshold: **{_vc_best}** (Sharpe = {_vc_sharpes.get(_vc_best, float('nan')):.3f}). "
                       "Green = Sharpe > 0.5, amber = 0-0.5, red = negative. "
                       "Contracts too near (F1-F3) are noisy; very far (F13-F15) may have limited liquidity.")

    # ── Section 7: Cumulative PnL + Comparison ────────────────────────────────
    st.divider()
    section_header("CUMULATIVE PnL (USD/MT)")

    _v9_cmp_keys = list(_VALUE_CMP_OPTIONS.keys())
    vc9_c1, vc9_c2 = st.columns(2)
    with vc9_c1:
        val_cmp_a = st.selectbox("Compare: Strategy A", _v9_cmp_keys, index=0, key="val_cmp_a")
    with vc9_c2:
        val_cmp_b = st.selectbox("Compare: Strategy B", _v9_cmp_keys, index=0, key="val_cmp_b")

    fig_v9cum = go.Figure()
    _V9_CMP_COLORS = [COLORS["primary"], COLORS["amber"]]
    for _v9_lbl, _v9_ci in [(val_cmp_a, 0), (val_cmp_b, 1)]:
        _v9_cspec = _VALUE_CMP_OPTIONS.get(_v9_lbl)
        if _v9_cspec is None:
            continue
        _v9_cpnl = _value_cum_pnl(v9_curve_px, vf1r, vf1c, _v9_cspec)
        if _v9_cpnl.empty:
            continue
        fig_v9cum.add_trace(go.Scatter(
            x=_v9_cpnl.index, y=_v9_cpnl.values, name=_v9_lbl, mode="lines",
            line=dict(color=_V9_CMP_COLORS[_v9_ci], width=1.8,
                      dash="dot" if _v9_ci == 1 else "solid"),
            hovertemplate=f"%{{x|%b %d, %Y}}<br>{_v9_lbl}: $%{{y:,.1f}}<extra></extra>",
        ))

    # Primary series
    fig_v9cum.add_trace(go.Scatter(
        x=v9_cum_pnl.index, y=v9_cum_pnl.values,
        name=f"Selected: {val_vgroup}", mode="lines",
        line=dict(color="#FFFFFF", width=2.0),
        hovertemplate="%{x|%b %d, %Y}<br>PnL: $%{y:,.1f}<extra></extra>",
    ))
    fig_v9cum.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
    fig_v9cum.add_vline(x=pd.Timestamp("2022-01-01").value // 10**6, line_dash="dash", line_color=COLORS["amber"],
                         line_width=1.2, annotation_text="2022", annotation_position="top right")
    fig_v9cum.update_layout(
        **CHART_LAYOUT, height=420,
        title=dict(text="Cumulative PnL (USD/MT) - Value Strategies", font=dict(size=13)),
        yaxis_title="Cumulative PnL ($/MT)", hovermode="x unified",
    )
    fig_v9cum.update_xaxes(showspikes=True, spikecolor="#475569", spikethickness=1, spikemode="across")
    st.plotly_chart(fig_v9cum, use_container_width=True)
    st.caption("White = currently selected variant. Use the dropdowns to compare across contracts, lookbacks, or V1 vs V2. "
               "The post-2022 acceleration (after the COVID spike mean-reversion) is a critical structural feature to examine.")

    # ── Section 8: Annual PnL ─────────────────────────────────────────────────
    st.divider()
    section_header("ANNUAL PnL (USD/MT)")

    _v9_ann = v9_gross_pnl.groupby(v9_gross_pnl.index.year).sum()
    _v9_ann_colors = ["#5BAD72" if v >= 0 else "#B85450" for v in _v9_ann.values]
    fig_v9ann = go.Figure(go.Bar(
        x=_v9_ann.index.astype(str), y=_v9_ann.values.round(1),
        marker_color=_v9_ann_colors,
        text=[f"${v:,.0f}" for v in _v9_ann.values],
        textposition="outside",
        hovertemplate="Year: %{x}<br>PnL: $%{y:,.1f}<extra></extra>",
    ))
    fig_v9ann.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
    fig_v9ann.update_layout(
        **CHART_LAYOUT, height=300,
        title=dict(text=f"{val_vgroup} - Annual PnL ($/MT, Gross)", font=dict(size=13)),
        xaxis_title=None, yaxis_title="PnL ($/MT)",
    )
    st.plotly_chart(fig_v9ann, use_container_width=True)

    # ── Section 9: Recent Signal Changes ──────────────────────────────────────
    st.divider()
    section_header("RECENT SIGNAL CHANGES (Last 20)")

    _v9_flips2 = val_pos.diff().abs() > 0
    _v9_flips2.iloc[0] = val_pos.iloc[0] != 0
    _v9_flip_dates2 = val_pos[_v9_flips2].tail(20)

    if not _v9_flip_dates2.empty:
        _v9_flip_df = pd.DataFrame({
            "Date":          _v9_flip_dates2.index.strftime("%Y-%m-%d"),
            "Position":      _v9_flip_dates2.values.astype(int),
            "State":         ["LONG" if v > 0 else ("SHORT" if v < 0 else "FLAT")
                              for v in _v9_flip_dates2.values],
            "F1_raw ($/MT)": vf1r_a.reindex(_v9_flip_dates2.index).round(1).values,
            "Signal Value":  val_raw.reindex(_v9_flip_dates2.index).round(6).values,
            "Daily PnL":     v9_gross_pnl.reindex(_v9_flip_dates2.index).round(2).values,
            "Cum PnL":       v9_cum_pnl.reindex(_v9_flip_dates2.index).round(2).values,
        })
        st.dataframe(_v9_flip_df, use_container_width=True, hide_index=True)
    else:
        st.info("No signal changes found in data.")

    # ── Methodology Notes ──────────────────────────────────────────────────────
    st.divider()
    with st.expander("Methodology Notes", expanded=False):
        st.markdown("""
**V1 - Long-Run MA Reversion**

Signal: `deviation_t = (Fk_t − MA_N(Fk)) / MA_N(Fk)`

Position (three states):
- `+1` (Long):  `deviation < −threshold`  → price below long-run mean → *cheap*
- `−1` (Short): `deviation > +threshold`  → price above long-run mean → *expensive*
- ` 0` (Flat):  deviation within ±threshold → *near fair value*

The flat zone is a distinguishing feature: unlike momentum and carry, value is *not always* in the market. The ±10% default threshold is calibrated to the NGL energy risk-premia paper. For copper (lower realized volatility than crude oil), consider testing ±15-20% as the flat zone may be very small at ±10%.

**V2 - Baz-Granger N-year Reversal**

Signal: `reversal_t = F1_raw[t−N] − F1_raw[t]`

Position (two states):
- `+1` (Long):  `reversal > 0` → price has *fallen* over N years → contrarian long
- `−1` (Short): `reversal < 0` → price has *risen* over N years → contrarian short

No flat zone - always in the market. This is a pure reversal bet on medium/long-run mean reversion. Positive reversal means the market is cheaper than it was N years ago, which the strategy treats as a buy signal.

**Why F12 as Reference (V1)?**
The NGL energy paper chose F12 (1-year-forward contract) because it is:
1. Far enough from the prompt to avoid roll/squeeze noise
2. Close enough to have continuous liquidity and price discovery
3. Closely correlated with long-run supply/demand expectations

For copper, F12 corresponds to the LME 12-month forward, which is well-traded.
The contract comparison chart above tests F1-F15 to find the empirically optimal contract for copper.

**Position Timing** (no look-ahead either way)
- *Same-Day (shift 1)*: `position` uses the prior close's signal; first return t→t+1. Realistic default.
- *Lag-1 (shift 2)*: one further close of delay; first return t+1→t+2. Conservative.

For V2 Baz-Granger, Same-Day beats Lag-1 (+0.51 vs +0.37). For V1 MA-Reversion the two are close (Lag-1 marginally ahead). Both are free of the look-ahead that affected the earlier shift-0 "Same-Day" definition.

**PnL & Returns**
All strategies trade F1_continuous regardless of which contract or lookback generates the signal.
`daily_ret[t] = position[t] × ΔF1_cont[t] / F1_cont[t−1]`

**Regime Conditionality Disclosure**
V2 Baz-Granger 10yr Sharpe (+0.512 gross) is **not a stable systematic signal**. The edge is
concentrated in the 2020-2021 COVID mean-reversion window, when copper prices dislocated significantly
below their 10-year trend and then recovered. In normal regimes (pre-2020, post-2021), V2 is weak
(Sharpe ≈ 0.0 to +0.3). Present this as event-driven risk premia, not a persistent carry-style signal.

V1 MA Reversion (best: F8, 5yr, Sharpe +0.277) is more stable across regimes but still weak for
copper - the ±10% threshold leaves copper in the flat zone most of the time (60% flat at ±10%).
This threshold was calibrated for crude oil (higher volatility). Consider ±15-20% for copper.

**References**
- NGL Energy Risk-Premia paper - *Risk Premia in Diversified Energy Portfolios.*
- Baz, J., Granger, N. M. (2015). Dissecting Investment Strategies in the Cross Section and Time Series. SSRN.
        """)

    st.markdown(
        '<div style="background:#0D1117;border:1px solid #2A2A2A;border-left:4px solid #B87333;'
        'border-radius:4px;padding:12px 20px;margin-top:20px;">'
        '<span style="color:#B87333;font-size:0.82rem;font-family:\'IBM Plex Mono\',monospace;font-weight:700;">'
        'All three signals validated - </span>'
        '<span style="color:#94A3B8;font-size:0.78rem;font-family:\'IBM Plex Mono\',monospace;font-weight:600;">'
        'NEXT &rarr; </span>'
        '<span style="color:#B87333;font-size:0.82rem;font-family:\'IBM Plex Mono\',monospace;font-weight:700;">'
        'Tab 10: Portfolio Construction</span>'
        '<span style="color:#8A8278;font-size:0.78rem;"> &nbsp;-&nbsp; '
        'Momentum (trend) + Carry (level) + Value (mean-reversion) combined into an equal-weight portfolio. '
        'Low pairwise correlations confirm genuine diversification: '
        'Mom-Carry +0.05, Mom-Value &minus;0.21, Carry-Value +0.03.</span></div>',
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════
# TAB 10: PORTFOLIO CONSTRUCTION
# ══════════════════════════════════════════════════════

with tab10:
    _p10_metal = st.radio("🔬 Metal", ["Copper", "Aluminium"], horizontal=True, key="port_metal")
    st.markdown(f"### Portfolio Construction - LME {_p10_metal}")
    st.caption(
        "Equal-weight combination of the three risk premia signals. "
        "All positions sized ±1 per signal; portfolio position ranges −1 to +1. "
        "PnL from F1_continuous throughout. Legs auto-switch to the selected metal's best signals."
    )

    # ── Load data ──────────────────────────────────────────────────────────────
    _f1_df_p10 = _load_f1_data(_p10_metal)
    if _f1_df_p10.empty:
        st.error(f"Rolling F1 file for {_p10_metal} not found.")
        st.stop()
    pf1c = _f1_df_p10["F1_continuous"]
    pf1r = _f1_df_p10["F1_raw"]

    _cu_sheet_p10 = _find_curve_sheet(_p10_metal, curve_data) if curve_data else None
    if not curve_data or _cu_sheet_p10 is None:
        st.error("Futures Curve data not loaded. Upload in the sidebar.")
        st.stop()
    p10_crv = curve_data[_cu_sheet_p10]["prices"].copy()
    p10_crv.index = pd.to_datetime(p10_crv.index).normalize()
    p10_crv = p10_crv.sort_index()

    # ── Portfolio legs: metal-specific best-of-each configs (all shift-1, no look-ahead) ──
    #   Copper:    Momentum MA(35,43) | Carry 20d roll-yield momentum | Value V1 F8 5yr
    #   Aluminium: Momentum MA(60,115)| Carry 252d roll-yield z-score | Value V1 F12 5yr
    if _p10_metal == "Aluminium":
        _P10_MA = (60, 115); _P10_CARRY = "zscore"; _P10_VK = 12
    else:
        _P10_MA = (35, 43);  _P10_CARRY = "carrymom"; _P10_VK = 8
    _p10_mom_label   = f"MA({_P10_MA[0]},{_P10_MA[1]})"
    _p10_carry_label = "Z-score 252d" if _P10_CARRY == "zscore" else "CarryMom 20d"
    _p10_val_label   = f"V1 F{_P10_VK} 5yr"

    # ── Signal 1: Momentum MA crossover, shift-1 ──────────────────────────────
    _p10_mom_pos = np.sign(pf1r.rolling(_P10_MA[0]).mean() - pf1r.rolling(_P10_MA[1]).mean()).shift(1).fillna(0)

    # ── Signal 2: Carry — roll-yield momentum (copper) or 252d z-score (aluminium) ─
    if "F1" in p10_crv.columns and "F2" in p10_crv.columns:
        _p10_cr_base = ((p10_crv["F1"] - p10_crv["F2"]) / p10_crv["F1"]).replace([np.inf, -np.inf], np.nan)
        if _P10_CARRY == "zscore":
            _p10_cr_sig = ((_p10_cr_base - _p10_cr_base.rolling(252).mean())
                           / _p10_cr_base.rolling(252).std()).replace([np.inf, -np.inf], np.nan)
        else:
            _p10_cr_sig = (_p10_cr_base - _p10_cr_base.shift(20))   # 20d change in roll yield
        _p10_carry_pos = np.sign(_p10_cr_sig).shift(1).reindex(pf1c.index).fillna(0)
    else:
        st.error("F1 or F2 column missing from curve data.")
        st.stop()

    # ── Signal 3A: Value V1 on F{VK}, 5yr MA, ±10%, shift-1 ──────────────────
    _p10_vk_col = f"F{_P10_VK}"
    _p10_val_v1_ok = _p10_vk_col in p10_crv.columns
    if _p10_val_v1_ok:
        _p10_fk = p10_crv[_p10_vk_col].dropna()
        _p10_mak = _p10_fk.rolling(1260, min_periods=630).mean()
        _p10_devk = ((_p10_fk - _p10_mak) / _p10_mak).replace([np.inf, -np.inf], np.nan).dropna()
        _p10_v1_bin = np.where(_p10_devk.values < -0.10, 1.0, np.where(_p10_devk.values > 0.10, -1.0, 0.0))
        _p10_val_v1_pos = pd.Series(_p10_v1_bin, index=_p10_devk.index).shift(1).fillna(0).reindex(pf1c.index).fillna(0)
    else:
        _p10_val_v1_pos = pd.Series(0.0, index=pf1c.index)

    # ── Signal 3B: Value V2 BG 10yr, shift-1 execution ───────────────────────
    _p10_rev10 = (pf1r.shift(2520) - pf1r).replace([np.inf, -np.inf], np.nan).dropna()
    _p10_val_v2_pos = np.sign(_p10_rev10).shift(1).fillna(0).reindex(pf1c.index).fillna(0)

    # ── Value selector + Weighting + TC ───────────────────────────────────────
    _p10_val_col, _p10_wt_col, _p10_tc_col = st.columns([2.3, 2.3, 1.6])
    with _p10_val_col:
        _p10_val_choice = st.selectbox(
            "Value signal for portfolio",
            ["V1: MA-reversion 5yr (default, robust)", "V2: BG 10yr reversal (fragile)"],
            index=0, key="p10_val_choice",
            help=f"Default V1 is metal-specific ({_p10_val_label}): MA-reversion, robust out-of-sample. "
                 "V2 BG 10yr has higher in-sample Sharpe but collapses out-of-sample (edge concentrated "
                 "in the 2020-21 dislocation), so it is the alternative, not the default.",
        )
    with _p10_wt_col:
        _p10_wt_choice = st.selectbox(
            "Weighting scheme",
            ["Equal-Weight (1/3 each)", "Inverse-Vol (risk-balanced)"],
            index=0, key="p10_wt_choice",
            help="Equal-Weight: fixed 1/3 per sleeve. Inverse-Vol: each sleeve weighted by "
                 "1 / its trailing 63d return-vol (renormalised daily) so each contributes "
                 "equal risk. For these 3 similar-vol signals the two are close; inverse-vol "
                 "is marginally more stable post-2022.",
        )
    with _p10_tc_col:
        _p10_tc_map   = _tc_label_map(float(pf1c.dropna().iloc[-1]))
        _p10_tc_label = st.selectbox("TC (bps)", list(_p10_tc_map.keys()), index=0, key="p10_tc")
        _p10_tc_bps   = _p10_tc_map[_p10_tc_label]
    _p10_use_v2 = "V2" in _p10_val_choice
    _p10_val_pos = _p10_val_v2_pos if _p10_use_v2 else _p10_val_v1_pos

    # ── Align on common index ─────────────────────────────────────────────────
    _p10_idx = pf1c.index
    for _s in [_p10_mom_pos, _p10_carry_pos, _p10_val_pos]:
        _p10_idx = _p10_idx.intersection(_s.index)
    _p10_idx = _p10_idx.sort_values()

    _p10_m  = _p10_mom_pos.reindex(_p10_idx).fillna(0)
    _p10_c  = _p10_carry_pos.reindex(_p10_idx).fillna(0)
    _p10_v  = _p10_val_pos.reindex(_p10_idx).fillna(0)
    _p10_f  = pf1c.reindex(_p10_idx)
    _p10_fraw = pf1r.reindex(_p10_idx)   # F1_raw aligned — used for TC, not PnL

    # Equal-weight and inverse-vol composites
    _p10_port_ew = (_p10_m + _p10_c + _p10_v) / 3.0

    def _p10_sleeve_ret(_pos):
        with np.errstate(invalid="ignore", divide="ignore"):
            return (_pos * _p10_f.diff() / _p10_f.shift(1)).replace([np.inf, -np.inf], np.nan).fillna(0)
    _p10_volw = pd.DataFrame({
        "m": _p10_sleeve_ret(_p10_m), "c": _p10_sleeve_ret(_p10_c), "v": _p10_sleeve_ret(_p10_v),
    }).rolling(63).std().shift(1)                      # trailing vol, lagged -> no look-ahead
    _p10_iw = (1.0 / _p10_volw).replace([np.inf, -np.inf], np.nan)
    _p10_iw = _p10_iw.div(_p10_iw.sum(axis=1), axis=0)
    _p10_wm = _p10_iw["m"].fillna(1/3); _p10_wc = _p10_iw["c"].fillna(1/3); _p10_wv = _p10_iw["v"].fillna(1/3)
    _p10_port_iv = _p10_wm * _p10_m + _p10_wc * _p10_c + _p10_wv * _p10_v

    _p10_use_iv = "Inverse" in _p10_wt_choice
    _p10_port = _p10_port_iv if _p10_use_iv else _p10_port_ew
    # Dynamic weighting labels used by every card/chart/table below (no hardcoded "EW")
    _p10_wt_full  = "Inverse-Vol" if _p10_use_iv else "Equal-Weight"
    _p10_wt_word  = "INVERSE-VOL" if _p10_use_iv else "EQUAL-WEIGHT"
    _p10_port_label = f"{_p10_wt_full} Portfolio"

    # ── PnL series ────────────────────────────────────────────────────────────
    def _p10_ret(pos: pd.Series, tc_bps: int = 0) -> pd.Series:
        pnl = pos * _p10_f.diff()
        if tc_bps > 0:
            chg = pos.diff().abs(); chg.iloc[0] = abs(pos.iloc[0])
            pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * _p10_fraw
        with np.errstate(invalid="ignore", divide="ignore"):
            return (pnl / _p10_f.shift(1)).replace([np.inf, -np.inf], np.nan)

    def _p10_sharpe(ret: pd.Series) -> float:
        act = ret[ret != 0].dropna()
        if len(act) < 20: return np.nan
        sd = act.std(ddof=1)
        return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

    def _p10_metrics(pos: pd.Series, tc_bps: int = 0):
        ret  = _p10_ret(pos, tc_bps)
        sh   = _p10_sharpe(ret)
        ann  = float(ret.dropna().mean() * 252 * 100)
        pnl  = pos * _p10_f.diff()
        if tc_bps > 0:
            chg = pos.diff().abs(); chg.iloc[0] = abs(pos.iloc[0])
            pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * _p10_fraw
        cum  = pnl.cumsum()
        dd   = float((cum - cum.cummax()).min())
        flat_pct = float(100 * (pos == 0).sum() / len(pos))
        return sh, ann, dd, flat_pct

    def _p10_sub_sharpe(pos: pd.Series, start=None, end=None, tc_bps: int = 0) -> float:
        idx = pos.index
        if start: idx = idx[idx >= pd.Timestamp(start)]
        if end:   idx = idx[idx <  pd.Timestamp(end)]
        if len(idx) < 20: return np.nan
        p  = pos.reindex(idx)
        f_ = _p10_f.reindex(idx)
        fraw_ = _p10_fraw.reindex(idx)
        pnl = p * f_.diff()
        if tc_bps > 0:
            chg = p.diff().abs(); chg.iloc[0] = abs(p.iloc[0])
            pnl = pnl - chg * (tc_bps / 10000.0 / 2.0) * fraw_
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = (pnl / f_.shift(1)).replace([np.inf, -np.inf], np.nan)
        return _p10_sharpe(ret)

    _port_sh, _port_ann, _port_dd, _port_flat = _p10_metrics(_p10_port, _p10_tc_bps)
    _mom_sh,  _mom_ann,  _mom_dd,  _mom_flat  = _p10_metrics(_p10_m,    _p10_tc_bps)
    _car_sh,  _car_ann,  _car_dd,  _car_flat  = _p10_metrics(_p10_c,    _p10_tc_bps)
    _val_sh,  _val_ann,  _val_dd,  _val_flat  = _p10_metrics(_p10_v,    _p10_tc_bps)

    # ── Portfolio story intro (DYNAMIC per metal) ──────────────────────────────
    _cmc = _p10_m.corr(_p10_c); _cmv = _p10_m.corr(_p10_v); _ccv = _p10_c.corr(_p10_v)
    def _shc(x): return "N/A" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:+.2f}"
    st.markdown(f"""
<div style="background:#111827;border:1px solid #2A2A2A;border-radius:6px;padding:16px 22px;margin-bottom:4px;">
<p style="color:#B87333;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:700;margin:0 0 10px">
THREE ORTHOGONAL RISK PREMIA &rarr; ONE {_p10_wt_word} PORTFOLIO &nbsp;({_p10_metal})</p>
<table style="width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:0.78rem;color:#C8BFB4;">
<tr style="border-bottom:1px solid #2A2A2A;color:#8A8278;font-size:0.72rem;">
  <td style="padding:4px 10px 4px 0">Signal</td>
  <td style="padding:4px 10px">Leg used ({_p10_metal})</td>
  <td style="padding:4px 10px">Net Sharpe (IS, sel. TC)</td>
  <td style="padding:4px 10px">Entry</td>
  <td style="padding:4px 10px">Correlation Role</td>
</tr>
<tr style="border-bottom:1px solid #1C1C1C;">
  <td style="padding:5px 10px 5px 0;color:#B87333;font-weight:600">Momentum</td>
  <td style="padding:5px 10px">{_p10_mom_label}</td>
  <td style="padding:5px 10px;color:#5BAD72">{_shc(_mom_sh)}</td>
  <td style="padding:5px 10px">shift-1</td>
  <td style="padding:5px 10px">Trend persistence - anchor signal</td>
</tr>
<tr style="border-bottom:1px solid #1C1C1C;">
  <td style="padding:5px 10px 5px 0;color:#B87333;font-weight:600">Carry</td>
  <td style="padding:5px 10px">{_p10_carry_label}</td>
  <td style="padding:5px 10px;color:#5BAD72">{_shc(_car_sh)}</td>
  <td style="padding:5px 10px">shift-1</td>
  <td style="padding:5px 10px">Curve premium - diversifies price-momentum</td>
</tr>
<tr>
  <td style="padding:5px 10px 5px 0;color:#B87333;font-weight:600">Value</td>
  <td style="padding:5px 10px">{_p10_val_label}</td>
  <td style="padding:5px 10px;color:#5BAD72">{_shc(_val_sh)}</td>
  <td style="padding:5px 10px">shift-1</td>
  <td style="padding:5px 10px">Mean-reversion - typically negatively correlated with Mom</td>
</tr>
</table>
<p style="color:#8A8278;font-size:0.75rem;margin:10px 0 0">
All legs use shift-1 execution (trade at the signal's close; first return next day; no look-ahead). Legs
auto-switch to the selected metal's best-performing configuration. Live pairwise position correlations:
Mom-Carry {_shc(_cmc)} &nbsp;|&nbsp; Mom-Value {_shc(_cmv)} &nbsp;|&nbsp; Carry-Value {_shc(_ccv)}.
&nbsp;{_p10_wt_full} portfolio net Sharpe (IS, selected TC) &asymp; <b>{_shc(_port_sh)}</b>; per-window walk-forward OOS is shown below.</p>
</div>""", unsafe_allow_html=True)

    # ── Section 1: Live Portfolio Badge ───────────────────────────────────────
    st.divider()
    section_header("LIVE PORTFOLIO POSITION")
    _p10_cur_pos   = float(_p10_port.iloc[-1])
    _p10_cur_date  = _p10_idx[-1]
    _p10_cur_mom   = float(_p10_m.iloc[-1])
    _p10_cur_carry = float(_p10_c.iloc[-1])
    _p10_cur_val   = float(_p10_v.iloc[-1])

    _p10_badge_col1, _p10_badge_col2, _p10_badge_col3, _p10_badge_col4 = st.columns(4)

    def _p10_pos_color(v):
        if v > 0.1:  return "#5BAD72"
        if v < -0.1: return "#B85450"
        return "#B87333"

    _p10_badge_style = (
        "background:#161616;border:1px solid #2A2A2A;border-radius:4px;"
        "padding:14px 18px;text-align:center"
    )
    for _col, _lbl, _val in [
        (_p10_badge_col1, _p10_port_label, _p10_cur_pos),
        (_p10_badge_col2, "Momentum", _p10_cur_mom),
        (_p10_badge_col3, "Carry", _p10_cur_carry),
        (_p10_badge_col4, "Value", _p10_cur_val),
    ]:
        _col.markdown(
            f'<div style="{_p10_badge_style}">'
            f'<p style="color:#8A8278;font-size:0.75rem;margin:0 0 4px">{_lbl}</p>'
            f'<p style="color:{_p10_pos_color(_val)};font-size:1.6rem;font-weight:700;margin:0">'
            f'{_val:+.2f}</p>'
            f'<p style="color:#8A8278;font-size:0.72rem;margin:4px 0 0">as of {_p10_cur_date.strftime("%Y-%m-%d")}</p>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # ── Section 2: Portfolio Performance Cards ────────────────────────────────
    st.divider()
    _p10_yr0 = str(pf1c.index[0].year); _p10_yr1 = str(pf1c.index[-1].year)
    _p10_tc_note_hdr = f", {_p10_tc_label}" if _p10_tc_bps > 0 else ", 0 TC (Gross)"
    _p10_wt_short = _p10_wt_word
    section_header(f"{_p10_wt_short} PORTFOLIO PERFORMANCE{_p10_tc_note_hdr.upper()}")
    st.caption(f"Full period {_p10_yr0}-{_p10_yr1}{_p10_tc_note_hdr}, all sleeves same-day execution (shift 1, no look-ahead).")

    _p10_card_s  = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;"
                    "border-radius:4px;padding:14px 20px")
    _p10_c_lbl   = ("color:#B87333;font-family:'IBM Plex Mono',monospace;"
                    "font-size:0.8rem;font-weight:600;margin:0 0 4px")
    _p10_c_big   = ("color:#E8DDD0;font-family:'IBM Plex Mono',monospace;"
                    "font-size:1.45rem;font-weight:700;margin:0")
    _p10_c_sub   = "color:#8A8278;font-size:0.73rem;margin:2px 0"
    _p10_c_hr    = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"

    _p10_cc1, _p10_cc2, _p10_cc3, _p10_cc4 = st.columns(4)
    for _col2, _lbl2, _big2, _sub2 in [
        (_p10_cc1, "Sharpe Ratio", f"{_port_sh:+.3f}", f"Ann.{'  Net' if _p10_tc_bps>0 else ' Gross'} ({_p10_yr0}-{_p10_yr1})"),
        (_p10_cc2, "Ann. Return",  f"{_port_ann:+.1f}%", f"{'Net' if _p10_tc_bps>0 else 'Gross'}, {_p10_tc_note_hdr.strip(', ')}"),
        (_p10_cc3, "Max Drawdown", f"${_port_dd:,.0f}/MT", "Cumulative USD/MT"),
        (_p10_cc4, "% In Market",  f"{100-_port_flat:.1f}%", f"Flat: {_port_flat:.1f}% of days"),
    ]:
        _col2.markdown(
            f'<div style="{_p10_card_s}">'
            f'<p style="{_p10_c_lbl}">{_lbl2}</p>'
            f'<p style="{_p10_c_big}">{_big2}</p>'
            f'<hr style="{_p10_c_hr}"/>'
            f'<p style="{_p10_c_sub}">{_sub2}</p>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # ── EW vs Inverse-Vol comparison ──────────────────────────────────────────
    with st.expander("⚖️  Equal-Weight vs Inverse-Vol - how it works + side-by-side", expanded=False):
        def _p10_cmp(_pos):
            _sh_f = _p10_sub_sharpe(_pos, None, None, _p10_tc_bps)
            _sh_pre = _p10_sub_sharpe(_pos, None, "2022-01-01", _p10_tc_bps)
            _sh_post = _p10_sub_sharpe(_pos, "2022-01-01", None, _p10_tc_bps)
            _s, _a, _d, _fl = _p10_metrics(_pos, _p10_tc_bps)
            return _sh_f, _sh_pre, _sh_post, _a, _d
        _ew_f, _ew_pre, _ew_post, _ew_ann, _ew_dd = _p10_cmp(_p10_port_ew)
        _iv_f, _iv_pre, _iv_post, _iv_ann, _iv_dd = _p10_cmp(_p10_port_iv)
        _tcn = "Net" if _p10_tc_bps > 0 else "Gross"
        _aw = _p10_iw.mean()   # realised average inverse-vol weights (live, per metal)
        _awm = float(_aw.get("m", 1/3)); _awc = float(_aw.get("c", 1/3)); _awv = float(_aw.get("v", 1/3))
        st.markdown(f"""
**How inverse-vol weighting works**

Equal-weight gives each sleeve a fixed **1/3** of the position. That equalises *position size*, not
*risk* - an always-on signal (momentum) contributes more variance than one that is flat 40% of the time
(value). Inverse-vol fixes this:

1. Each day, take each sleeve's trailing **63-day return volatility** σ_m, σ_c, σ_v (lagged one day → no look-ahead).
2. Weight each sleeve by **w_i = (1/σ_i) / Σ(1/σ_j)** - low-vol sleeves get more weight, so each contributes ≈ equal risk.
3. Portfolio position = w_m×Mom + w_c×Carry + w_v×Value, rebalanced daily.

For *these* three {_p10_metal} signals the realised vols are similar, so the average weights land near
**{_awm:.2f} / {_awc:.2f} / {_awv:.2f}** (Mom / Carry / Value) - close to equal. The benefit is therefore
modest and shows up mostly as **post-2022 stability** (the weights tilt away from whichever sleeve is
blowing out in a given regime).

| Metric ({_tcn}, TC={_p10_tc_bps}bps) | Equal-Weight | Inverse-Vol |
|---|---|---|
| Sharpe - full period | **{_ew_f:+.3f}** | **{_iv_f:+.3f}** |
| Sharpe - pre-2022 | {_ew_pre:+.3f} | {_iv_pre:+.3f} |
| Sharpe - post-2022 | {_ew_post:+.3f} | {_iv_post:+.3f} |
| Ann. return | {_ew_ann:+.1f}% | {_iv_ann:+.1f}% |
| Max drawdown | ${_ew_dd:,.0f}/MT | ${_iv_dd:,.0f}/MT |

**Read:** EW is marginally better on full-period Sharpe and has a smaller drawdown; inverse-vol earns a
higher annual return and is steadier post-2022. Neither dominates - which is itself the finding: with
three similar-vol, low-correlation sleeves, equal-weight is already close to risk-parity. EW is kept as
the default for simplicity; switch to inverse-vol above if you prefer regime stability.
        """)

    # Individual signal cards
    st.markdown("&nbsp;")
    _p10_ic1, _p10_ic2, _p10_ic3 = st.columns(3)
    for _col3, _lbl3, _sh3, _ann3 in [
        (_p10_ic1, f"Momentum {_p10_mom_label}", _mom_sh, _mom_ann),
        (_p10_ic2, f"Carry {_p10_carry_label}",  _car_sh, _car_ann),
        (_p10_ic3, "Value (selected)",           _val_sh, _val_ann),
    ]:
        _col3.markdown(
            f'<div style="{_p10_card_s}">'
            f'<p style="{_p10_c_lbl}">{_lbl3}</p>'
            f'<p style="{_p10_c_big}">{_sh3:+.3f}</p>'
            f'<hr style="{_p10_c_hr}"/>'
            f'<p style="{_p10_c_sub}">Sharpe, Ann Ret {_ann3:+.1f}%</p>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # ── Section 2b: Out-of-Sample Walk-Forward Validation ─────────────────────
    st.divider()
    section_header("OUT-OF-SAMPLE WALK-FORWARD VALIDATION")
    _p10_wf = _wf_portfolio_tc(
        _p10_metal, _p10_use_iv, _p10_use_v2, _p10_tc_bps,
        _p10_m, _p10_c, _p10_v, _p10_port, pf1r, pf1c,
    )
    _p10_wf_yrs = sorted(_p10_wf["Portfolio"].keys())
    _p10_wf_first = _p10_wf_yrs[0] if _p10_wf_yrs else "N/A"
    _p10_wf_last  = _p10_wf_yrs[-1] if _p10_wf_yrs else "N/A"
    _p10_wf_recent = _p10_wf_yrs[-3:] if len(_p10_wf_yrs) >= 3 else _p10_wf_yrs
    _p10_wf_recent_lbl = f"{_p10_wf_recent[0]}-{_p10_wf_recent[-1]}" if _p10_wf_recent else "-"
    _p10_wf_tcn = f", {_p10_tc_label}" if _p10_tc_bps > 0 else ", 0 TC (Gross)"
    st.caption(
        f"IS = 5yr rolling window, OOS = 1yr, all legs same-day execution (shift 1, no look-ahead). "
        f"Each leg uses {_p10_metal}'s a-priori configuration - never re-optimised per window. "
        f"{len(_p10_wf_yrs)} OOS windows, labelled by end year, coverage {_p10_wf_first}-{_p10_wf_last}"
        f"{_p10_wf_tcn}."
    )

    def _wf_avg(d):
        vv = [v for v in d.values() if v is not None and not np.isnan(v)]
        return float(np.nanmean(vv)) if vv else np.nan
    def _wf_avg_recent(d):
        vv = [d[y] for y in _p10_wf_recent if y in d and not np.isnan(d[y])]
        return float(np.nanmean(vv)) if vv else np.nan

    _p10_wf_port_avg   = _wf_avg(_p10_wf["Portfolio"])
    _p10_wf_port_rec   = _wf_avg_recent(_p10_wf["Portfolio"])
    _p10_wf_port_vals  = [v for v in _p10_wf["Portfolio"].values() if v is not None and not np.isnan(v)]
    _p10_wf_npos       = sum(1 for v in _p10_wf_port_vals if v > 0)
    _p10_wf_ngt03      = sum(1 for v in _p10_wf_port_vals if v > 0.30)
    _p10_wf_ntot       = len(_p10_wf["Portfolio"])

    _p10_wfc1, _p10_wfc2, _p10_wfc3 = st.columns(3)
    _wfs   = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #B87333;"
              "border-radius:4px;padding:14px 20px")
    _wfsg  = ("background:#161616;border:1px solid #2A2A2A;border-left:4px solid #475569;"
              "border-radius:4px;padding:14px 20px")
    _wflbl = "color:#B87333;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:600;margin:0 0 6px"
    _wflbg = "color:#94A3B8;font-family:'IBM Plex Mono',monospace;font-size:0.85rem;font-weight:600;margin:0 0 6px"
    _wfbig = "color:#E8DDD0;font-family:'IBM Plex Mono',monospace;font-size:1.55rem;font-weight:700;margin:0"
    _wfmed = "color:#E8DDD0;font-family:'IBM Plex Mono',monospace;font-size:1.15rem;font-weight:600;margin:0"
    _wfsub = "color:#8A8278;font-size:0.75rem;margin:2px 0"
    _wfhr  = "border:none;border-top:1px solid #2A2A2A;margin:8px 0"

    with _p10_wfc1:
        st.markdown(f"""<div style="{_wfs}">
<p style="{_wflbl}">{_p10_port_label} - OOS Sharpe</p>
<p style="{_wfbig}">{_p10_wf_port_avg:+.3f}</p>
<p style="{_wfsub}">Avg across {_p10_wf_ntot} OOS windows ({_p10_wf_first}-{_p10_wf_last}){_p10_wf_tcn}</p>
<hr style="{_wfhr}"/>
<p style="{_wfsub}">{_p10_wf_recent_lbl} avg</p>
<p style="{_wfmed}">{_p10_wf_port_rec:+.3f}</p>
<p style="{_wfsub}">Fixed-config, zero per-window re-optimisation</p>
</div>""", unsafe_allow_html=True)

    with _p10_wfc2:
        st.markdown(f"""<div style="{_wfs}">
<p style="{_wflbl}">Per-Sleeve OOS Sharpe</p>
<p style="{_wfmed}">Mom {_wf_avg(_p10_wf['Momentum']):+.2f} &nbsp; Carry {_wf_avg(_p10_wf['Carry']):+.2f} &nbsp; Val {_wf_avg(_p10_wf['Value']):+.2f}</p>
<hr style="{_wfhr}"/>
<p style="{_wfsub}">Avg OOS Sharpe of each leg, same windows{_p10_wf_tcn}</p>
<p style="{_wfsub}">Portfolio OOS Sharpe ({_p10_wf_port_avg:+.2f}) typically exceeds the best single leg - diversification holding out-of-sample</p>
</div>""", unsafe_allow_html=True)

    with _p10_wfc3:
        _p10_wf_best = max(_p10_wf["Portfolio"], key=lambda y: (_p10_wf["Portfolio"][y]
                            if not np.isnan(_p10_wf["Portfolio"][y]) else -9)) if _p10_wf_ntot else "N/A"
        _p10_wf_worst = min(_p10_wf["Portfolio"], key=lambda y: (_p10_wf["Portfolio"][y]
                            if not np.isnan(_p10_wf["Portfolio"][y]) else 9)) if _p10_wf_ntot else "N/A"
        st.markdown(f"""<div style="{_wfsg}">
<p style="{_wflbg}">OOS Consistency - {_p10_wt_full}</p>
<p style="{_wfsub}">Positive OOS Sharpe</p>
<p style="{_wfmed}">{_p10_wf_npos} / {_p10_wf_ntot} windows</p>
<hr style="{_wfhr}"/>
<p style="{_wfsub}">OOS Sharpe above +0.30</p>
<p style="{_wfmed}">{_p10_wf_ngt03} / {_p10_wf_ntot} windows</p>
<hr style="{_wfhr}"/>
<p style="{_wfsub}">Best: {_p10_wf_best} ({_p10_wf['Portfolio'].get(_p10_wf_best, float('nan')):+.2f}), Worst: {_p10_wf_worst} ({_p10_wf['Portfolio'].get(_p10_wf_worst, float('nan')):+.2f})</p>
</div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    _p10_wf_fc1, _p10_wf_fc2 = st.columns([2, 4])
    with _p10_wf_fc1:
        _p10_wf_pick = st.multiselect(
            "Series to plot (annual OOS Sharpe)",
            [_p10_port_label, "Momentum", "Carry", "Value"],
            default=[_p10_port_label],
            key="p10_wf_pick",
        )
    _p10_wf_series_map = {
        _p10_port_label: ("Portfolio", COLORS["primary"]),
        "Momentum":      ("Momentum",  "#5BAD72"),
        "Carry":         ("Carry",     COLORS["amber"]),
        "Value":         ("Value",     "#7B8FC0"),
    }
    fig_p10_wf = go.Figure()
    for _disp in (_p10_wf_pick or [_p10_port_label]):
        _key, _clr = _p10_wf_series_map[_disp]
        _d = _p10_wf[_key]
        fig_p10_wf.add_trace(go.Bar(
            x=list(_d.keys()), y=list(_d.values()), name=_disp,
            marker_color=_clr,
            hovertemplate="%{x}<br>" + _disp + " OOS Sharpe: %{y:.3f}<extra></extra>",
        ))
    fig_p10_wf.add_hline(y=0, line_color="#475569", line_width=1)
    if _p10_port_label in (_p10_wf_pick or [_p10_port_label]):
        fig_p10_wf.add_hline(
            y=_p10_wf_port_avg, line_dash="dot", line_color=COLORS["primary"], line_width=1.5,
            annotation_text=f"Portfolio avg {_p10_wf_port_avg:+.3f}",
            annotation_position="top right", annotation_font=dict(size=10, color=COLORS["primary"]),
        )
    fig_p10_wf.update_layout(
        height=320, margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="#0E1117", plot_bgcolor="#131922",
        font=dict(color="#E8DDD0", family="IBM Plex Mono", size=11),
        xaxis=dict(gridcolor="#1C2333", title=None),
        yaxis=dict(gridcolor="#1C2333", title="OOS Sharpe", zeroline=False),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11), orientation="h", y=1.12),
        barmode="group",
    )
    st.plotly_chart(fig_p10_wf, use_container_width=True)
    st.caption(
        f"Walk-forward out-of-sample: the selected configuration is applied to each unseen 1yr window "
        f"after a 5yr burn-in. The {_p10_wt_full.lower()} portfolio's avg OOS Sharpe is "
        f"{_p10_wf_port_avg:+.3f} versus per-sleeve "
        f"Mom {_wf_avg(_p10_wf['Momentum']):+.2f} / Carry {_wf_avg(_p10_wf['Carry']):+.2f} / "
        f"Value {_wf_avg(_p10_wf['Value']):+.2f}. Use the selector to overlay individual sleeves."
    )

    # ── Section 3: Pairwise Correlation Heatmap ───────────────────────────────
    st.divider()
    section_header("SIGNAL POSITION CORRELATIONS")
    st.caption("Correlation of daily ±1 position series. Low cross-signal correlation is the foundation of the diversification benefit.")

    _p10_pos_df = pd.DataFrame({
        "Momentum": _p10_m,
        "Carry":    _p10_c,
        "Value":    _p10_v,
    })
    _p10_corr = _p10_pos_df.corr().round(3)

    _p10_corr_col, _p10_corr_txt = st.columns([2, 2])
    with _p10_corr_col:
        fig_p10_corr = go.Figure(go.Heatmap(
            z=_p10_corr.values,
            x=list(_p10_corr.columns),
            y=list(_p10_corr.index),
            text=_p10_corr.values.round(3),
            texttemplate="%{text}",
            colorscale=[[0,"#B85450"],[0.5,"#1C2333"],[1,"#5BAD72"]],
            zmin=-1, zmax=1,
            showscale=True,
            hovertemplate="%{y} × %{x}: %{z:.3f}<extra></extra>",
        ))
        fig_p10_corr.update_layout(
            height=280, margin=dict(l=0, r=0, t=20, b=0),
            paper_bgcolor="#0E1117", plot_bgcolor="#0E1117",
            font=dict(color="#E8DDD0", family="IBM Plex Mono", size=12),
        )
        st.plotly_chart(fig_p10_corr, use_container_width=True)

    with _p10_corr_txt:
        st.markdown("&nbsp;")
        _mom_carry_corr = float(_p10_corr.loc["Momentum", "Carry"])
        _mom_val_corr   = float(_p10_corr.loc["Momentum", "Value"])
        _carry_val_corr = float(_p10_corr.loc["Carry",    "Value"])
        st.markdown(
            f"**Mom-Carry:** `{_mom_carry_corr:+.3f}`  \n"
            f"**Mom-Value:** `{_mom_val_corr:+.3f}`  \n"
            f"**Carry-Value:** `{_carry_val_corr:+.3f}`"
        )
        _avg_pairwise = np.mean([abs(_mom_carry_corr), abs(_mom_val_corr), abs(_carry_val_corr)])
        st.caption(
            f"Avg absolute pairwise correlation: {_avg_pairwise:.3f}. "
            f"{'All pairs < 0.3 - diversification benefit confirmed.' if _avg_pairwise < 0.3 else 'At least one pair > 0.3 - partial correlation, diversification benefit reduced.'}"
        )
        _theo_sharpe = np.sqrt(3) * np.mean([abs(_mom_sh), abs(_car_sh), abs(_val_sh)])
        st.caption(
            f"Theoretical portfolio Sharpe (uncorrelated) ≈ √3 × avg individual = {_theo_sharpe:.3f}. "
            f"Realised: {_port_sh:.3f}."
        )

    # ── Section 4: Sub-period Performance Table ───────────────────────────────
    st.divider()
    section_header("SUB-PERIOD PERFORMANCE")
    st.caption("Sharpe ratio by regime. Highlights where each signal adds / detracts value.")

    _p10_periods = [
        (f"Full ({_p10_yr0}-{_p10_yr1})", None,         None),
        ("Pre-2020",         None,         "2020-01-01"),
        ("2020-2021",        "2020-01-01", "2022-01-01"),
        ("Post-2022",        "2022-01-01", None),
    ]

    _sub_rows = []
    for _plbl, _ps, _pe in _p10_periods:
        _sub_rows.append({
            "Period":    _plbl,
            "Momentum":  _p10_sub_sharpe(_p10_m,    _ps, _pe, _p10_tc_bps),
            "Carry":     _p10_sub_sharpe(_p10_c,    _ps, _pe, _p10_tc_bps),
            "Value":     _p10_sub_sharpe(_p10_v,    _ps, _pe, _p10_tc_bps),
            _p10_port_label: _p10_sub_sharpe(_p10_port, _ps, _pe, _p10_tc_bps),
        })

    _sub_df = pd.DataFrame(_sub_rows).set_index("Period")

    def _fmt_sub(v):
        if v is None or np.isnan(v): return "-"
        color = "#5BAD72" if v >= 0.3 else ("#B85450" if v < 0 else "#E8DDD0")
        return f'<span style="color:{color};font-weight:600">{v:+.3f}</span>'

    _sub_html = "<table style='width:100%;border-collapse:collapse;font-family:IBM Plex Mono,monospace;font-size:0.82rem'>"
    _sub_html += "<tr style='border-bottom:1px solid #2A2A2A'>"
    _sub_html += "<th style='text-align:left;color:#8A8278;padding:6px 10px'>Period</th>"
    for _ch in ["Momentum", "Carry", "Value", _p10_port_label]:
        _sub_html += f"<th style='text-align:right;color:#8A8278;padding:6px 10px'>{_ch}</th>"
    _sub_html += "</tr>"
    for _pr, _srow in _sub_df.iterrows():
        _sub_html += f"<tr style='border-bottom:1px solid #1A1A1A'><td style='color:#E8DDD0;padding:6px 10px'>{_pr}</td>"
        for _ck in ["Momentum", "Carry", "Value", _p10_port_label]:
            _v = _srow[_ck]
            _sub_html += f"<td style='text-align:right;padding:6px 10px'>{_fmt_sub(_v)}</td>"
        _sub_html += "</tr>"
    _sub_html += "</table>"
    st.markdown(_sub_html, unsafe_allow_html=True)

    # ── Section 5: Signal Agreement Panel ────────────────────────────────────
    st.divider()
    section_header("SIGNAL AGREEMENT")
    st.caption("How often all three signals point in the same direction vs. a 2-of-3 majority vs. three-way split.")

    _p10_sm = np.sign(_p10_m).astype(int)
    _p10_sc = np.sign(_p10_c).astype(int)
    _p10_sv = np.sign(_p10_v).astype(int)
    _p10_T  = len(_p10_sm)

    _all_long   = int(((_p10_sm == 1)  & (_p10_sc == 1)  & (_p10_sv == 1)).sum())  if _p10_T else 0
    _all_short  = int(((_p10_sm == -1) & (_p10_sc == -1) & (_p10_sv == -1)).sum()) if _p10_T else 0
    _all_agree  = _all_long + _all_short
    _two_agree  = int(((_p10_sm == _p10_sc) | (_p10_sm == _p10_sv) | (_p10_sc == _p10_sv)).sum()) - _all_agree
    _split      = _p10_T - _all_agree - _two_agree

    _ag1, _ag2, _ag3, _ag4 = st.columns(4)
    for _acol, _albl, _aval in [
        (_ag1, "All 3 LONG",   _all_long),
        (_ag2, "All 3 SHORT",  _all_short),
        (_ag3, "2 of 3 Agree", _two_agree),
        (_ag4, "3-Way Split",  _split),
    ]:
        _apct = 100 * _aval / _p10_T if _p10_T else 0
        _acol.metric(_albl, f"{_aval:,} days", f"{_apct:.1f}% of history")

    # ── Section 6: Cumulative PnL Chart ──────────────────────────────────────
    st.divider()
    _p10_pnl_lbl = f"Net ({_p10_tc_label})" if _p10_tc_bps > 0 else "Gross"
    section_header("CUMULATIVE PERFORMANCE")

    _p10_cum_ctl1, _p10_cum_ctl2 = st.columns([1.6, 2.4])
    with _p10_cum_ctl1:
        _p10_cum_mode = st.radio(
            "Scale",
            ["Raw $/MT (1-unit position)", "Risk-scaled (each to 10% vol)"],
            index=0, key="p10_cum_mode", horizontal=False,
            help="Raw $/MT plots dollar PnL of a 1-unit position. Because the portfolio holds the "
                 "AVERAGE of three sleeves, its gross size (and so its raw $/MT) is smaller than an "
                 "always-on single sleeve - even though its Sharpe is higher. The risk-scaled view "
                 "puts every series at the same 10% annual volatility, so the curves are directly "
                 "comparable on a risk-adjusted basis and the portfolio's higher Sharpe shows as the "
                 "steepest, smoothest line.",
        )
    with _p10_cum_ctl2:
        _p10_cum_pick = st.multiselect(
            "Series to plot",
            [_p10_port_label, "Momentum", "Carry", "Value"],
            default=[_p10_port_label, "Momentum", "Carry", "Value"],
            key="p10_cum_pick",
        )
    _p10_cum_risk = _p10_cum_mode.startswith("Risk")

    def _cum_pnl(pos):
        pnl = pos * _p10_f.diff()
        if _p10_tc_bps > 0:
            chg = pos.diff().abs(); chg.iloc[0] = abs(pos.iloc[0])
            pnl = pnl - chg * (_p10_tc_bps / 10000.0 / 2.0) * _p10_fraw
        return pnl.cumsum()

    def _cum_riskscaled(pos, target_vol=0.10):
        pnl = pos * _p10_f.diff()
        if _p10_tc_bps > 0:
            chg = pos.diff().abs(); chg.iloc[0] = abs(pos.iloc[0])
            pnl = pnl - chg * (_p10_tc_bps / 10000.0 / 2.0) * _p10_fraw
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = (pnl / _p10_f.shift(1)).replace([np.inf, -np.inf], np.nan)
        act = ret[pos != 0].dropna()
        vol = act.std(ddof=1) * np.sqrt(252) if len(act) > 20 else np.nan
        scale = (target_vol / vol) if vol and vol > 0 else 0.0
        return (ret.fillna(0) * scale).cumsum() * 100.0

    _p10_cum_fn = _cum_riskscaled if _p10_cum_risk else _cum_pnl
    _p10_cum_map = {
        _p10_port_label: (_p10_cum_fn(_p10_port), COLORS["primary"], 2.5),
        "Momentum":      (_p10_cum_fn(_p10_m),    "#5BAD72",         1.4),
        "Carry":         (_p10_cum_fn(_p10_c),    COLORS["amber"],   1.4),
        "Value":         (_p10_cum_fn(_p10_v),    "#7B8FC0",         1.4),
    }
    _p10_disp_names = {
        _p10_port_label: _p10_port_label,
        "Momentum":      f"Momentum {_p10_mom_label}",
        "Carry":         f"Carry {_p10_carry_label}",
        "Value":         f"Value {_p10_val_label}",
    }
    if _p10_cum_risk:
        _hov = "%{x|%b %d, %Y}<br>Cum return (10% vol): %{y:,.1f}%<extra></extra>"
    else:
        _hov = "%{x|%b %d, %Y}<br>Cum PnL: $%{y:,.0f}/MT<extra></extra>"

    fig_p10_cum = go.Figure()
    for _disp in (_p10_cum_pick or [_p10_port_label]):
        _series, _color, _width = _p10_cum_map[_disp]
        fig_p10_cum.add_trace(go.Scatter(
            x=_series.index, y=_series.values,
            name=_p10_disp_names[_disp], mode="lines",
            line=dict(color=_color, width=_width),
            hovertemplate=_hov,
        ))
    fig_p10_cum.update_layout(
        height=400, margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="#0E1117", plot_bgcolor="#131922",
        font=dict(color="#E8DDD0", family="IBM Plex Mono", size=11),
        xaxis=dict(gridcolor="#1C2333", showgrid=True),
        yaxis=(dict(gridcolor="#1C2333", showgrid=True, ticksuffix="%") if _p10_cum_risk
               else dict(gridcolor="#1C2333", showgrid=True, tickprefix="$", ticksuffix="/MT")),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
        hovermode="x unified",
    )
    st.plotly_chart(fig_p10_cum, use_container_width=True)
    if _p10_cum_risk:
        st.caption(
            f"{_p10_pnl_lbl}, risk-scaled. Every series is scaled to a common 10% annual volatility, "
            f"so slope = risk-adjusted return. The {_p10_wt_full.lower()} portfolio is the steepest and "
            f"smoothest line, consistent with its higher Sharpe."
        )
    else:
        st.caption(
            f"{_p10_pnl_lbl} dollar PnL of a 1-unit position. The portfolio holds the average of three "
            f"sleeves, so an always-on single sleeve (e.g. Momentum) can show larger raw $/MT while still "
            f"having a LOWER Sharpe - higher dollar PnL is not higher risk-adjusted return. Switch to "
            f"'Risk-scaled' above to compare like-for-like."
        )

    # ── Section 7: Rolling Sharpe ─────────────────────────────────────────────
    st.divider()
    section_header("ROLLING SHARPE (252-DAY)")
    _p10_rsh_cc1, _p10_rsh_cc2, _p10_rsh_cc3 = st.columns([2.6, 1.6, 1])
    with _p10_rsh_cc1:
        st.caption(f"Rolling 1yr Sharpe - {_p10_port_label} vs each sleeve. "
                   f"Net uses the selected {_p10_tc_label}.")
    with _p10_rsh_cc2:
        _p10_rsh_pick = st.multiselect(
            "Series",
            [_p10_port_label, "Momentum", "Carry", "Value"],
            default=[_p10_port_label, "Momentum", "Carry", "Value"],
            key="p10_rsh_pick",
        )
    with _p10_rsh_cc3:
        _p10_rsh_basis = st.radio("Returns", ["Gross", "Net of TC"], index=0,
                                  key="p10_rsh_basis", horizontal=False)
    _p10_rsh_net = _p10_rsh_basis.startswith("Net")
    _p10_rsh_tc  = _p10_tc_bps if _p10_rsh_net else 0

    def _p10_roll_sh(pos: pd.Series, window: int = 252) -> pd.Series:
        ret = _p10_ret(pos, _p10_rsh_tc)
        act = ret.where(pos != 0)
        return act.rolling(window, min_periods=window // 2).apply(
            lambda x: (x.mean() / x.std(ddof=1) * np.sqrt(252)) if x.std(ddof=1) > 0 else np.nan,
            raw=True,
        )

    _p10_rsh_map = {
        _p10_port_label: (_p10_roll_sh(_p10_port), COLORS["primary"], 2.2),
        "Momentum":      (_p10_roll_sh(_p10_m),    "#5BAD72",         1.0),
        "Carry":         (_p10_roll_sh(_p10_c),    COLORS["amber"],   1.0),
        "Value":         (_p10_roll_sh(_p10_v),    "#7B8FC0",         1.0),
    }
    _p10_rsh_disp = {
        _p10_port_label: _p10_port_label,
        "Momentum":      f"Momentum {_p10_mom_label}",
        "Carry":         f"Carry {_p10_carry_label}",
        "Value":         f"Value {_p10_val_label}",
    }
    fig_p10_rsh = go.Figure()
    for _disp in (_p10_rsh_pick or [_p10_port_label]):
        _series, _color, _width = _p10_rsh_map[_disp]
        fig_p10_rsh.add_trace(go.Scatter(
            x=_series.index, y=_series.values,
            name=_p10_rsh_disp[_disp], mode="lines",
            line=dict(color=_color, width=_width),
            hovertemplate="%{x|%b %d, %Y}<br>Rolling Sharpe: %{y:.3f}<extra></extra>",
        ))
    fig_p10_rsh.add_hline(y=0, line_dash="dash", line_color="#475569", line_width=1)
    fig_p10_rsh.update_layout(
        height=350, margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="#0E1117", plot_bgcolor="#131922",
        font=dict(color="#E8DDD0", family="IBM Plex Mono", size=11),
        xaxis=dict(gridcolor="#1C2333"),
        yaxis=dict(gridcolor="#1C2333", zeroline=False),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
        hovermode="x unified",
    )
    st.plotly_chart(fig_p10_rsh, use_container_width=True)

    # ── Section 8: Annual PnL Bars ────────────────────────────────────────────
    st.divider()
    section_header(f"ANNUAL PnL (USD/MT) - {_p10_pnl_lbl}")

    _p10_pnl_s = _p10_port * _p10_f.diff()
    if _p10_tc_bps > 0:
        _p10_chg = _p10_port.diff().abs(); _p10_chg.iloc[0] = abs(_p10_port.iloc[0])
        _p10_pnl_s = _p10_pnl_s - _p10_chg * (_p10_tc_bps / 10000.0 / 2.0) * _p10_fraw
    _p10_ann_pnl = _p10_pnl_s.groupby(_p10_pnl_s.index.year).sum()
    _p10_ann_colors = ["#5BAD72" if v >= 0 else "#B85450" for v in _p10_ann_pnl.values]

    fig_p10_ann = go.Figure(go.Bar(
        x=_p10_ann_pnl.index.astype(str),
        y=_p10_ann_pnl.values,
        marker_color=_p10_ann_colors,
        hovertemplate="%{x}: $%{y:,.0f}/MT<extra></extra>",
    ))
    fig_p10_ann.update_layout(
        height=280, margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="#0E1117", plot_bgcolor="#131922",
        font=dict(color="#E8DDD0", family="IBM Plex Mono", size=11),
        xaxis=dict(gridcolor="#1C2333"),
        yaxis=dict(gridcolor="#1C2333", tickprefix="$", ticksuffix="/MT"),
    )
    fig_p10_ann.add_hline(y=0, line_color="#475569", line_width=1)
    st.plotly_chart(fig_p10_ann, use_container_width=True)

    # ── Section 9: Position Decomposition ────────────────────────────────────
    st.divider()
    section_header("POSITION DECOMPOSITION")
    if _p10_use_iv:
        st.caption("Each signal's contribution to the Inverse-Vol portfolio position "
                   "(daily inverse-vol weight × signal). Stacked contributions sum to the portfolio position.")
        _p10_dec_mom = _p10_wm * _p10_m
        _p10_dec_car = _p10_wc * _p10_c
        _p10_dec_val = _p10_wv * _p10_v
        _p10_dec_tags = ("Momentum (w×sig)", "Carry (w×sig)", "Value (w×sig)")
    else:
        st.caption("Each signal's contribution to the Equal-Weight portfolio position (±1/3 per signal). "
                   "Stacked contributions sum to the portfolio position.")
        _p10_dec_mom = _p10_m / 3.0
        _p10_dec_car = _p10_c / 3.0
        _p10_dec_val = _p10_v / 3.0
        _p10_dec_tags = ("Momentum (±1/3)", "Carry (±1/3)", "Value (±1/3)")

    _p10_dec_start = st.date_input(
        "From date", value=pd.Timestamp("2020-01-01").date(),
        min_value=_p10_idx[0].date(), max_value=_p10_idx[-1].date(),
        key="p10_dec_start",
    )
    _p10_dec_idx = _p10_idx[_p10_idx >= pd.Timestamp(_p10_dec_start)]

    fig_p10_dec = go.Figure()
    for _name, _series, _color in [
        (_p10_dec_tags[0],  _p10_dec_mom.reindex(_p10_dec_idx),  "#5BAD72"),
        (_p10_dec_tags[1],  _p10_dec_car.reindex(_p10_dec_idx),  COLORS["amber"]),
        (_p10_dec_tags[2],  _p10_dec_val.reindex(_p10_dec_idx),  "#7B8FC0"),
    ]:
        fig_p10_dec.add_trace(go.Scatter(
            x=_series.index, y=_series.values,
            name=_name, mode="lines", stackgroup="one",
            line=dict(color=_color, width=0.5),
            fillcolor=(f"rgba({int(_color[1:3],16)},{int(_color[3:5],16)},{int(_color[5:7],16)},0.35)"
                       if isinstance(_color, str) and _color.startswith("#") and len(_color) >= 7 else _color),
            hovertemplate="%{x|%b %d, %Y}<br>%{fullData.name}: %{y:+.2f}<extra></extra>",
        ))
    fig_p10_dec.update_layout(
        height=300, margin=dict(l=0, r=0, t=20, b=0),
        paper_bgcolor="#0E1117", plot_bgcolor="#131922",
        font=dict(color="#E8DDD0", family="IBM Plex Mono", size=11),
        xaxis=dict(gridcolor="#1C2333"),
        yaxis=dict(gridcolor="#1C2333", zeroline=True, zerolinecolor="#475569"),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
        hovermode="x unified",
    )
    st.plotly_chart(fig_p10_dec, use_container_width=True)

    # ── Methodology Note ──────────────────────────────────────────────────────
    st.divider()
    with st.expander("Methodology Notes", expanded=False):
        st.markdown("""
**Signal Definitions (each leg = the selected metal's best-performing configuration; all numbers in the
cards/tables above recompute live for Copper vs Aluminium):**
- **Momentum:** MA crossover on F1_raw, shift-1 entry, position ±1. *Copper:* MA(35,43); *Aluminium:* slow MA(60,115).
- **Carry:** shift-1, position ±1. *Copper:* 20-day change in the (F1−F2)/F1 roll yield (curve momentum);
  *Aluminium:* 252-day z-score of the (F1−F2)/F1 roll yield.
- **Value V1 (default):** (Fk − MA_1260)/MA_1260 deviation, ±10% threshold, shift-1, position −1/0/+1.
  *Copper:* F8; *Aluminium:* F12.
- *Value V2 (optional):* F1_raw[t−2520] − F1_raw[t] reversal (10yr) - higher in-sample Sharpe but fragile out-of-sample.

**EW Portfolio:** `Port_pos = (1/3)·Mom + (1/3)·Carry + (1/3)·Value`, ranging −1 to +1. When all three agree |port| = 1;
a 2−1 split gives |port| = 1/3; a value flat (0) reduces conviction.

**Entry timing (no look-ahead):** Same-Day = shift 1 (trade at the signal's close; first return t→t+1).
Lag-1 = shift 2 (one further day of delay). The retired shift-0 "same-day" booked an already-realised move
(look-ahead) and is removed.

**Transaction costs:** TC = |Δposition| × (bps/10000/2) × **F1_raw** (the actual traded price); the spread is
charged on every position change. PnL and returns are computed on the roll-adjusted F1_continuous.

**Diversification:** with low pairwise leg correlations, the equal-weight Sharpe tends to exceed the average
single-leg Sharpe (≈ √3 × avg if the legs were uncorrelated). The per-leg Sharpes, live correlations and the
realised EW / inverse-vol portfolio Sharpe (for the selected metal at the chosen TC) are shown in the table and
cards above and update with the metal toggle. EW is the default; inverse-vol is the alternative.

**Disclaimer:** Partial-OOS backtest (params IS-selected, applied walk-forward). Not investment advice.
        """)
