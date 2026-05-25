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
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');

    /* Global */
    .stApp { font-family: 'DM Sans', sans-serif; }

    /* Hide default streamlit elements */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

    /* Metric cards */
    .metric-card {
        background: linear-gradient(135deg, #1A1F2E 0%, #0F1724 100%);
        border: 1px solid #2D3748;
        border-radius: 12px;
        padding: 16px 20px;
        margin: 4px 0;
    }
    .metric-card h4 {
        color: #94A3B8;
        font-size: 0.75rem;
        font-weight: 500;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        margin: 0 0 4px 0;
    }
    .metric-card .value {
        color: #E2E8F0;
        font-family: 'JetBrains Mono', monospace;
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0;
    }
    .metric-card .delta-pos { color: #34D399; font-size: 0.85rem; }
    .metric-card .delta-neg { color: #F87171; font-size: 0.85rem; }

    /* Section headers */
    .section-header {
        font-family: 'DM Sans', sans-serif;
        color: #E2E8F0;
        font-size: 1.1rem;
        font-weight: 700;
        border-bottom: 2px solid #3B82F6;
        padding-bottom: 8px;
        margin: 24px 0 16px 0;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background-color: #0F1724;
        padding: 4px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 20px;
        font-weight: 500;
    }

    /* Backwardation / Contango badges */
    .badge-backwardation {
        background: rgba(52, 211, 153, 0.15);
        color: #34D399;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    .badge-contango {
        background: rgba(248, 113, 113, 0.15);
        color: #F87171;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
    }

    /* Title */
    .main-title {
        font-family: 'DM Sans', sans-serif;
        font-weight: 700;
        font-size: 1.8rem;
        color: #E2E8F0;
        margin-bottom: 0;
    }
    .main-subtitle {
        color: #64748B;
        font-size: 0.9rem;
        margin-top: 0;
    }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# CHART THEME
# ═══════════════════════════════════════════════

CHART_LAYOUT = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(14,17,23,0.8)",
    font=dict(family="DM Sans, sans-serif", color="#94A3B8"),
    xaxis=dict(gridcolor="rgba(45,55,72,0.5)", zerolinecolor="rgba(45,55,72,0.5)"),
    yaxis=dict(gridcolor="rgba(45,55,72,0.5)", zerolinecolor="rgba(45,55,72,0.5)"),
    margin=dict(l=60, r=30, t=50, b=50),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    hoverlabel=dict(bgcolor="#1A1F2E", font_size=12, font_family="JetBrains Mono"),
)

COLORS = {
    "primary": "#3B82F6",
    "secondary": "#8B5CF6",
    "accent": "#06B6D4",
    "green": "#34D399",
    "red": "#F87171",
    "amber": "#FBBF24",
    "orange": "#FB923C",
    "pink": "#F472B6",
    "slate": "#64748B",
}

METAL_COLORS = {
    "Copper": "#FB923C",
    "Aluminium": "#94A3B8",
    "Zinc": "#3B82F6",
    "Nickel": "#8B5CF6",
    "Lead": "#64748B",
    "Tin": "#06B6D4",
    "Gold": "#FBBF24",
    "Silver": "#CBD5E1",
    "Platinum": "#A78BFA",
    "Palladium": "#F472B6",
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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Market Overview",
    "📈 Term Structure",
    "💰 Cash vs 3M (Carry)",
    "📉 Volume & Open Interest",
    "🔗 Copper LME-CME Spread",
    "📋 Statistics"
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
                            date_options = [d.strftime("%Y-%m-%d") for d in reversed(available_dates)]
                            selected_date_str = st.selectbox(
                                "Select Date", date_options, index=0, key="curve_date_select"
                            )
                            selected_curve_date = pd.Timestamp(selected_date_str)

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

