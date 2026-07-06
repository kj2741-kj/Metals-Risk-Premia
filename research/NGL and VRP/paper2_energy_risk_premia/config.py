"""
Paper 2 – Risk Premia in Diversified Energy Portfolios: configuration.
"""

from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR.parent   # NGL and VRP folder

# ── Data files ────────────────────────────────────────────────────────────────
NGL_FILE    = DATA_DIR / "NGL_PAPER1.xlsx"
NG_QS_FILE  = DATA_DIR / "NG_QS_Futures.xlsx"
CRACK_FILE  = DATA_DIR / "Crack_futures_data.xlsx"
EXPIRY_FILE = DATA_DIR / "expiry_calendars_20260526.xlsx"

# ── Output ────────────────────────────────────────────────────────────────────
OUTPUT_DIR = BASE_DIR / "output"
CHARTS_DIR = OUTPUT_DIR / "charts"

# ── Sample periods ────────────────────────────────────────────────────────────
# Data from 2010 used to initialize long-horizon signals (10yr value MA etc.)
DATA_START   = "2010-01-04"
# Formal backtest window: Jan 2, 2015 – Dec 1, 2025  (paper Section 4)
FULL_START   = "2015-01-02"
FULL_END     = "2025-12-01"

# Sub-sample break at Jan 2, 2022 (pandemic recovery / Russia-Ukraine regime shift)
SUBSAMPLE_1  = ("2015-01-02", "2021-12-31")   # pre-2022
SUBSAMPLE_2  = ("2022-01-02", "2025-12-01")   # post-2022

# ── Instrument universe (Paper 2 core-6 + NGLs) ──────────────────────────────
# Bloomberg tickers used as keys throughout the project
INSTRUMENTS = {
    # Crack / crude futures  (Crack_futures_data.xlsx)
    "CL":  {"name": "WTI Crude",        "file": "crack", "sheet": "CL",  "unit": "usd_bbl"},
    "CO":  {"name": "Brent Crude",      "file": "crack", "sheet": "CO",  "unit": "usd_bbl"},
    "XB":  {"name": "RBOB Gasoline",    "file": "crack", "sheet": "XB",  "unit": "cps_gal"},
    "HO":  {"name": "ULSD Heating Oil", "file": "crack", "sheet": "HO",  "unit": "cps_gal"},
    # Natural gas / gasoil   (NG_QS_Futures.xlsx)
    "NG":  {"name": "Henry Hub Gas",    "file": "ng_qs", "sheet": "NG",  "unit": "usd_mmbtu"},
    "QS":  {"name": "ICE Gasoil",       "file": "ng_qs", "sheet": "QS",  "unit": "usd_mt"},
    # NGLs                   (NGL_PAPER1.xlsx)
    "CAP": {"name": "Ethane",           "file": "ngl",   "sheet": "CAP", "unit": "usd_gal"},
    "BAP": {"name": "Propane",          "file": "ngl",   "sheet": "BAP", "unit": "usd_gal"},
    "DAE": {"name": "Normal Butane",    "file": "ngl",   "sheet": "DAE", "unit": "usd_gal"},
}

# Instruments that have actual expiry dates in the calendar
EXPIRY_TICKERS = ["CL", "CO", "XB", "HO", "NG", "QS"]
# NGL swaps use end-of-month rule (no exact expiry dates in calendar)
NGL_TICKERS = ["CAP", "BAP", "DAE"]
# Always skip F1 for NGLs (swap averaging); use F2 as front contract
NGL_SKIP_FRONT = True

# ── Unit conversion to $/bbl (Appendix E of the paper) ───────────────────────
# Applied after roll-adjusted series are built
# cps_gal = US cents per gallon; 42 gal/bbl → multiply by 42/100 = 0.42
# usd_mt  = USD per metric ton; 1 mt gasoil ≈ 7.45 bbl → multiply by 1/7.45
BBL_CONVERSIONS = {
    "usd_bbl":    1.0,
    "cps_gal":    0.42,          # ¢/gal → $/bbl
    "usd_gal":   42.0,           # $/gal → $/bbl
    "usd_mmbtu":  5.8,           # $/MMBtu → $/bbl  (crude energy equiv.)
    "usd_mt":    (1 / 7.45),     # $/mt  → $/bbl  (gasoil density)
}

# ── Strategy parameters ───────────────────────────────────────────────────────
TRADING_DAYS_PER_YEAR = 252
ROLL_COST_BPS = 0.002           # 0.20% per leg (decimal fraction, not basis points)

# Momentum – three MA crossovers (in trading days); signal = sign(average of three)
# (1,5), (5,20), (10,60): short / medium / longer daily-lookback trend signals
MOMENTUM_PAIRS = [(1, 5), (5, 20), (10, 60)]

# Carry – use F4 vs F15; threshold = 0 (any backwardation/contango)
CARRY_SHORT_TENOR  = 4
CARRY_LONG_TENOR   = 15

# Value – 10-year (2520 TD) running MA on F12/F13; threshold = 10%
VALUE_LOOKBACK     = 2520        # trading days ≈ 10 years
VALUE_CONTRACT     = 12          # use F12 as the "1-year-out" price
VALUE_THRESHOLD    = 0.10        # 10% band

# Stat-arb spreads (Paper 2, Table 5) – (leg1_ticker, leg2_ticker)
STAT_ARB_SPREADS = {
    "CL-CO":   ("CL",  "CO"),
    "XB-HO":   ("XB",  "HO"),
    "XB-CL":   ("XB",  "CL"),
    "HO-CL":   ("HO",  "CL"),
    "NG-CL":   ("NG",  "CL"),
    "CAP-BAP": ("CAP", "BAP"),
    "BAP-DAE": ("BAP", "DAE"),
    "CAP-DAE": ("CAP", "DAE"),
}
STAT_ARB_LOOKBACK   = 20         # days
STAT_ARB_THRESHOLD  = 0.10       # 10% of rolling std
STAT_ARB_CONTRACT   = 12         # use year-out futures (F12/F13) for stat-arb

# ── Risk model ────────────────────────────────────────────────────────────────
EWMA_VOL_HALFLIFE   = 20         # days  (vol normalisation)
EWMA_COV_HALFLIFE   = 60         # days  (covariance matrix)
LW_SHRINKAGE_ALPHA  = 0.5        # Ledoit-Wolf blend: α×LW + (1-α)×sample
