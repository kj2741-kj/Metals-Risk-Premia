from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR.parent  # NGL and VRP folder

NGL_FILE = DATA_DIR / "NGL_PAPER1.xlsx"
OUTPUT_DIR = BASE_DIR / "output"
CHARTS_DIR = OUTPUT_DIR / "charts"

# Sample period from paper
START_DATE = "2011-07-07"
END_DATE   = "2025-02-24"

# Regime shading boundaries (inclusive)
REGIMES = {
    "Pre-Shale":  ("2011-07-07", "2014-12-31"),
    "Shale Boom": ("2015-01-01", "2019-12-31"),
    "COVID-19":   ("2020-01-01", "2022-12-31"),
}
REGIME_COLORS = {
    "Pre-Shale":  "#FFF3CD",
    "Shale Boom": "#D4EDDA",
    "COVID-19":   "#CCE5FF",
}

# Petrochemical densities (lb/gal) from Engineering Toolbox
ETHYLENE_DENSITY  = 4.91   # lb/gal
PROPYLENE_DENSITY = 4.16   # lb/gal

# Transaction cost per leg per roll
ROLL_COST = 0.002  # 0.2%

# Optimization grids
LOOKBACK_GRID         = [30, 60, 90, 120, 150, 200, 250]   # days
INTRA_THRESHOLD_GRID  = [0.00, 0.15, 0.30, 0.45, 0.60, 0.75, 0.90]  # $/gal
INTER_THRESHOLD_GRID  = [0.00, 0.02, 0.04, 0.08, 0.16, 0.32, 0.64]  # $/gal

TRADING_DAYS_PER_YEAR = 252
