"""
Load NGL_PAPER1.xlsx.

Sheet structure (rows are 1-indexed in Excel):
  Row 1: title header (skip)
  Row 2: contract labels  – Date | F1 | None | None | F2 | None | None | …
  Row 3: sub-labels       – None | Price | Volume | Open | Price | Volume | Open | …
  Row 4+: data

Each contract occupies 3 columns: [Price, Volume, Open].
F1 starts at column-index 1, F2 at 4, Fn at 1 + (n-1)*3.

Bloomberg stores:
  CAP, BAP, DAE  → $/gal   (e.g. 0.50 = 50¢/gal)
  PCW, PGP       → $/lb    (e.g. 0.40 = 40¢/lb)
"""

import pandas as pd
from pathlib import Path


def _load_sheet(filepath: Path, sheet: str, max_tenors: int = 2) -> pd.DataFrame:
    """Return a DataFrame with Date + F1_Price, F2_Price (and F1_Volume)."""
    wb_raw = pd.read_excel(filepath, sheet_name=sheet, header=None)

    # Row index 1 (0-based) is contract row: Date, F1, None, None, F2, ...
    # Row index 2 is sub-label row: None, Price, Volume, Open, Price, ...
    # Data starts at row index 3.

    col_map = {"Date": 0}
    for n in range(1, max_tenors + 1):
        base = 1 + (n - 1) * 3
        col_map[f"F{n}_Price"]  = base
        col_map[f"F{n}_Volume"] = base + 1
        col_map[f"F{n}_Open"]   = base + 2

    df = wb_raw.iloc[3:].copy()   # data rows only
    df = df.rename(columns={v: k for k, v in col_map.items()})
    df = df[list(col_map.keys())].copy()

    # Parse dates
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df = df.set_index("Date").sort_index()

    # Coerce prices to numeric
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def load_ngl_data(filepath: Path) -> dict[str, pd.DataFrame]:
    """
    Returns dict keyed by ticker (CAP, BAP, DAE, PCW, PGP).
    Each value is a DataFrame indexed by Date with columns:
      F1_Price, F2_Price, F1_Volume, F1_Open, F2_Volume, F2_Open
    """
    tickers = ["CAP", "BAP", "DAE", "PCW", "PGP"]
    data = {}
    for t in tickers:
        data[t] = _load_sheet(filepath, sheet=t, max_tenors=2)
    return data
