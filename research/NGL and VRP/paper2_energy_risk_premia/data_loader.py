"""
Load all three Excel files for Paper 2.

Sheet structures:
  NGL_PAPER1.xlsx   : Row 1 skip, Row 2 contract labels, Row 3 sub-labels, Row 4+ data
                      3 cols per contract [Price, Volume, Open]; NGL tickers
  NG_QS_Futures.xlsx: same 3-col structure; NG (Henry Hub) and QS (ICE Gasoil)
  Crack_futures_data.xlsx: 5 cols per contract [Price, Volume, Open, High, Low]

Returns a nested dict:
  {ticker: {F1: Series, F2: Series, ..., F27: Series}}
  Each Series is indexed by Date, values are prices in raw Bloomberg units.
"""

import pandas as pd
from pathlib import Path
from config import (
    NGL_FILE, NG_QS_FILE, CRACK_FILE,
    INSTRUMENTS, NGL_TICKERS,
)


MAX_TENORS = 27


def _load_3col_sheet(filepath: Path, sheet: str, max_tenors: int = MAX_TENORS) -> dict[str, pd.Series]:
    """
    Parse a sheet with 3 columns per contract (Price, Volume, Open).
    Returns {F1_Price: Series, F2_Price: Series, …}.
    """
    wb = pd.read_excel(filepath, sheet_name=sheet, header=None)

    col_map = {0: "Date"}
    for n in range(1, max_tenors + 1):
        base = 1 + (n - 1) * 3
        col_map[base]     = f"F{n}_Price"
        col_map[base + 1] = f"F{n}_Volume"
        col_map[base + 2] = f"F{n}_Open"

    data = wb.iloc[3:].copy()
    # Rename only columns that exist in this sheet
    existing = {k: v for k, v in col_map.items() if k < data.shape[1]}
    data = data.rename(columns=existing)
    named_cols = [v for k, v in sorted(existing.items())]
    data = data[named_cols].copy()

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    data = data.dropna(subset=["Date"]).set_index("Date").sort_index()
    for col in data.columns:
        data[col] = pd.to_numeric(data[col], errors="coerce")

    return {col: data[col] for col in data.columns}


def _load_5col_sheet(filepath: Path, sheet: str, max_tenors: int = MAX_TENORS) -> dict[str, pd.Series]:
    """
    Parse a sheet with 5 columns per contract (Price, Volume, Open, High, Low).
    Returns {F1_Price: Series, F2_Price: Series, …}.
    """
    wb = pd.read_excel(filepath, sheet_name=sheet, header=None)

    col_map = {0: "Date"}
    for n in range(1, max_tenors + 1):
        base = 1 + (n - 1) * 5
        col_map[base]     = f"F{n}_Price"
        col_map[base + 1] = f"F{n}_Volume"
        col_map[base + 2] = f"F{n}_Open"
        col_map[base + 3] = f"F{n}_High"
        col_map[base + 4] = f"F{n}_Low"

    data = wb.iloc[3:].copy()
    existing = {k: v for k, v in col_map.items() if k < data.shape[1]}
    data = data.rename(columns=existing)
    named_cols = [v for k, v in sorted(existing.items())]
    data = data[named_cols].copy()

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    data = data.dropna(subset=["Date"]).set_index("Date").sort_index()
    for col in data.columns:
        data[col] = pd.to_numeric(data[col], errors="coerce")

    return {col: data[col] for col in data.columns}


def load_all_data() -> dict[str, dict[str, pd.Series]]:
    """
    Load all instruments.
    Returns {ticker: {F1_Price: Series, F2_Price: Series, …}}.
    """
    result: dict[str, dict[str, pd.Series]] = {}

    # NGL tickers (3-col)
    for ticker in NGL_TICKERS:
        series_map = _load_3col_sheet(NGL_FILE, sheet=ticker)
        result[ticker] = series_map

    # Natural gas & gasoil (3-col)
    for ticker in ["NG", "QS"]:
        series_map = _load_3col_sheet(NG_QS_FILE, sheet=ticker)
        result[ticker] = series_map

    # Crack / crude futures (5-col)
    for ticker in ["CL", "CO", "XB", "HO"]:
        series_map = _load_5col_sheet(CRACK_FILE, sheet=ticker)
        result[ticker] = series_map

    return result


def get_price_strip(raw: dict[str, dict[str, pd.Series]], ticker: str) -> pd.DataFrame:
    """
    Return a DataFrame with columns F1, F2, … for the given ticker,
    index = Date, values = raw Bloomberg prices.
    """
    series_map = raw[ticker]
    price_cols = {k: v for k, v in series_map.items() if k.endswith("_Price")}
    df = pd.DataFrame(price_cols)
    df.columns = [c.replace("_Price", "") for c in df.columns]
    df = df.sort_index()
    return df
