"""
Value signal (Paper 2, Section 3.3).

Uses the 1-year-out contract (F12/F13) to avoid front-month distortions.

Value signal:
  ma_t = rolling(VALUE_LOOKBACK)-day mean of F12 price
  deviation = (F12_t - ma_t) / ma_t

  +1 if deviation < -VALUE_THRESHOLD  (price well below long-run mean → buy)
  -1 if deviation >  VALUE_THRESHOLD  (price well above long-run mean → sell)
   0 otherwise

Signal lagged 1 day.
"""

import numpy as np
import pandas as pd
from config import VALUE_LOOKBACK, VALUE_CONTRACT, VALUE_THRESHOLD, INSTRUMENTS, FULL_START
from data_loader import get_price_strip


def value_signal(raw: dict, ticker: str) -> pd.Series:
    """Return value signal (+1/-1/0) for one ticker, lagged 1 day."""
    strip = get_price_strip(raw, ticker)
    col   = f"F{VALUE_CONTRACT}"
    if col not in strip.columns:
        return pd.Series(dtype=float, name=ticker)

    price = strip[col].dropna()
    ma    = price.rolling(VALUE_LOOKBACK, min_periods=VALUE_LOOKBACK // 2).mean()
    dev   = (price - ma) / ma.replace(0, np.nan)

    sig = np.where(dev < -VALUE_THRESHOLD, 1,
           np.where(dev >  VALUE_THRESHOLD, -1, 0))
    s = pd.Series(sig, index=price.index, dtype=float).fillna(0.0)
    return s.shift(1).fillna(0.0).rename(ticker)


def all_value_signals(raw: dict) -> pd.DataFrame:
    """Value signals for all instruments. Returns Date × ticker DataFrame."""
    signals = {}
    for ticker in INSTRUMENTS:
        signals[ticker] = value_signal(raw, ticker)
    return pd.DataFrame(signals)
