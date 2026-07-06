"""
Carry (term-structure) signal (Paper 2, Section 3.2).

Carry = slope of the futures curve, measured as:
  carry_t = F4_t - F15_t

Signal:
  +1 if carry_t > 0   (backwardation → long)
  -1 if carry_t < 0   (contango → short)
   0 if carry_t = 0

Threshold = 0 (any non-zero slope triggers a position).
Signal lagged 1 day.
"""

import numpy as np
import pandas as pd
from config import CARRY_SHORT_TENOR, CARRY_LONG_TENOR, INSTRUMENTS
from data_loader import get_price_strip


def carry_signal(
    raw: dict,
    ticker: str,
) -> pd.Series:
    """
    Return carry signal (+1/-1/0) for one ticker, lagged 1 day.
    raw: output of load_all_data()
    """
    strip  = get_price_strip(raw, ticker)
    f_near = strip.get(f"F{CARRY_SHORT_TENOR}")
    f_far  = strip.get(f"F{CARRY_LONG_TENOR}")

    if f_near is None or f_far is None:
        return pd.Series(dtype=float, name=ticker)

    carry = f_near - f_far
    sig   = np.sign(carry).fillna(0.0)
    return sig.shift(1).fillna(0.0).rename(ticker)


def all_carry_signals(raw: dict) -> pd.DataFrame:
    """Carry signals for all instruments. Returns Date × ticker DataFrame."""
    signals = {}
    for ticker in INSTRUMENTS:
        signals[ticker] = carry_signal(raw, ticker)
    return pd.DataFrame(signals)
