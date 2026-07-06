"""
Time-series momentum signal (Paper 2, Section 3.1).

Three MA crossovers are computed for each instrument:
  (1, 5), (5, 20), (10, 60)

Signal for each crossover:
  +1 if short MA > long MA   (uptrend)
  -1 if short MA < long MA   (downtrend)
   0 if equal (rare)

Final momentum signal = sign(sum of the three sub-signals).
If sum = 0, signal = 0 (no position).

Signal is lagged by 1 day (computed at close, applied next open/close).
"""

import numpy as np
import pandas as pd
from config import MOMENTUM_PAIRS, INSTRUMENTS


def _ma_cross(price: pd.Series, short: int, long: int) -> pd.Series:
    """Return +1/-1 signal from a single MA crossover pair."""
    s_ma = price.rolling(short, min_periods=short).mean()
    l_ma = price.rolling(long,  min_periods=long).mean()
    raw  = np.sign(s_ma - l_ma)
    return raw.fillna(0.0)


def momentum_signal(price: pd.Series) -> pd.Series:
    """
    Compute the combined momentum signal for one price series.
    Returns daily signal series, lagged 1 day.
    """
    subs = [_ma_cross(price, s, l) for s, l in MOMENTUM_PAIRS]
    combined = sum(subs)          # sum of three ±1 signals
    sig = np.sign(combined).fillna(0.0)
    return sig.shift(1).fillna(0.0).rename(price.name)


def all_momentum_signals(
    cont: dict[str, tuple[pd.Series, pd.Series]],
) -> pd.DataFrame:
    """
    Compute momentum signals for all instruments.
    Returns DataFrame: rows=Date, cols=ticker.
    """
    signals = {}
    for ticker, (adj, _) in cont.items():
        signals[ticker] = momentum_signal(adj.dropna())
    return pd.DataFrame(signals)
