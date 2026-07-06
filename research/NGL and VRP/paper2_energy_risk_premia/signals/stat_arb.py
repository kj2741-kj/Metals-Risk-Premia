"""
Statistical arbitrage signal (Paper 2, Section 3.4).

For each defined spread pair (leg1, leg2):
  spread_t = F12_leg1_t - F12_leg2_t      (year-out contracts, same units after conversion)

  rolling_std = std(spread, window=STAT_ARB_LOOKBACK)
  ma          = mean(spread, window=STAT_ARB_LOOKBACK)
  z           = (spread_t - ma_t) / rolling_std_t

  signal:
    +1 if z < -STAT_ARB_THRESHOLD   (spread below mean → leg1 cheap vs leg2 → buy spread)
    -1 if z >  STAT_ARB_THRESHOLD   (spread above mean → leg1 dear vs leg2 → sell spread)
     0 otherwise

Signal lagged 1 day.
P/L per pair = signal × Δspread.
"""

import numpy as np
import pandas as pd
from config import (
    STAT_ARB_SPREADS, STAT_ARB_LOOKBACK,
    STAT_ARB_THRESHOLD, STAT_ARB_CONTRACT, INSTRUMENTS,
)
from data_loader import get_price_strip
from units import to_usd_bbl


def _get_f12_usd(raw: dict, ticker: str) -> pd.Series:
    """Return the F12 (or nearest available tenor) price in $/bbl."""
    strip = get_price_strip(raw, ticker)
    col   = f"F{STAT_ARB_CONTRACT}"
    if col not in strip.columns:
        # Fall back to F1
        col = "F1"
    price = strip[col].rename(ticker)
    return to_usd_bbl(price, ticker)


def stat_arb_signal(
    raw: dict,
    spread_name: str,
) -> tuple[pd.Series, pd.Series]:
    """
    Compute stat-arb signal for one spread pair.
    Returns (signal_series, spread_series).
    """
    leg1_ticker, leg2_ticker = STAT_ARB_SPREADS[spread_name]
    p1 = _get_f12_usd(raw, leg1_ticker)
    p2 = _get_f12_usd(raw, leg2_ticker)

    combined = pd.concat([p1, p2], axis=1).dropna()
    spread   = combined.iloc[:, 0] - combined.iloc[:, 1]
    spread.name = spread_name

    ma  = spread.rolling(STAT_ARB_LOOKBACK, min_periods=STAT_ARB_LOOKBACK // 2).mean()
    std = spread.rolling(STAT_ARB_LOOKBACK, min_periods=STAT_ARB_LOOKBACK // 2).std()

    # Relative z-score threshold
    threshold = STAT_ARB_THRESHOLD * std.abs()
    dev = spread - ma

    sig = np.where(dev < -threshold, 1,
           np.where(dev >  threshold, -1, 0))
    signal = pd.Series(sig, index=spread.index, dtype=float).fillna(0.0)
    return signal.shift(1).fillna(0.0).rename(spread_name), spread


def all_stat_arb_signals(
    raw: dict,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compute stat-arb signals for all spread pairs.
    Returns (signals_df, spreads_df) both indexed by Date.
    """
    signals_dict = {}
    spreads_dict = {}
    for spread_name in STAT_ARB_SPREADS:
        sig, spread = stat_arb_signal(raw, spread_name)
        signals_dict[spread_name] = sig
        spreads_dict[spread_name] = spread
    return pd.DataFrame(signals_dict), pd.DataFrame(spreads_dict)
