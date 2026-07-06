"""
Core backtesting engine for Paper 1.

Both intra- and inter-spread strategies share the same signal logic:
  signal_t = -1  if  (series_t - MA_t(series; n)) >  ε
           = +1  if  (series_t - MA_t(series; n)) < -ε
           =  0  otherwise

Signal is lagged by 1 day before P/L calculation (no look-ahead).

Rolling transaction cost:
  Applied on the last trading day of each calendar month.
  Only charged when |signal_prev| > 0 (position is open).
  Cost = ROLL_COST × |signal_prev| × sum(F1_prices of all legs, in $/gal).

For intra-spread (2 legs): legs are petrochem_$/gal and NGL_$/gal.
For inter-spread (4 legs): legs of spread1 + legs of spread2.
"""

import numpy as np
import pandas as pd
from config import ROLL_COST, TRADING_DAYS_PER_YEAR
from spreads import INTRA_SPREAD_DEFS, INTER_SPREAD_DEFS


# ─── helpers ──────────────────────────────────────────────────────────────────

def _signal(series: pd.Series, n: int, eps: float) -> pd.Series:
    """Return the raw (un-lagged) signal series."""
    ma  = series.rolling(n, min_periods=n).mean()
    dev = series - ma
    sig = np.where(dev > eps, -1, np.where(dev < -eps, 1, 0))
    return pd.Series(sig, index=series.index, dtype=float)


def _eom_dates(index: pd.DatetimeIndex) -> pd.DatetimeIndex:
    """Return dates that are the last trading day in each calendar month."""
    s = pd.Series(index, index=index)
    return s.groupby([s.dt.year, s.dt.month]).last().values


def _leg_f1_sum(spread_name: str, prices: pd.DataFrame) -> pd.Series:
    """Sum of F1 prices of both legs for an intra-spread ($/gal)."""
    pet_col, _, ngl_col = INTRA_SPREAD_DEFS[spread_name]
    return prices[pet_col] + prices[ngl_col]


def _inter_leg_f1_sum(inter_name: str, prices: pd.DataFrame) -> pd.Series:
    """Sum of F1 prices of all 4 legs for an inter-spread ($/gal)."""
    s1_name, s2_name = INTER_SPREAD_DEFS[inter_name]
    pet1, _, ngl1 = INTRA_SPREAD_DEFS[s1_name]
    pet2, _, ngl2 = INTRA_SPREAD_DEFS[s2_name]
    return prices[pet1] + prices[ngl1] + prices[pet2] + prices[ngl2]


# ─── performance metrics ──────────────────────────────────────────────────────

def _performance(daily_pnl: pd.Series) -> dict:
    """Compute Sharpe, annualized P/L, max drawdown, and RoD."""
    pnl = daily_pnl.fillna(0.0)
    ann_pnl  = pnl.mean() * TRADING_DAYS_PER_YEAR
    vol      = pnl.std()
    sharpe   = (pnl.mean() / vol * np.sqrt(TRADING_DAYS_PER_YEAR)) if vol > 1e-12 else 0.0
    cum      = pnl.cumsum()
    max_dd   = (cum.cummax() - cum).max()
    rod      = (ann_pnl / max_dd) if max_dd > 1e-12 else np.nan
    return {
        "Sharpe Ratio":   round(sharpe,  5),
        "Annualized P/L": round(ann_pnl, 6),   # $/gal
        "Max Drawdown":   round(max_dd,  6),
        "RoD":            round(rod,     4),
    }


# ─── intra-spread backtest ────────────────────────────────────────────────────

def backtest_intra(
    spread_name: str,
    spreads: pd.DataFrame,
    prices: pd.DataFrame,
    n: int,
    eps: float,
) -> tuple[pd.Series, dict]:
    """
    Run intra-spread strategy.
    Returns (daily_pnl series, performance dict).
    """
    s        = spreads[spread_name].dropna()
    leg_sum  = _leg_f1_sum(spread_name, prices).reindex(s.index)

    raw_sig  = _signal(s, n, eps)
    sig      = raw_sig.shift(1).fillna(0.0)            # lag 1 day (no look-ahead)

    pnl      = sig * s.diff()
    pnl.iloc[0] = 0.0

    # Roll cost on last trading day of each month
    eom = pd.DatetimeIndex(_eom_dates(s.index))
    for d in eom:
        if d in pnl.index and abs(sig.loc[d]) > 0:
            pnl.loc[d] -= ROLL_COST * abs(sig.loc[d]) * leg_sum.reindex([d]).iloc[0]

    perf = _performance(pnl)
    return pnl.rename(spread_name), perf


# ─── inter-spread backtest ────────────────────────────────────────────────────

def backtest_inter(
    inter_name: str,
    spreads: pd.DataFrame,
    prices: pd.DataFrame,
    n: int,
    eps: float,
) -> tuple[pd.Series, dict]:
    """
    Run inter-spread strategy on R = spread1 - spread2.
    Returns (daily_pnl series, performance dict).
    """
    s1_name, s2_name = INTER_SPREAD_DEFS[inter_name]
    combined = pd.concat([spreads[s1_name], spreads[s2_name]], axis=1).dropna()
    r        = combined.iloc[:, 0] - combined.iloc[:, 1]
    r.name   = inter_name

    leg_sum  = _inter_leg_f1_sum(inter_name, prices).reindex(r.index)

    raw_sig  = _signal(r, n, eps)
    sig      = raw_sig.shift(1).fillna(0.0)

    pnl      = sig * r.diff()
    pnl.iloc[0] = 0.0

    eom = pd.DatetimeIndex(_eom_dates(r.index))
    for d in eom:
        if d in pnl.index and abs(sig.loc[d]) > 0:
            pnl.loc[d] -= ROLL_COST * abs(sig.loc[d]) * leg_sum.reindex([d]).iloc[0]

    perf = _performance(pnl)
    return pnl.rename(inter_name), perf
