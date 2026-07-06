"""
Roll-adjusted continuous price series for Paper 2.

Rolling logic:
  Instruments with expiry calendar (CL, CO, XB, HO, NG, QS):
    Roll on LAST_TRADEABLE_DT of the front contract.
    On roll day: sell F1, buy F2.

  NGL swaps (CAP, BAP, DAE):
    No exact expiry — roll on last trading day of each calendar month.
    Always use F2 as effective front contract (skip averaging F1).

Adjustment method: RATIO backward adjustment.
  On each roll: multiply ALL prior prices by (F2_close / F1_close).
  This keeps prices positive and ensures log-returns are continuous.
  Roll transaction cost is subtracted as a discrete return on roll dates.
"""

import numpy as np
import pandas as pd
from pathlib import Path

from config import (
    EXPIRY_FILE, EXPIRY_TICKERS, NGL_TICKERS,
    DATA_START, FULL_START, FULL_END, ROLL_COST_BPS,
)
from data_loader import load_all_data, get_price_strip


# ── Expiry calendar ────────────────────────────────────────────────────────────

def load_expiry_calendar(filepath: Path | None = None) -> pd.DataFrame:
    fp = filepath or EXPIRY_FILE
    cal = pd.read_excel(fp)
    cal.columns = [c.strip() for c in cal.columns]
    cal["LAST_TRADEABLE_DT"] = pd.to_datetime(cal["LAST_TRADEABLE_DT"], errors="coerce")
    cal = cal.dropna(subset=["LAST_TRADEABLE_DT"])

    if "Ticker" in cal.columns:
        cal["root"] = cal["Ticker"].str[:2].str.upper()
    elif "Contract" in cal.columns:
        cal["root"] = cal["Contract"].str[:2].str.upper()
    else:
        raise ValueError("Cannot find Ticker/Contract column in expiry calendar")

    return cal[["root", "Year", "Month", "LAST_TRADEABLE_DT"]].copy()


def _get_roll_dates(cal: pd.DataFrame, ticker: str, idx: pd.DatetimeIndex | None = None) -> pd.DatetimeIndex:
    """Return roll dates from calendar; fall back to EOM dates if ticker not found."""
    sub = cal[cal["root"] == ticker].sort_values("LAST_TRADEABLE_DT")
    if len(sub) > 0:
        return pd.DatetimeIndex(sub["LAST_TRADEABLE_DT"].values)
    # Fall back to end-of-month roll dates derived from price index
    if idx is not None:
        return _eom_dates_for(idx)
    return pd.DatetimeIndex([])


# ── EOM helper for NGL swaps ──────────────────────────────────────────────────

def _eom_dates_for(index: pd.DatetimeIndex) -> pd.DatetimeIndex:
    s = pd.Series(index, index=index)
    return pd.DatetimeIndex(s.groupby([s.dt.year, s.dt.month]).last().values)


# ── Snap to nearest available trading day ─────────────────────────────────────

def _snap_date(dt: pd.Timestamp, idx: pd.DatetimeIndex) -> pd.Timestamp | None:
    pos = idx.searchsorted(dt, side="right") - 1
    if pos < 0:
        return None
    return idx[pos]


# ── Core roll engine (ratio adjustment) ────────────────────────────────────────

def build_continuous_series(
    strip: pd.DataFrame,
    roll_dates: pd.DatetimeIndex | None,
    is_ngl: bool = False,
) -> tuple[pd.Series, pd.Series]:
    """
    Build a ratio-backward-adjusted continuous price series.

    Ratio adjustment: on each roll date, multiply ALL prior prices by (F2/F1).
    This keeps prices positive and gives correct daily percentage returns.

    Returns
    -------
    (adjusted, unadjusted)  — both pd.Series indexed by Date in raw Bloomberg units.
    """
    strip = strip.loc[DATA_START:FULL_END].copy()
    idx   = strip.index

    if is_ngl:
        front_col = "F2"
        next_col  = "F3"
        roll_dts  = _eom_dates_for(idx)
    else:
        front_col = "F1"
        next_col  = "F2"
        roll_dts  = roll_dates

    if front_col not in strip.columns:
        raise ValueError(f"Column {front_col} missing in strip")

    unadjusted = strip[front_col].copy()
    adjusted   = strip[front_col].copy()

    for rd in roll_dts:
        snapped = _snap_date(rd, idx)
        if snapped is None or snapped >= idx[-1]:
            continue

        f1 = strip[front_col].loc[snapped]
        if next_col not in strip.columns:
            continue
        f2 = strip[next_col].loc[snapped]

        if pd.isna(f1) or pd.isna(f2) or f1 <= 0 or f2 <= 0:
            continue

        ratio = f2 / f1  # > 1 in contango, < 1 in backwardation
        # Multiply all prices UP TO AND INCLUDING roll date by this ratio
        # so the transition from pre-roll to post-roll is seamless
        adjusted.loc[:snapped] *= ratio

    return adjusted, unadjusted


def build_return_series(
    strip: pd.DataFrame,
    roll_dates: pd.DatetimeIndex | None,
    is_ngl: bool = False,
) -> pd.Series:
    """
    Build daily log-return series from the front contract.

    On roll dates: return = -ROLL_COST_BPS (just the transaction cost).
    Between rolls: return = log(F1_t / F1_{t-1}).

    Returns pd.Series of daily log-returns.
    """
    strip = strip.loc[DATA_START:FULL_END].copy()
    idx   = strip.index

    if is_ngl:
        front_col = "F2"
        roll_dts  = set(_eom_dates_for(idx))
    else:
        front_col = "F1"
        roll_dts_idx = roll_dates if roll_dates is not None else pd.DatetimeIndex([])
        # Snap all roll dates to actual trading days
        snapped = set()
        for rd in roll_dts_idx:
            s = _snap_date(rd, idx)
            if s is not None:
                snapped.add(s)
        roll_dts = snapped

    if front_col not in strip.columns:
        return pd.Series(dtype=float)

    prices = strip[front_col].copy()
    log_ret = np.log(prices / prices.shift(1))

    # Replace roll-day returns with just the negative transaction cost
    for rd in roll_dts:
        if rd in log_ret.index:
            log_ret.loc[rd] = -ROLL_COST_BPS

    return log_ret.replace([np.inf, -np.inf], np.nan)


def build_all_continuous(
    raw: dict | None = None,
    cal: pd.DataFrame | None = None,
) -> dict[str, tuple[pd.Series, pd.Series]]:
    """
    Returns {ticker: (adjusted_series, unadjusted_series)} for all instruments.
    Prices in raw Bloomberg units (not yet converted to $/bbl).
    """
    if raw is None:
        raw = load_all_data()
    if cal is None:
        cal = load_expiry_calendar()

    result: dict[str, tuple[pd.Series, pd.Series]] = {}

    for ticker in list(EXPIRY_TICKERS) + list(NGL_TICKERS):
        strip      = get_price_strip(raw, ticker)
        is_ngl     = ticker in NGL_TICKERS
        roll_dates = None if is_ngl else _get_roll_dates(cal, ticker, idx=strip.index)

        adj, unadj = build_continuous_series(strip, roll_dates, is_ngl=is_ngl)
        adj.name   = ticker
        unadj.name = ticker
        result[ticker] = (adj, unadj)

    return result


def build_all_returns(
    raw: dict | None = None,
    cal: pd.DataFrame | None = None,
) -> dict[str, pd.Series]:
    """
    Returns {ticker: log_return_series} for all instruments.
    Returns already account for roll costs.
    """
    if raw is None:
        raw = load_all_data()
    if cal is None:
        cal = load_expiry_calendar()

    result: dict[str, pd.Series] = {}

    for ticker in list(EXPIRY_TICKERS) + list(NGL_TICKERS):
        strip      = get_price_strip(raw, ticker)
        is_ngl     = ticker in NGL_TICKERS
        roll_dates = None if is_ngl else _get_roll_dates(cal, ticker, idx=strip.index)

        ret = build_return_series(strip, roll_dates, is_ngl=is_ngl)
        ret.name = ticker
        result[ticker] = ret

    return result


def get_strip_at_tenor(
    raw: dict,
    ticker: str,
    tenor: int,
) -> pd.Series:
    """Return the Fn price series for a specific tenor."""
    strip = get_price_strip(raw, ticker)
    col   = f"F{tenor}"
    if col not in strip.columns:
        return pd.Series(dtype=float, name=f"{ticker}_F{tenor}")
    return strip[col].rename(f"{ticker}_F{tenor}")
