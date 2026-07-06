"""
Backtest engine for Paper 2.

Computes strategy returns for each signal type (momentum, carry, value, stat-arb)
and for each portfolio scheme (EW, DRP).

Performance metrics (log-return based):
  Information Ratio (IR) = AnnReturn / AnnVol  (= Sharpe assuming 0 risk-free rate)
  Annualised return = mean(daily_ret) * 252
  Annualised vol    = std(daily_ret) * sqrt(252)
  Max drawdown      = max(cummax - cum)
  Calmar ratio      = AnnReturn / MaxDD

Split-sample periods are defined in config.py.
"""

import numpy as np
import pandas as pd
from config import TRADING_DAYS_PER_YEAR, FULL_START, FULL_END, SUBSAMPLE_1, SUBSAMPLE_2
from portfolio.allocation import equal_weight_portfolio, dynamic_risk_parity_portfolio


def _performance(ret: pd.Series, label: str = "") -> dict:
    """Compute annualised performance metrics from a daily return series."""
    r = ret.dropna()
    if len(r) < 2:
        return {k: np.nan for k in ["Label", "IR", "AnnReturn", "AnnVol", "MaxDD", "Calmar", "N"]}

    ann_ret = r.mean() * TRADING_DAYS_PER_YEAR
    ann_vol = r.std()  * np.sqrt(TRADING_DAYS_PER_YEAR)
    ir      = ann_ret / ann_vol if ann_vol > 1e-10 else np.nan
    cum     = r.cumsum()
    max_dd  = (cum.cummax() - cum).max()
    calmar  = ann_ret / max_dd if max_dd > 1e-10 else np.nan

    return {
        "Label":     label,
        "IR":        round(ir,      4),
        "AnnReturn": round(ann_ret, 6),
        "AnnVol":    round(ann_vol, 6),
        "MaxDD":     round(max_dd,  6),
        "Calmar":    round(calmar,  4),
        "N":         len(r),
    }


def _period_perf(ret: pd.Series, label: str, start: str, end: str) -> dict:
    return _performance(ret.loc[start:end], label=f"{label} [{start[:4]}-{end[:4]}]")


def backtest_strategy(
    signals: pd.DataFrame,
    returns: pd.DataFrame,
    strategy_name: str,
) -> dict:
    """
    Run EW and DRP portfolios for a given signal DataFrame.
    Returns dict with keys: 'ew_returns', 'drp_returns', 'performance'.
    """
    ew_ret  = equal_weight_portfolio(signals, returns)
    drp_ret = dynamic_risk_parity_portfolio(signals, returns)

    periods = [
        ("Full sample 2015-2025",    FULL_START,       FULL_END),
        ("Pre-2022 (2015-2021)",     SUBSAMPLE_1[0],   SUBSAMPLE_1[1]),
        ("Post-2022 (2022-2025)",    SUBSAMPLE_2[0],   SUBSAMPLE_2[1]),
    ]

    rows = []
    for period_label, s, e in periods:
        rows.append(_period_perf(ew_ret,  f"{strategy_name} EW  – {period_label}", s, e))
        rows.append(_period_perf(drp_ret, f"{strategy_name} DRP – {period_label}", s, e))

    return {
        "ew_returns":  ew_ret,
        "drp_returns": drp_ret,
        "performance": pd.DataFrame(rows).set_index("Label"),
    }


def combined_portfolio(
    strategy_returns: dict[str, pd.Series],
    scheme: str = "EW",
) -> pd.Series:
    """
    Equal-weight blend of returns across strategy types.
    strategy_returns: {strategy_name: daily_return_series}
    """
    df = pd.DataFrame(strategy_returns)
    return df.mean(axis=1).rename(f"Combined_{scheme}")
