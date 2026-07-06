"""
Portfolio allocation schemes for Paper 2.

Two schemes:
  1. Equal Weight (EW)     – each instrument gets weight 1/N after vol-scaling
  2. Dynamic Risk Parity (DRP) – weights inversely proportional to EWMA variance,
                                  further adjusted with Ledoit-Wolf covariance

Both operate on *vol-scaled* signals:
  scaled_signal_i = signal_i / ewma_vol_i

Then:
  EW     : w_i = scaled_signal_i / N
  DRP    : solve min w'Σw  s.t. Σ w_i = 1  (risk-parity via inv-variance approximation)
           w_i = (1/σ²_i) / Σ(1/σ²_j)     then multiply by scaled_signal direction
"""

import numpy as np
import pandas as pd
from config import INSTRUMENTS, TRADING_DAYS_PER_YEAR
from portfolio.risk_model import ewma_vol_all, shrunk_cov_series


def _vol_scale_signals(
    signals: pd.DataFrame,
    returns: pd.DataFrame,
) -> pd.DataFrame:
    """
    Divide each signal column by its EWMA vol (annualised to 1).
    Returns vol-scaled signals, same shape.
    """
    vols = ewma_vol_all(returns) * np.sqrt(TRADING_DAYS_PER_YEAR)
    vols = vols.reindex(signals.index)
    # Avoid division by zero/NaN
    vols = vols.replace(0, np.nan)
    return signals.div(vols).fillna(0.0)


def equal_weight_portfolio(
    signals: pd.DataFrame,
    returns: pd.DataFrame,
) -> pd.Series:
    """
    Equal-weight portfolio: each instrument gets 1/N weight.
    No vol-scaling at position level — paper uses plain signal/N × return.
    Returns a pd.Series of daily log-returns.
    """
    n = signals.shape[1]
    common = signals.index.intersection(returns.index)
    sig  = signals.loc[common]
    rets = returns.loc[common]
    port_ret = (sig / n * rets).sum(axis=1)
    return port_ret.rename("EW")


def dynamic_risk_parity_portfolio(
    signals: pd.DataFrame,
    returns: pd.DataFrame,
) -> pd.Series:
    """
    Dynamic Risk Parity portfolio using EWMA+LW covariance.
    Weights: w_i ∝ 1/σ²_i × signal_direction_i
    Returns a pd.Series of portfolio log-returns.
    """
    common  = signals.index.intersection(returns.index)
    signals = signals.loc[common]
    rets    = returns.loc[common]

    vols    = ewma_vol_all(rets) * np.sqrt(TRADING_DAYS_PER_YEAR)
    cov_list = shrunk_cov_series(rets)

    port_rets = []
    for i, date in enumerate(common):
        sig  = signals.loc[date].values
        var  = np.diag(cov_list[i]) if not np.any(np.isnan(cov_list[i])) else (vols.loc[date].values ** 2)
        var  = np.where(var <= 0, np.nan, var)

        # Inverse-variance weights with signal direction
        inv_var = np.where(np.isnan(var), 0.0, 1.0 / var)
        total   = inv_var.sum()
        if total < 1e-12:
            w = np.zeros_like(sig)
        else:
            w = inv_var / total * np.sign(sig)

        r    = rets.loc[date].values
        daily_ret = float(np.nansum(w * r))
        port_rets.append(daily_ret)

    return pd.Series(port_rets, index=common, name="DRP")
