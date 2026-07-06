"""
Risk model for Paper 2:
  - EWMA volatility (20-day half-life) for signal vol-scaling
  - EWMA covariance matrix (60-day half-life) for Dynamic Risk Parity
  - Ledoit-Wolf shrinkage (α=0.5) applied to the EWMA covariance

EWMA decay factor: λ = exp(-ln(2) / halflife)
"""

import numpy as np
import pandas as pd
from config import EWMA_VOL_HALFLIFE, EWMA_COV_HALFLIFE, LW_SHRINKAGE_ALPHA


# ── EWMA volatility ────────────────────────────────────────────────────────────

def ewma_vol(returns: pd.Series, halflife: int = EWMA_VOL_HALFLIFE) -> pd.Series:
    """
    EWMA standard deviation of daily returns.
    Returns a Series with the same index as `returns`.
    """
    var = returns.ewm(halflife=halflife, min_periods=halflife // 2).var()
    return np.sqrt(var).rename(returns.name)


def ewma_vol_all(returns: pd.DataFrame) -> pd.DataFrame:
    """EWMA vol for every column. Returns same shape as returns."""
    return returns.apply(lambda col: ewma_vol(col, EWMA_VOL_HALFLIFE))


# ── EWMA covariance ────────────────────────────────────────────────────────────

def ewma_cov(returns: pd.DataFrame, halflife: int = EWMA_COV_HALFLIFE) -> list[np.ndarray]:
    """
    Rolling EWMA covariance matrices.
    Returns list of (N×N) numpy arrays, one per row of `returns`.
    Index-aligned: result[i] corresponds to returns.index[i].
    """
    n = returns.shape[1]
    lam = np.exp(-np.log(2) / halflife)

    cov_mat = np.full((n, n), np.nan)
    results = []
    count   = 0

    for i, row in enumerate(returns.values):
        if np.any(np.isnan(row)):
            results.append(np.full((n, n), np.nan))
            continue

        x = row.reshape(-1, 1)
        if count == 0:
            cov_mat = x @ x.T
        else:
            cov_mat = lam * cov_mat + (1 - lam) * (x @ x.T)
        count += 1
        results.append(cov_mat.copy())

    return results


# ── Ledoit-Wolf shrinkage ─────────────────────────────────────────────────────

def ledoit_wolf_shrink(S: np.ndarray, alpha: float = LW_SHRINKAGE_ALPHA) -> np.ndarray:
    """
    Blend EWMA covariance S with its diagonal (shrinkage target = diag(S)).
    Shrunk = α × diag(S) + (1 - α) × S
    """
    if np.any(np.isnan(S)):
        return S
    target = np.diag(np.diag(S))
    return alpha * target + (1 - alpha) * S


def shrunk_cov_series(returns: pd.DataFrame) -> list[np.ndarray]:
    """Compute EWMA cov then apply LW shrinkage. Returns list aligned to returns index."""
    raw_covs = ewma_cov(returns)
    return [ledoit_wolf_shrink(C) for C in raw_covs]
