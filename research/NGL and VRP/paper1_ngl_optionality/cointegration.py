"""
Engle-Granger cointegration tests for Paper 1.

For intra-spreads:  coint(petrochem_$/gal, NGL_$/gal)
For inter-spreads:  coint(spread1, spread2)

Uses statsmodels.tsa.stattools.coint with trend='c' (constant, no trend),
which runs OLS regression then ADF on the residuals — the standard EG test.
"""

import numpy as np
import pandas as pd
from statsmodels.tsa.stattools import coint

from spreads import INTRA_SPREAD_DEFS, INTER_SPREAD_DEFS


def _run_coint(x: pd.Series, y: pd.Series) -> tuple[float, float]:
    """Return (test_stat, p_value) from Engle-Granger cointegration test."""
    xy = pd.concat([x, y], axis=1).dropna()
    stat, pval, _ = coint(xy.iloc[:, 0], xy.iloc[:, 1], trend="c", autolag="aic")
    return float(stat), float(pval)


def run_intra_coint(prices: pd.DataFrame) -> pd.DataFrame:
    """
    Test cointegration for all 5 intra-spreads using pairwise-complete data
    (only rows where BOTH legs are non-NaN, not all-5-complete).
    Returns DataFrame with columns: Spread, Statistic, P-Value, Cointegrated.
    """
    rows = []
    for name, (pet_col, _, ngl_col) in INTRA_SPREAD_DEFS.items():
        # Pairwise complete: drop only where this specific pair has NaNs
        pair = prices[[pet_col, ngl_col]].dropna()
        stat, pval = _run_coint(pair[pet_col], pair[ngl_col])
        rows.append({
            "Spread":       name,
            "Statistic":    round(stat, 6),
            "P-Value":      round(pval, 5),
            "Cointegrated": pval < 0.05,
        })
    return pd.DataFrame(rows).set_index("Spread")


def run_inter_coint(spreads: pd.DataFrame) -> pd.DataFrame:
    """
    Test cointegration for all inter-spread pairs using pairwise-complete data.
    Returns DataFrame with columns: Spread Pair, Statistic, P-Value, Cointegrated.
    """
    rows = []
    for pair_name, (s1, s2) in INTER_SPREAD_DEFS.items():
        if s1 not in spreads.columns or s2 not in spreads.columns:
            continue
        pair = spreads[[s1, s2]].dropna()
        stat, pval = _run_coint(pair[s1], pair[s2])
        rows.append({
            "Spread Pair":  pair_name,
            "Statistic":    round(stat, 6),
            "P-Value":      round(pval, 5),
            "Cointegrated": pval < 0.05,
        })
    return pd.DataFrame(rows).set_index("Spread Pair")


def get_cointegrated_intra(intra_results: pd.DataFrame) -> list[str]:
    return intra_results[intra_results["Cointegrated"]].index.tolist()


def get_cointegrated_inter(inter_results: pd.DataFrame) -> list[str]:
    return inter_results[inter_results["Cointegrated"]].index.tolist()
