"""
Grid search over (n, ε) to maximise Sharpe ratio.
"""

import numpy as np
import pandas as pd
from itertools import product

from config import (
    LOOKBACK_GRID,
    INTRA_THRESHOLD_GRID,
    INTER_THRESHOLD_GRID,
)
from strategy import backtest_intra, backtest_inter


def optimize_intra(
    spread_name: str,
    spreads: pd.DataFrame,
    prices: pd.DataFrame,
) -> tuple[pd.DataFrame, dict, int, float]:
    """
    Grid search over LOOKBACK_GRID × INTRA_THRESHOLD_GRID.
    Returns (sharpe_heatmap, best_performance, best_n, best_eps).
    """
    results = {}
    best_sharpe = -np.inf
    best_params = (LOOKBACK_GRID[0], INTRA_THRESHOLD_GRID[0])
    best_perf   = {}
    best_pnl    = None

    for n, eps in product(LOOKBACK_GRID, INTRA_THRESHOLD_GRID):
        pnl, perf = backtest_intra(spread_name, spreads, prices, n, eps)
        results[(n, eps)] = perf["Sharpe Ratio"]
        if perf["Sharpe Ratio"] > best_sharpe:
            best_sharpe = perf["Sharpe Ratio"]
            best_params = (n, eps)
            best_perf   = perf
            best_pnl    = pnl

    # Build heatmap DataFrame  (rows=n, cols=eps)
    heatmap = pd.DataFrame(
        index=LOOKBACK_GRID,
        columns=INTRA_THRESHOLD_GRID,
        dtype=float,
    )
    for (n, eps), sr in results.items():
        heatmap.loc[n, eps] = sr
    heatmap.index.name   = "Lookback Period (days)"
    heatmap.columns.name = "Threshold"

    return heatmap, best_perf, best_params[0], best_params[1]


def optimize_inter(
    inter_name: str,
    spreads: pd.DataFrame,
    prices: pd.DataFrame,
) -> tuple[pd.DataFrame, dict, int, float]:
    """
    Grid search over LOOKBACK_GRID × INTER_THRESHOLD_GRID.
    Returns (sharpe_heatmap, best_performance, best_n, best_eps).
    """
    results = {}
    best_sharpe = -np.inf
    best_params = (LOOKBACK_GRID[0], INTER_THRESHOLD_GRID[0])
    best_perf   = {}
    best_pnl    = None

    for n, eps in product(LOOKBACK_GRID, INTER_THRESHOLD_GRID):
        pnl, perf = backtest_inter(inter_name, spreads, prices, n, eps)
        results[(n, eps)] = perf["Sharpe Ratio"]
        if perf["Sharpe Ratio"] > best_sharpe:
            best_sharpe = perf["Sharpe Ratio"]
            best_params = (n, eps)
            best_perf   = perf
            best_pnl    = pnl

    heatmap = pd.DataFrame(
        index=LOOKBACK_GRID,
        columns=INTER_THRESHOLD_GRID,
        dtype=float,
    )
    for (n, eps), sr in results.items():
        heatmap.loc[n, eps] = sr
    heatmap.index.name   = "Lookback Period (days)"
    heatmap.columns.name = "Threshold"

    return heatmap, best_perf, best_params[0], best_params[1]
