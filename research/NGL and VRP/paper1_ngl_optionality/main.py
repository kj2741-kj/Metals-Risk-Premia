"""
Paper 1 – NGL Optionality: full pipeline.

Run from the paper1_ngl_optionality directory:
    python main.py

Outputs written to output/:
    charts/  – all PNG figures
    intra_coint_results.csv
    inter_coint_results.csv
    intra_perf_summary.csv
    inter_perf_summary.csv
    intra_sharpe_heatmaps/  – one CSV per spread
    inter_sharpe_heatmaps/  – one CSV per inter-spread
"""

import sys
import warnings
from pathlib import Path

import pandas as pd

# Allow imports from same directory when run directly
sys.path.insert(0, str(Path(__file__).parent))

warnings.filterwarnings("ignore")

from config import NGL_FILE, OUTPUT_DIR, CHARTS_DIR
from data_loader import load_ngl_data
from spreads import align_all
from cointegration import (
    run_intra_coint, run_inter_coint,
    get_cointegrated_intra, get_cointegrated_inter,
)
from optimize import optimize_intra, optimize_inter
from strategy import backtest_intra, backtest_inter
from spreads import INTRA_SPREAD_DEFS, INTER_SPREAD_DEFS
from plots import (
    plot_spreads,
    plot_intra_equity,
    plot_intra_focus,
    plot_inter_equity,
    plot_inter_focus,
    plot_coint_summary,
)


def _ensure_dirs() -> None:
    for d in [OUTPUT_DIR, CHARTS_DIR,
              OUTPUT_DIR / "intra_sharpe_heatmaps",
              OUTPUT_DIR / "inter_sharpe_heatmaps"]:
        d.mkdir(parents=True, exist_ok=True)


def main() -> None:
    _ensure_dirs()

    # ── 1. Load data ──────────────────────────────────────────────────────────
    print("Loading NGL data …")
    raw = load_ngl_data(NGL_FILE)
    prices, f2_prices, spreads = align_all(raw)
    print(f"  Sample: {prices.index[0].date()} to {prices.index[-1].date()}  "
          f"({len(prices):,} trading days)")

    # ── 2. Figure 2: spread time series ──────────────────────────────────────
    print("Plotting spreads (Figure 2) …")
    plot_spreads(spreads)

    # ── 3. Cointegration tests ─────────────────────────────────────────────────
    print("Running cointegration tests …")
    intra_coint = run_intra_coint(prices)
    inter_coint = run_inter_coint(spreads)

    intra_coint.to_csv(OUTPUT_DIR / "intra_coint_results.csv")
    inter_coint.to_csv(OUTPUT_DIR / "inter_coint_results.csv")

    coint_intra_list = get_cointegrated_intra(intra_coint)
    coint_inter_list = get_cointegrated_inter(inter_coint)

    print("\n  Cointegrated intra-spreads:")
    print(intra_coint.to_string())
    print("\n  Cointegrated inter-spreads:")
    print(inter_coint.to_string())

    plot_coint_summary(intra_coint, inter_coint)

    # ── 4. Grid search – intra ─────────────────────────────────────────────────
    print("\nOptimising intra-spread strategies …")
    intra_best_pnls: dict[str, pd.Series] = {}
    intra_perf_rows = []

    for spread_name in INTRA_SPREAD_DEFS:
        print(f"  {spread_name} …", end="", flush=True)
        heatmap, best_perf, best_n, best_eps = optimize_intra(
            spread_name, spreads, prices
        )
        heatmap.to_csv(OUTPUT_DIR / "intra_sharpe_heatmaps" / f"{spread_name.replace(' ', '_')}.csv")

        # Re-run with best params to get the P/L series
        pnl, _ = backtest_intra(spread_name, spreads, prices, best_n, best_eps)
        intra_best_pnls[spread_name] = pnl

        row = {"Spread": spread_name, "Best n": best_n, "Best eps": best_eps}
        row.update(best_perf)
        intra_perf_rows.append(row)

        # Figure 4 for every spread (named fig4_*)
        plot_intra_focus(spread_name, pnl, heatmap, best_n, best_eps)
        print(f" Sharpe={best_perf['Sharpe Ratio']:.3f}  n={best_n}  eps={best_eps:.2f}")

    intra_perf = pd.DataFrame(intra_perf_rows).set_index("Spread")
    intra_perf.to_csv(OUTPUT_DIR / "intra_perf_summary.csv")

    print("\n  Intra-spread performance summary:")
    print(intra_perf.to_string())

    # ── 5. Figure 3: intra equity curves ─────────────────────────────────────
    plot_intra_equity(intra_best_pnls)

    # ── 6. Grid search – inter ─────────────────────────────────────────────────
    print("\nOptimising inter-spread strategies …")
    inter_best_pnls: dict[str, pd.Series] = {}
    inter_perf_rows = []

    for inter_name in INTER_SPREAD_DEFS:
        # Both constituent spreads must exist in spreads DataFrame
        s1, s2 = INTER_SPREAD_DEFS[inter_name]
        if s1 not in spreads.columns or s2 not in spreads.columns:
            print(f"  Skipping {inter_name} (missing constituent spread data)")
            continue

        print(f"  {inter_name} …", end="", flush=True)
        heatmap, best_perf, best_n, best_eps = optimize_inter(
            inter_name, spreads, prices
        )
        heatmap.to_csv(OUTPUT_DIR / "inter_sharpe_heatmaps" / f"{inter_name.replace(' ', '_')}.csv")

        pnl, _ = backtest_inter(inter_name, spreads, prices, best_n, best_eps)
        inter_best_pnls[inter_name] = pnl

        row = {"Inter-Spread": inter_name, "Best n": best_n, "Best eps": best_eps}
        row.update(best_perf)
        inter_perf_rows.append(row)

        # Figure 6 for every inter-spread
        plot_inter_focus(inter_name, pnl, heatmap, best_n, best_eps)
        print(f" Sharpe={best_perf['Sharpe Ratio']:.3f}  n={best_n}  eps={best_eps:.2f}")

    inter_perf = pd.DataFrame(inter_perf_rows).set_index("Inter-Spread")
    inter_perf.to_csv(OUTPUT_DIR / "inter_perf_summary.csv")

    print("\n  Inter-spread performance summary:")
    print(inter_perf.to_string())

    # ── 7. Figure 5: inter equity curves ─────────────────────────────────────
    if inter_best_pnls:
        plot_inter_equity(inter_best_pnls)

    # ── 8. Summary ────────────────────────────────────────────────────────────
    print(f"\nAll outputs written to: {OUTPUT_DIR.resolve()}")
    print("Charts:")
    for p in sorted(CHARTS_DIR.glob("*.png")):
        print(f"  {p.name}")


if __name__ == "__main__":
    main()
