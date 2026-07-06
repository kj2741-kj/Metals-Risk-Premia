"""
Paper 2 – Risk Premia in Diversified Energy Portfolios: full pipeline.

Run from the paper2_energy_risk_premia directory:
    python main.py

Outputs written to output/:
    charts/                  – PNG figures
    continuous_prices.csv    – roll-adjusted prices ($/bbl equiv.)
    momentum_signals.csv
    carry_signals.csv
    value_signals.csv
    statarb_signals.csv
    statarb_spreads.csv
    perf_momentum.csv
    perf_carry.csv
    perf_value.csv
    perf_statarb.csv
    perf_combined.csv
"""

import sys
import warnings
from pathlib import Path

import pandas as pd
import numpy as np

sys.path.insert(0, str(Path(__file__).parent))
warnings.filterwarnings("ignore")

from config import OUTPUT_DIR, CHARTS_DIR, DATA_START, FULL_START, FULL_END, INSTRUMENTS
from data_loader import load_all_data, get_price_strip
from rolling import build_all_continuous, build_all_returns, load_expiry_calendar
from units import convert_all
from signals.momentum import all_momentum_signals
from signals.carry import all_carry_signals
from signals.value import all_value_signals
from signals.stat_arb import all_stat_arb_signals
from portfolio.backtest import backtest_strategy, combined_portfolio
from plots import (
    plot_continuous_prices,
    plot_momentum_equity,
    plot_carry_equity,
    plot_value_equity,
    plot_statarb_equity,
    plot_combined_equity,
    plot_correlation_heatmap,
    plot_ir_bars,
)


def _ensure_dirs() -> None:
    for d in [OUTPUT_DIR, CHARTS_DIR]:
        d.mkdir(parents=True, exist_ok=True)


def _log_returns(raw: dict, cal: pd.DataFrame) -> pd.DataFrame:
    """
    Build daily log-returns from DATA_START (2010) for full initialization depth,
    then the backtest slice FULL_START:FULL_END is extracted at use time.
    """
    ret_dict = build_all_returns(raw, cal)
    df = pd.DataFrame(ret_dict)
    return df.loc[DATA_START:FULL_END]


def main() -> None:
    _ensure_dirs()

    # ── 1. Load raw data ──────────────────────────────────────────────────────
    print("Loading raw data …")
    raw = load_all_data()
    cal = load_expiry_calendar()

    # ── 2. Build roll-adjusted continuous series ──────────────────────────────
    print("Building roll-adjusted continuous price series …")
    cont_raw = build_all_continuous(raw, cal)

    # ── 3. Convert to $/bbl ───────────────────────────────────────────────────
    print("Converting units to $/bbl equivalent …")
    cont = convert_all(cont_raw)

    # Save continuous prices
    prices_df = pd.DataFrame({t: adj for t, (adj, _) in cont.items()})
    prices_df = prices_df.loc[FULL_START:FULL_END]
    prices_df.to_csv(OUTPUT_DIR / "continuous_prices.csv")

    # Figure 1
    plot_continuous_prices(cont)
    print("  Figure 1 written.")

    # ── 4. Log returns (from front contract prices, roll costs on roll days) ───
    returns = _log_returns(raw, cal)

    # ── 5. Signals ────────────────────────────────────────────────────────────
    print("Computing signals …")

    mom_signals = all_momentum_signals(cont)
    car_signals = all_carry_signals(raw)
    val_signals = all_value_signals(raw)
    sa_signals, sa_spreads = all_stat_arb_signals(raw)

    mom_signals.to_csv(OUTPUT_DIR / "momentum_signals.csv")
    car_signals.to_csv(OUTPUT_DIR / "carry_signals.csv")
    val_signals.to_csv(OUTPUT_DIR / "value_signals.csv")
    sa_signals.to_csv(OUTPUT_DIR / "statarb_signals.csv")
    sa_spreads.to_csv(OUTPUT_DIR / "statarb_spreads.csv")

    print(f"  Signals shape — Momentum: {mom_signals.shape}, "
          f"Carry: {car_signals.shape}, Value: {val_signals.shape}, "
          f"StatArb: {sa_signals.shape}")

    # ── 6. Backtests ──────────────────────────────────────────────────────────
    print("Running backtests …")

    # Instruments-level returns — restrict to formal backtest window
    ret_cols = [t for t in INSTRUMENTS if t in returns.columns]
    returns_aligned = returns[ret_cols].loc[FULL_START:FULL_END]

    # Stat-arb returns: signal × (r_leg1 - r_leg2)
    # Using log returns for compatibility with other strategies (dimensionless %-return)
    from config import STAT_ARB_SPREADS
    sa_rets_dict = {}
    for col, (leg1, leg2) in STAT_ARB_SPREADS.items():
        if col not in sa_signals.columns:
            continue
        if leg1 not in returns_aligned.columns or leg2 not in returns_aligned.columns:
            continue
        sig = sa_signals[col].reindex(returns_aligned.index).fillna(0.0)
        spread_ret = returns_aligned[leg1] - returns_aligned[leg2]
        sa_rets_dict[col] = (sig * spread_ret).fillna(0.0)

    # Align signal DataFrames to returns index
    def _align(sig: pd.DataFrame) -> pd.DataFrame:
        cols = [c for c in sig.columns if c in ret_cols]
        return sig[cols].reindex(returns_aligned.index).fillna(0.0)

    mom_bt  = backtest_strategy(_align(mom_signals), returns_aligned, "Momentum")
    car_bt  = backtest_strategy(_align(car_signals), returns_aligned, "Carry")
    val_bt  = backtest_strategy(_align(val_signals), returns_aligned, "Value")

    # Stat-arb backtest: treat each spread as its own return series
    sa_returns_df = pd.DataFrame(sa_rets_dict).loc[FULL_START:FULL_END]
    # For the stat-arb portfolio we use EW across spreads
    sa_ew_ret  = sa_returns_df.mean(axis=1).rename("StatArb_EW")
    sa_drp_ret = sa_ew_ret.copy().rename("StatArb_DRP")  # same: single-spread has no DRP distinction

    # Save performance tables
    mom_bt["performance"].to_csv(OUTPUT_DIR / "perf_momentum.csv")
    car_bt["performance"].to_csv(OUTPUT_DIR / "perf_carry.csv")
    val_bt["performance"].to_csv(OUTPUT_DIR / "perf_value.csv")

    # ── 7. Plots ──────────────────────────────────────────────────────────────
    print("Generating figures …")

    plot_momentum_equity(mom_bt["ew_returns"], mom_bt["drp_returns"])
    plot_carry_equity   (car_bt["ew_returns"], car_bt["drp_returns"])
    plot_value_equity   (val_bt["ew_returns"], val_bt["drp_returns"])
    plot_statarb_equity (sa_rets_dict)

    # Combined portfolio (EW blend across 4 strategies)
    strategy_ew_rets = {
        "Momentum":  mom_bt["ew_returns"],
        "Carry":     car_bt["ew_returns"],
        "Value":     val_bt["ew_returns"],
        "Stat-Arb":  sa_ew_ret,
    }
    combo = combined_portfolio(strategy_ew_rets, scheme="EW")
    plot_combined_equity(strategy_ew_rets, combo)
    plot_correlation_heatmap({**strategy_ew_rets, "Combined": combo})

    # IR bar chart
    all_perf = pd.concat([
        mom_bt["performance"],
        car_bt["performance"],
        val_bt["performance"],
    ])
    all_perf.to_csv(OUTPUT_DIR / "perf_combined.csv")
    plot_ir_bars(all_perf)

    # ── 8. Summary ────────────────────────────────────────────────────────────
    print(f"\nAll outputs written to: {OUTPUT_DIR.resolve()}")
    print("Charts:")
    for p in sorted(CHARTS_DIR.glob("*.png")):
        print(f"  {p.name}")

    print("\nMomentum performance (EW):")
    ew_rows = mom_bt["performance"][mom_bt["performance"].index.str.contains("EW")]
    print(ew_rows[["IR", "AnnReturn", "AnnVol", "MaxDD"]].to_string())


if __name__ == "__main__":
    main()
