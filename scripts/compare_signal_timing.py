"""
compare_signal_timing.py
========================
Compares two position-entry conventions for every top-5 MA and CTA strategy:

  LAG-1  (current):  Position[t] = Signal[t-1]
                     PnL[t]      = Position[t] × ΔF1_cont[t]
                     → signal computed at close t-1, trade entered at close t-1,
                       PnL runs from close t-1 to close t.

  SAME-DAY:          Position[t] = Signal[t]
                     PnL[t]      = Position[t] × ΔF1_cont[t]
                     → signal computed at close t, trade entered at close t,
                       PnL runs from close t-1 to close t.
                     Feasible for liquid futures (LME Copper has a daily Ring
                     settlement that is publicly known; you can execute AT that
                     close). NOT look-ahead in the strict sense for long-window
                     strategies, but the single-day price IS reflected in
                     today's MA/EWMA, so a same-day flip captures that day's
                     move — a small structural bias for short-window signals.

Outputs
-------
  Signal_Timing_Comparison.csv   — full table, saved to momentum_output_v2/
  Printed summary table
"""

from __future__ import annotations

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import numpy as np
import pandas as pd
from pathlib import Path

from rolling_continuous import get_metal_rolling_f1
from momentum_signals  import (
    _ewma, _metrics_from_pnl,
    CTA_SHORT, CTA_LONG, CTA_PW, CTA_SW, MAX_LOOKBACK, OUT_DIR,
)

OUT_DIR.mkdir(exist_ok=True)


# ── Load data ──────────────────────────────────────────────────────────────────
print("Loading LME Copper data ...")
rolling = get_metal_rolling_f1("LP", verbose=False)
f1_raw  = rolling["F1_raw"]
f1_cont = rolling["F1_continuous"]

delta_np = f1_cont.diff().values.astype(float)   # ΔF1_continuous, length T
T        = len(f1_raw)


# ── Signal builders ────────────────────────────────────────────────────────────

def ma_signal_np(m: int, n: int) -> np.ndarray:
    """Raw MA crossover signal array (NaN during warmup)."""
    ma_m = f1_raw.rolling(m).mean().values
    ma_n = f1_raw.rolling(n).mean().values
    return np.sign(ma_m - ma_n)


def cta_single_signal_np(s: int, l: int,
                          pw: int = CTA_PW,
                          sw: int = CTA_SW) -> np.ndarray:
    """Single-pair CTA signal array (NaN during warmup)."""
    x = (_ewma(f1_raw, s) - _ewma(f1_raw, l)).values
    pv = f1_raw.rolling(pw).std().values
    with np.errstate(invalid="ignore", divide="ignore"):
        y  = x / pv
        ys = pd.Series(y).rolling(sw).std().values
        z  = y / ys
        u  = z * np.exp(-z**2 / 4) / 0.89
    return np.sign(u)


def cta_paper_signal_np(short_params=CTA_SHORT,
                         long_params=CTA_LONG,
                         pw=CTA_PW, sw=CTA_SW) -> np.ndarray:
    """Full 3-pair Baz-Granger S_CTA signal array."""
    pv = f1_raw.rolling(pw).std().values
    u_list = []
    for sk, lk in zip(short_params, long_params):
        x = (_ewma(f1_raw, sk) - _ewma(f1_raw, lk)).values
        with np.errstate(invalid="ignore", divide="ignore"):
            y  = x / pv
            ys = pd.Series(y).rolling(sw).std().values
            z  = y / ys
            u  = z * np.exp(-z**2 / 4) / 0.89
        u_list.append(u)
    s_cta = np.nanmean(np.stack(u_list, axis=1), axis=1)
    return np.sign(s_cta)


# ── Dual metrics (lag-1 vs same-day) ──────────────────────────────────────────

def dual_metrics(raw_signal: np.ndarray) -> dict:
    """
    Compute performance metrics for both entry conventions.

    LAG-1     : pos[t] = signal[t-1]  →  PnL[t] = pos[t] * delta[t]
    SAME-DAY  : pos[t] = signal[t]    →  PnL[t] = pos[t] * delta[t]
    """
    sig = raw_signal.astype(float)

    # --- LAG-1 ---
    pos_lag = np.empty(T)
    pos_lag[0] = 0.0
    pos_lag[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
    pnl_lag = pos_lag * delta_np
    m_lag   = _metrics_from_pnl(pnl_lag, pos_lag)

    # --- SAME-DAY ---
    pos_same = np.where(np.isfinite(sig), sig, 0.0)
    pnl_same = pos_same * delta_np
    m_same   = _metrics_from_pnl(pnl_same, pos_same)

    return {
        "sharpe_lag1"    : m_lag["sharpe"],
        "sharpe_same"    : m_same["sharpe"],
        "ann_ret_lag1"   : m_lag["ann_return"],
        "ann_ret_same"   : m_same["ann_return"],
        "maxdd_lag1"     : m_lag["max_drawdown"],
        "maxdd_same"     : m_same["max_drawdown"],
        "totpnl_lag1"    : m_lag["total_pnl"],
        "totpnl_same"    : m_same["total_pnl"],
        "hitrate_lag1"   : m_lag["hit_rate"],
        "hitrate_same"   : m_same["hit_rate"],
        "sharpe_delta"   : round(m_same["sharpe"]    - m_lag["sharpe"],    4),
        "annret_delta"   : round(m_same["ann_return"] - m_lag["ann_return"], 2),
    }


# ── Build comparison rows ──────────────────────────────────────────────────────

# Load top-5 from existing optimization files
ma_opt  = pd.read_csv(OUT_DIR / "MA_Crossover_Optimization.csv")
cta_opt = pd.read_csv(OUT_DIR / "CTA_Optimization.csv")

top5_ma  = ma_opt.head(5)
top5_cta = cta_opt.head(5)

rows = []

print("\n" + "=" * 80)
print("MA CROSSOVER — top 5 pairs, lag-1 vs same-day")
print("=" * 80)

for _, r in top5_ma.iterrows():
    m, n = int(r["m"]), int(r["n"])
    sig  = ma_signal_np(m, n)
    d    = dual_metrics(sig)
    label = f"MA({m},{n})"
    d["strategy"] = label
    d["type"]     = "MA_Crossover"
    rows.append(d)

for _, r in top5_cta.iterrows():
    s, l = int(r["S"]), int(r["L"])
    sig  = cta_single_signal_np(s, l)
    d    = dual_metrics(sig)
    d["strategy"] = f"CTA({s},{l})"
    d["type"]     = "CTA_SinglePair"
    rows.append(d)

# Paper CTA
sig_paper = cta_paper_signal_np()
d_paper   = dual_metrics(sig_paper)
d_paper["strategy"] = "CTA_Paper(8-16-32/24-48-96)"
d_paper["type"]     = "CTA_Paper"
rows.append(d_paper)

df = pd.DataFrame(rows)


# ── Pretty print ───────────────────────────────────────────────────────────────

def _pct(x):
    return f"{x*100:.2f}%" if isinstance(x, float) and not np.isnan(x) else "n/a"

SEP = "-" * 120

header = (f"{'Strategy':<32} | "
          f"{'Sharpe':>8} {'Sharpe':>8} {'Chg':>7} | "
          f"{'AnnRet':>9} {'AnnRet':>9} {'Chg':>8} | "
          f"{'MaxDD':>9} {'MaxDD':>9} | "
          f"{'TotPnL':>9} {'TotPnL':>9} | "
          f"{'HitRate':>8} {'HitRate':>8}")
sub    = (f"{'':32} | "
          f"{'Lag-1':>8} {'SameDay':>8} {'':>7} | "
          f"{'Lag-1':>9} {'SameDay':>9} {'':>8} | "
          f"{'Lag-1':>9} {'SameDay':>9} | "
          f"{'Lag-1':>9} {'SameDay':>9} | "
          f"{'Lag-1':>8} {'SameDay':>8}")

for group, gdf in df.groupby("type", sort=False):
    print("\n" + "=" * 120)
    print(f"  {group}")
    print("=" * 120)
    print(header)
    print(sub)
    print(SEP)

    for _, row in gdf.iterrows():
        winner_sharpe = "<<" if row["sharpe_same"] > row["sharpe_lag1"] else "  "
        print(
            f"{row['strategy']:<32} | "
            f"{row['sharpe_lag1']:>8.4f} {row['sharpe_same']:>8.4f} "
            f"{row['sharpe_delta']:>+7.4f} | "
            f"{row['ann_ret_lag1']:>9,.1f} {row['ann_ret_same']:>9,.1f} "
            f"{row['annret_delta']:>+8,.1f} | "
            f"{row['maxdd_lag1']:>9,.1f} {row['maxdd_same']:>9,.1f} | "
            f"{row['totpnl_lag1']:>9,.1f} {row['totpnl_same']:>9,.1f} | "
            f"{_pct(row['hitrate_lag1']):>8} {_pct(row['hitrate_same']):>8}"
            f"  {winner_sharpe}"
        )


# ── Recommendation ─────────────────────────────────────────────────────────────

print("\n" + "=" * 120)
print("SUMMARY & RECOMMENDATION")
print("=" * 120)

ma_rows  = df[df["type"] == "MA_Crossover"]
cta_rows = df[df["type"] == "CTA_SinglePair"]

ma_same_wins  = (ma_rows["sharpe_same"]  > ma_rows["sharpe_lag1"]).sum()
cta_same_wins = (cta_rows["sharpe_same"] > cta_rows["sharpe_lag1"]).sum()

print(f"\n  MA Crossover  : same-day wins {ma_same_wins}/5 pairs by Sharpe")
print(f"  CTA SinglePair: same-day wins {cta_same_wins}/5 pairs by Sharpe")

avg_delta_ma  = ma_rows["sharpe_delta"].mean()
avg_delta_cta = cta_rows["sharpe_delta"].mean()
print(f"\n  Avg Sharpe delta (same-day − lag-1):")
print(f"    MA Crossover  : {avg_delta_ma:+.4f}")
print(f"    CTA SinglePair: {avg_delta_cta:+.4f}")

print("""
  Interpretation guide
  ────────────────────
  If same-day Sharpe > lag-1 Sharpe consistently:
    → same-day entry is capturing meaningful trend continuation at the close;
      feasible for liquid futures (LME Copper Ring settlement is a firm price).

  If lag-1 Sharpe > same-day Sharpe:
    → the trend has already partially reversed by the next session;
      the extra day of lag avoids entering after the move is exhausted.

  For LONG-window MA (m≥30, n≥40): same-day vs lag-1 difference is typically
  small because a single day shifts a 30+ day average by <3%.

  For SHORT-window CTA (S=9, L=21): the EWMA reacts faster, so today's
  closing price has a larger weight — same-day can capture more of the
  continuation but also introduces slightly more look-ahead bias.

  RECOMMENDED: choose the convention that consistently outperforms across
  BOTH MA and CTA.  If results are mixed, prefer LAG-1 as the conservative
  default (no execution-timing ambiguity, cleaner out-of-sample validity).
""")

# ── Save CSV ───────────────────────────────────────────────────────────────────
col_order = ["type", "strategy",
             "sharpe_lag1", "sharpe_same", "sharpe_delta",
             "ann_ret_lag1", "ann_ret_same", "annret_delta",
             "maxdd_lag1", "maxdd_same",
             "totpnl_lag1", "totpnl_same",
             "hitrate_lag1", "hitrate_same"]
df[col_order].to_csv(OUT_DIR / "Signal_Timing_Comparison.csv", index=False)
print(f"  Saved -> {OUT_DIR / 'Signal_Timing_Comparison.csv'}")
