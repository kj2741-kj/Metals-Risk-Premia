"""
weight_sensitivity.py
Three analyses on fixed anchors MA(10,25) + MA(35,43) + MA(63,100):

  1. IS-optimised weights (max-Sharpe QP each 5yr IS window → apply OOS 1yr)
  2. Weight sensitivity grid (all w1+w2+w3=1, step 0.1 → avg OOS Sharpe surface)
  3. Rebalancing intervals (how often to re-optimise: 1yr / 3yr / 5yr)

Baseline: MA(35,43) fixed = +0.492 avg OOS
"""
import warnings, sys, io, os
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import numpy as np
import pandas as pd
from scipy.optimize import minimize

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
DATA_DIR = os.path.join(_REPO_ROOT, "data")
OUTPUTS_DIR = os.path.join(_REPO_ROOT, "outputs")
BASE = DATA_DIR

f1_df = pd.read_csv(f"{BASE}\\LME_Copper_Rolling_F1_v2.csv",
                    parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()
print(f"F1: {len(f1r)} rows | {f1r.index[0].date()} → {f1r.index[-1].date()}")

IS, OOS = 1260, 252
ANCHORS = [(10, 25), (35, 43), (63, 100)]   # fast, medium, slow
NAMES   = ["MA(10,25)", "MA(35,43)", "MA(63,100)"]

def L1(sig): return sig.shift(1).fillna(0)

def sharpe_from_ret(ret_s):
    act = ret_s[ret_s != 0].dropna()
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def sharpe_pos(pos, f1c_s):
    idx = pos.index.intersection(f1c_s.index)
    p = pos.reindex(idx); c = f1c_s.reindex(idx)
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (p * c.diff() / c.shift(1)).replace([np.inf,-np.inf], np.nan)
    return sharpe_from_ret(ret)

def F(x): return f"{x:+.3f}" if (x is not None and not np.isnan(x)) else "  —"

# ── Pre-build anchor signal series (full dataset) ──────────────────────────────
print("Building anchor position series...", flush=True)
anchor_pos = []
for am, an in ANCHORS:
    anchor_pos.append(L1(np.sign(f1r.rolling(am).mean() - f1r.rolling(an).mean())))

# MA(35,43) fixed
ma35_pos = anchor_pos[1]

# Pre-compute daily returns for each anchor
T = len(f1r); dates = f1r.index
delta  = f1c.diff()
prev   = f1c.shift(1)

anchor_ret = []
for pos in anchor_pos:
    with np.errstate(invalid="ignore", divide="ignore"):
        r = (pos * delta / prev).replace([np.inf,-np.inf], np.nan).reindex(dates)
    anchor_ret.append(r)

# Stack into (T, 3) return matrix
ret_mat = pd.concat(anchor_ret, axis=1)
ret_mat.columns = NAMES

# ── Helper: max-Sharpe weights for a return matrix window ─────────────────────
def max_sharpe_weights(ret_win):
    """IS optimisation: find w*=(w1,w2,w3) summing to 1 maximising Sharpe."""
    def neg_sh(w):
        combined = ret_win.values @ w
        act = combined[combined != 0]
        if len(act) < 20: return 0.0
        sd = act.std(ddof=1)
        return -(act.mean() / sd * np.sqrt(252)) if sd > 0 else 0.0

    best_val, best_w = np.inf, np.array([1/3, 1/3, 1/3])
    # Try multiple starting points to avoid local optima
    for w0 in [[1/3,1/3,1/3],[0.8,0.1,0.1],[0.1,0.8,0.1],[0.1,0.1,0.8],
               [0.5,0.5,0],[0.5,0,0.5],[0,0.5,0.5]]:
        try:
            res = minimize(neg_sh, w0,
                           method="SLSQP",
                           bounds=[(0,1)]*3,
                           constraints=[{"type":"eq","fun":lambda w: w.sum()-1}],
                           options={"ftol":1e-9,"maxiter":500})
            if res.fun < best_val:
                best_val, best_w = res.fun, res.x
        except: pass
    return best_w / best_w.sum()   # re-normalise for numerical safety

# ── Walk-forward OOS windows (shared for all analyses) ────────────────────────
oos_windows = []
oos_s = IS
while oos_s < T:
    oos_e     = min(oos_s + OOS, T)
    oos_dates = dates[oos_s:oos_e]
    oos_windows.append((oos_s, oos_e, oos_dates))
    oos_s += OOS

print(f"{len(oos_windows)} OOS windows  ({dates[IS].year}–{dates[min(IS+len(oos_windows)*OOS-1,T-1)].year})")

# ═══════════════════════════════════════════════════════════════════════════════
# ANALYSIS 1 — IS-OPTIMISED WEIGHTS (annual rebalance)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n--- Analysis 1: IS-optimised weights (annual rebalance) ---", flush=True)

rows1 = []
for oos_s, oos_e, oos_dates in oos_windows:
    is_s    = oos_s - IS
    oos_yr  = str(dates[oos_s].year) + ("*" if (oos_e-oos_s) < OOS else "")
    is_lbl  = f"{dates[is_s].year}–{dates[oos_s-1].year}"

    is_rets = ret_mat.iloc[is_s:oos_s]
    w_opt   = max_sharpe_weights(is_rets)

    oos_rets = ret_mat.reindex(oos_dates)
    combined_oos = oos_rets.values @ w_opt
    combined_ser = pd.Series(combined_oos, index=oos_dates)

    ma35_sh = sharpe_from_ret(ret_mat["MA(35,43)"].reindex(oos_dates))
    eq_sh   = sharpe_from_ret(oos_rets.mean(axis=1))
    opt_sh  = sharpe_from_ret(combined_ser)

    rows1.append({
        "OOS": oos_yr, "IS window": is_lbl,
        "MA(35,43)": F(ma35_sh),
        "EW anchors": F(eq_sh),
        "IS-opt weights": F(opt_sh),
        "w_fast": f"{w_opt[0]:.2f}",
        "w_med":  f"{w_opt[1]:.2f}",
        "w_slow": f"{w_opt[2]:.2f}",
    })
    print(f"  {oos_yr} ({is_lbl}): MA35={F(ma35_sh)}  EW={F(eq_sh)}  "
          f"OPT={F(opt_sh)}  w=({w_opt[0]:.2f},{w_opt[1]:.2f},{w_opt[2]:.2f})", flush=True)

df1 = pd.DataFrame(rows1)

# ═══════════════════════════════════════════════════════════════════════════════
# ANALYSIS 2 — WEIGHT SENSITIVITY GRID
# ═══════════════════════════════════════════════════════════════════════════════
print("\n--- Analysis 2: Weight sensitivity grid (step=0.1) ---", flush=True)

# Pre-compute OOS return series per anchor per window
oos_ret_blocks = []   # list of (oos_dates, ret_df) per window
for oos_s, oos_e, oos_dates in oos_windows:
    oos_ret_blocks.append(ret_mat.reindex(oos_dates))

# Grid search
STEP = 0.1
grid_rows = []
best_sh, best_w = -np.inf, (1/3, 1/3, 1/3)
w_vals = np.round(np.arange(0, 1+STEP, STEP), 2)

for wf in w_vals:
    for wm in w_vals:
        ws = round(1 - wf - wm, 2)
        if ws < -1e-9: continue
        ws = max(ws, 0.0)
        w = np.array([wf, wm, ws])
        # Avg OOS Sharpe across all windows (NOTE: post-hoc — uses OOS data to find best w)
        sharpes = []
        for block in oos_ret_blocks:
            combined = block.values @ w
            ser = pd.Series(combined, index=block.index)
            sh  = sharpe_from_ret(ser)
            if not np.isnan(sh): sharpes.append(sh)
        avg_sh = np.mean(sharpes) if sharpes else np.nan
        grid_rows.append({"w_fast": wf, "w_med": wm, "w_slow": ws, "avg_OOS_Sharpe": avg_sh})
        if not np.isnan(avg_sh) and avg_sh > best_sh:
            best_sh, best_w = avg_sh, (wf, wm, ws)

grid_df = pd.DataFrame(grid_rows)
print(f"  Grid size: {len(grid_df)} combinations")
print(f"  Best weight combo (POST-HOC — NOT valid OOS): "
      f"w=({best_w[0]:.1f},{best_w[1]:.1f},{best_w[2]:.1f}) → avg Sharpe={best_sh:+.3f}")

# ═══════════════════════════════════════════════════════════════════════════════
# ANALYSIS 3 — REBALANCING INTERVALS (1yr / 3yr / 5yr)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n--- Analysis 3: Rebalancing intervals ---", flush=True)

def run_rebalance(rebal_every_n_windows):
    """Re-optimise weights every N OOS windows; hold fixed in between."""
    sharpes = []
    current_w = np.array([1/3, 1/3, 1/3])
    for i, (oos_s, oos_e, oos_dates) in enumerate(oos_windows):
        if i % rebal_every_n_windows == 0:
            is_s    = oos_s - IS
            is_rets = ret_mat.iloc[is_s:oos_s]
            current_w = max_sharpe_weights(is_rets)
        oos_rets  = ret_mat.reindex(oos_dates)
        combined  = pd.Series(oos_rets.values @ current_w, index=oos_dates)
        sh = sharpe_from_ret(combined)
        if not np.isnan(sh): sharpes.append(sh)
    return np.mean(sharpes) if sharpes else np.nan

rebal_results = {}
for n, label in [(1, "Annual (1yr)"), (3, "Every 3yr"), (5, "Every 5yr")]:
    sh = run_rebalance(n)
    rebal_results[label] = sh
    print(f"  {label}: avg OOS Sharpe = {sh:+.3f}", flush=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PRINT RESULTS
# ═══════════════════════════════════════════════════════════════════════════════
pd.set_option("display.width", 200)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

def avg_col(df, col):
    vals = []
    for v in df[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print("\n" + "="*150)
print("ANALYSIS 1 — IS-OPTIMISED WEIGHTS (annual rebalance, valid walk-forward)")
print("="*150)
print(df1.to_string(index=False))
print(f"\nAverage OOS Sharpe:")
print(f"  MA(35,43) fixed    : {avg_col(df1, 'MA(35,43)'):+.3f}")
print(f"  Equal-weight EW    : {avg_col(df1, 'EW anchors'):+.3f}")
print(f"  IS-opt weights     : {avg_col(df1, 'IS-opt weights'):+.3f}")

print("\n" + "="*100)
print("ANALYSIS 2 — WEIGHT SENSITIVITY GRID  (⚠ post-hoc: OOS data used to find best w)")
print("Rows = w_fast | Cols = w_med | cell = avg OOS Sharpe | w_slow = 1 - w_fast - w_med")
print("="*100)

# Pivot table: rows=w_fast, cols=w_med
pivot = grid_df.pivot_table(index="w_fast", columns="w_med",
                             values="avg_OOS_Sharpe", aggfunc="mean")
pivot.index   = [f"wF={v:.1f}" for v in pivot.index]
pivot.columns = [f"wM={v:.1f}" for v in pivot.columns]
print(pivot.round(3).to_string())
print(f"\n  Peak: w=({best_w[0]:.1f} fast, {best_w[1]:.1f} med, {best_w[2]:.1f} slow) "
      f"→ avg OOS Sharpe = {best_sh:+.3f}  [⚠ post-hoc]")
print(f"  MA(35,43) fixed = +0.492  (no grid search needed)")

# Show top-10 weight combos
top10 = grid_df.nlargest(10, "avg_OOS_Sharpe")[["w_fast","w_med","w_slow","avg_OOS_Sharpe"]]
print(f"\n  Top-10 weight combos (post-hoc):")
print(top10.to_string(index=False))

print("\n" + "="*60)
print("ANALYSIS 3 — REBALANCING INTERVALS")
print("="*60)
print(f"  MA(35,43) fixed (no rebalance) : +0.492")
print(f"  Equal-weight fixed             : -0.012")
for label, sh in rebal_results.items():
    print(f"  IS-opt weights, {label:15s}: {sh:+.3f}")

print("\n" + "="*60)
print("FULL SUMMARY")
print("="*60)
print(f"  MA(35,43) fixed                : +0.492  ← baseline")
print(f"  Equal-weight anchors           : -0.012")
print(f"  IS-opt weights (annual)        : {avg_col(df1, 'IS-opt weights'):+.3f}")
for label, sh in rebal_results.items():
    print(f"  IS-opt {label:22s}: {sh:+.3f}")
print(f"  Best fixed weight (post-hoc)   : {best_sh:+.3f}  w=({best_w[0]:.1f},{best_w[1]:.1f},{best_w[2]:.1f})")
