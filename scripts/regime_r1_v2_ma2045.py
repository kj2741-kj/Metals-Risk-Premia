"""
regime_r1_v2.py -- R1 Term-Structure Regime Portfolio, v2 (Copper)

Changes from v1 (regime_r1.py):
  1. Carry leg swapped from CarryMom-20d to V3 Z-score-252d for regime testing.
     Reason: the regime classifier is sign(roll_yield); CarryMom-20d is the
     20-day CHANGE in that same roll_yield, so it shares a data-generating
     variable with the classifier. Z-score-252d is a standardised LEVEL
     measure -- a cleaner "carry" leg that isn't a derivative of the same
     thing defining the regime. (CarryMom-20d stays the production signal
     on the standalone Carry tab / EW portfolio -- that's a separate,
     already-settled question; this swap is regime-research-only.)
  2. Regime weights are IS-fit per walk-forward window (QP max-Sharpe,
     >=0, sum=1), separately for backwardation days and contango days
     within each 5yr IS window, then applied OOS -- no full-sample fit,
     no hand-picked weight presets.
  3. Diagnostic breakdown per window: regime composition (%BD), and each
     leg's conditional Sharpe within backwardation vs contango, to see
     WHY EW vs R1 diverge post-COVID / post-2022.
"""
import os, sys, io, warnings
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
IS_W, OOS_W = 1260, 252
TC_BPS = 5

# ── Load F1 + curve (same loaders as regime_r1.py) ─────────────────────────────
f1_df = pd.read_csv(f"{BASE}\\LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()

xls   = pd.ExcelFile(f"{BASE}\\Metals Futures Curve.csv")
cu_sh = next((s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()),
             next((s for s in xls.sheet_names if "copper" in s.lower()), xls.sheet_names[0]))
df_raw = pd.read_excel(xls, sheet_name=cu_sh, header=None, nrows=5)
hr = [i for i in range(min(4, len(df_raw)))
      if any(kw in " ".join(str(v).lower() for v in df_raw.iloc[i].values if pd.notna(v))
             for kw in ["date", "f1", "f2", "price"])]
df_c = pd.read_excel(xls, sheet_name=cu_sh, header=hr if hr else [0, 1])
if isinstance(df_c.columns, pd.MultiIndex):
    df_c.columns = ["_".join(str(p) for p in t if pd.notna(p) and "Unnamed" not in str(p) and str(p).strip())
                    or str(t) for t in df_c.columns]
df_c.columns = [str(c).strip() for c in df_c.columns]
dc = [c for c in df_c.columns if "date" in c.lower()]
if dc: df_c = df_c.rename(columns={dc[0]: "Date"})
df_c["Date"] = pd.to_datetime(df_c["Date"], errors="coerce")
df_c = df_c.dropna(subset=["Date"]).set_index("Date").sort_index()
df_c.index = df_c.index.normalize()

curve_prices = {}
for col in df_c.columns:
    cl = col.lower().replace(" ", "_")
    for i in range(1, 27):
        if f"f{i}_" in cl and "price" in cl:
            curve_prices[f"F{i}"] = pd.to_numeric(df_c[col], errors="coerce")
            break
crv = pd.DataFrame(curve_prices, index=df_c.index)

# ── Legs ────────────────────────────────────────────────────────────────────────
mom_pos = np.sign(f1r.rolling(20).mean() - f1r.rolling(45).mean()).shift(1).fillna(0)  # updated to new Cu default MA(20,45)

roll_yield = ((crv["F1"] - crv["F2"]) / crv["F1"]).replace([np.inf, -np.inf], np.nan)

# Carry leg for regime research = V3 Z-score-252d (NOT CarryMom-20d, see docstring)
zsc = ((roll_yield - roll_yield.rolling(252).mean()) / roll_yield.rolling(252).std()
       ).replace([np.inf, -np.inf], np.nan)
carry_pos = np.sign(zsc).shift(1).reindex(f1c.index).fillna(0)

def v1_pos(k, N=1260, thr=0.10):
    p   = crv[f"F{k}"].dropna()
    ma  = p.rolling(N, min_periods=N // 2).mean()
    dev = ((p - ma) / ma).replace([np.inf, -np.inf], np.nan).dropna()
    sig = np.where(dev < -thr, 1.0, np.where(dev > thr, -1.0, 0.0))
    return pd.Series(sig, index=dev.index).shift(1).fillna(0)

val_pos = v1_pos(8)

regime_bd = (roll_yield.shift(1) > 0).reindex(f1c.index).fillna(False)

# ── Align ───────────────────────────────────────────────────────────────────────
idx = f1c.index
for s in [mom_pos, carry_pos, val_pos, regime_bd]:
    idx = idx.intersection(s.index)
idx = idx.sort_values()

mom_a   = mom_pos.reindex(idx).fillna(0)
carry_a = carry_pos.reindex(idx).fillna(0)
val_a   = val_pos.reindex(idx).fillna(0)
bd_a    = regime_bd.reindex(idx).fillna(False)
f1c_a   = f1c.reindex(idx)
f1r_a   = f1r.reindex(idx)
T       = len(idx)

def ret_from_pos(pos):
    with np.errstate(invalid="ignore", divide="ignore"):
        return (pos * f1c_a.diff() / f1c_a.shift(1)).replace([np.inf, -np.inf], np.nan)

mom_r, carry_r, val_r = ret_from_pos(mom_a), ret_from_pos(carry_a), ret_from_pos(val_a)
leg_ret = np.column_stack([mom_r.fillna(0).values, carry_r.fillna(0).values, val_r.fillna(0).values])

def sharpe_of(ret_arr, pos_arr):
    act = ret_arr[pos_arr != 0]
    act = act[~np.isnan(act)]
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def opt_w(ret_mat_sub):
    def neg_sh(w):
        r = ret_mat_sub @ w
        return -sharpe_of(r, r) if not np.isnan(sharpe_of(r, r)) else 0.0
    best_v, best_w = np.inf, np.array([1/3, 1/3, 1/3])
    for w0 in [[1/3,1/3,1/3],[1,0,0],[0,1,0],[0,0,1],[0.5,0.5,0],[0.5,0,0.5],[0,0.5,0.5]]:
        try:
            res = minimize(neg_sh, w0, method="SLSQP", bounds=[(0,1)]*3,
                           constraints=[{"type":"eq","fun":lambda w: w.sum()-1}],
                           options={"ftol":1e-9,"maxiter":300})
            if res.fun < best_v: best_v, best_w = res.fun, res.x
        except Exception:
            pass
    w = np.clip(best_w, 0, 1); return w / w.sum()

print("=" * 112)
print("R1 v2 -- IS-OPTIMISED REGIME WEIGHTS (Mom / Carry-Zscore / Value), walk-forward, net 5bps")
print("=" * 112)
hdr = (f"{'OOS':7s} {'%BD':>6s} | {'w_BD (M/C/V)':>16s} | {'w_CTG (M/C/V)':>16s} | "
       f"{'EW':>7s} {'R1v2':>7s} {'Diff':>7s} | "
       f"{'Mom@BD':>7s} {'Mom@CTG':>8s} {'Carry@BD':>9s} {'Carry@CTG':>10s} {'Val@BD':>7s} {'Val@CTG':>8s}")
print(hdr)
print("-" * 150)

rows = []
oos_s = IS_W
port_ew_full = (mom_a + carry_a + val_a) / 3.0

while oos_s < T:
    oos_e = min(oos_s + OOS_W, T)
    oos_yr = str(idx[oos_s].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
    is_slice = slice(oos_s - IS_W, oos_s)
    bd_is = bd_a.iloc[is_slice].values
    ret_is = leg_ret[is_slice]

    w_bd  = opt_w(ret_is[bd_is])  if bd_is.sum()  >= 20 else np.array([1/3,1/3,1/3])
    w_ctg = opt_w(ret_is[~bd_is]) if (~bd_is).sum() >= 20 else np.array([1/3,1/3,1/3])

    # Conditional per-leg IS Sharpes (diagnostic)
    def cond_sh(leg_idx, mask):
        r = ret_is[:, leg_idx][mask]
        return sharpe_of(r, r)
    diag = [cond_sh(0, bd_is), cond_sh(0, ~bd_is), cond_sh(1, bd_is), cond_sh(1, ~bd_is),
            cond_sh(2, bd_is), cond_sh(2, ~bd_is)]

    oos_dt = idx[oos_s:oos_e]
    bd_oos = bd_a.reindex(oos_dt).values
    w_oos  = np.where(bd_oos[:, None], w_bd, w_ctg)
    m_oos  = mom_a.reindex(oos_dt).values
    c_oos  = carry_a.reindex(oos_dt).values
    v_oos  = val_a.reindex(oos_dt).values
    port_r1v2 = pd.Series(w_oos[:,0]*m_oos + w_oos[:,1]*c_oos + w_oos[:,2]*v_oos, index=oos_dt)
    port_ew   = port_ew_full.reindex(oos_dt)

    def win_sharpe_net(pos_slice):
        c = f1c_a.reindex(oos_dt); r = f1r_a.reindex(oos_dt)
        pnl = pos_slice.values * np.diff(c.values, prepend=np.nan)
        chg = np.abs(np.diff(pos_slice.values, prepend=pos_slice.values[0]))
        pnl = pnl - chg * (TC_BPS / 10000.0 / 2.0) * r.values
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = pnl / np.concatenate([[np.nan], c.values[:-1]])
        act_mask = (pos_slice.values != 0) & ~np.isnan(ret)
        act = ret[act_mask]
        if len(act) < 20: return np.nan
        sd = act.std(ddof=1)
        return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

    sh_ew   = win_sharpe_net(port_ew)
    sh_r1v2 = win_sharpe_net(port_r1v2)
    bd_pct  = 100 * bd_oos.mean()

    def F(x): return f"{x:+.2f}" if not np.isnan(x) else "  - "

    print(f"{oos_yr:7s} {bd_pct:5.1f}% | "
          f"{w_bd[0]:.2f}/{w_bd[1]:.2f}/{w_bd[2]:.2f}    | "
          f"{w_ctg[0]:.2f}/{w_ctg[1]:.2f}/{w_ctg[2]:.2f}    | "
          f"{sh_ew:+7.3f} {sh_r1v2:+7.3f} {sh_r1v2-sh_ew:+7.3f} | "
          f"{F(diag[0]):>7s} {F(diag[1]):>8s} {F(diag[2]):>9s} {F(diag[3]):>10s} {F(diag[4]):>7s} {F(diag[5]):>8s}")

    rows.append({"OOS": oos_yr, "EW": sh_ew, "R1v2": sh_r1v2, "BDpct": bd_pct,
                 "wM_bd": w_bd[0], "wC_bd": w_bd[1], "wV_bd": w_bd[2],
                 "wM_ctg": w_ctg[0], "wC_ctg": w_ctg[1], "wV_ctg": w_ctg[2]})
    oos_s += OOS_W

df = pd.DataFrame(rows)
print("-" * 150)
avg_ew, avg_r1 = np.nanmean(df["EW"]), np.nanmean(df["R1v2"])
print(f"AVERAGE OOS SHARPE:  EW={avg_ew:+.3f}   R1v2={avg_r1:+.3f}   Diff={avg_r1-avg_ew:+.3f}")
post22 = df[df["OOS"].str.replace("*","").astype(int) >= 2022]
print(f"POST-2022 AVERAGE:   EW={np.nanmean(post22['EW']):+.3f}   R1v2={np.nanmean(post22['R1v2']):+.3f}   "
      f"Diff={np.nanmean(post22['R1v2'])-np.nanmean(post22['EW']):+.3f}")
print()
print("Average IS-fit weights across all windows:")
print(f"  Backwardation:  Mom={df['wM_bd'].mean():.3f}  Carry(Z)={df['wC_bd'].mean():.3f}  Val={df['wV_bd'].mean():.3f}")
print(f"  Contango:       Mom={df['wM_ctg'].mean():.3f}  Carry(Z)={df['wC_ctg'].mean():.3f}  Val={df['wV_ctg'].mean():.3f}")
