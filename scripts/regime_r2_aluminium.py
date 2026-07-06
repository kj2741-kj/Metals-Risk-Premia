"""
regime_r2.py -- R2 Volatility Regime Portfolio (Copper)

Regime = rolling 21-day realised vol of F1_continuous daily returns, shift-1
(no look-ahead). Unlike R1, the split point (high vol vs low vol) is not a
free absolute number picked by eye -- it is defined as the MEDIAN of the
21d-vol series over each walk-forward window's own IS period, then applied
to classify both IS (for weight-fitting) and OOS days. This keeps the same
no-look-ahead, no-hand-picked-constant discipline as the R1 tests.

Legs = current Copper production config:
  Momentum: MA(20,45)            Carry: CarryMom-20d           Value: V1 F8, 5yr

Weights are IS-fit per walk-forward window (QP max-Sharpe, >=0, sum=1),
separately for high-vol and low-vol days within each 5yr IS window, then
applied OOS -- identical methodology to regime_r1_v2.py, for a clean
mechanism-vs-mechanism comparison (term-structure regime vs vol regime).
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

f1_df = pd.read_csv(f"{BASE}\\LME_Aluminium_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()

xls   = pd.ExcelFile(f"{BASE}\\Metals Futures Curve.csv")
cu_sh = "ALuminium LME"
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

# ── Legs (Copper production config: MA(20,45) / CarryMom-20d / V1 F8) ─────────
mom_pos = np.sign(f1r.rolling(60).mean() - f1r.rolling(115).mean()).shift(1).fillna(0)  # Aluminium leg

roll_yield = ((crv["F1"] - crv["F2"]) / crv["F1"]).replace([np.inf, -np.inf], np.nan)
zsc = ((roll_yield - roll_yield.rolling(252).mean()) / roll_yield.rolling(252).std()).replace([np.inf, -np.inf], np.nan)
carry_pos  = np.sign(zsc).shift(1).reindex(f1c.index).fillna(0)  # Aluminium production carry leg (Z-score)

def v1_pos(k, N=1260, thr=0.10):
    p   = crv[f"F{k}"].dropna()
    ma  = p.rolling(N, min_periods=N // 2).mean()
    dev = ((p - ma) / ma).replace([np.inf, -np.inf], np.nan).dropna()
    sig = np.where(dev < -thr, 1.0, np.where(dev > thr, -1.0, 0.0))
    return pd.Series(sig, index=dev.index).shift(1).fillna(0)

val_pos = v1_pos(12)  # Aluminium value contract

# ── R2 regime: rolling 21d realised vol of F1_continuous returns, shift-1 ─────
daily_ret = f1c.pct_change()
vol21 = daily_ret.rolling(21).std().shift(1)

# ── Align ───────────────────────────────────────────────────────────────────────
idx = f1c.index
for s in [mom_pos, carry_pos, val_pos, vol21]:
    idx = idx.intersection(s.dropna().index if s.name != vol21.name else s.index)
idx = f1c.index.intersection(mom_pos.index).intersection(carry_pos.index).intersection(val_pos.index).intersection(vol21.dropna().index)
idx = idx.sort_values()

mom_a   = mom_pos.reindex(idx).fillna(0)
carry_a = carry_pos.reindex(idx).fillna(0)
val_a   = val_pos.reindex(idx).fillna(0)
vol_a   = vol21.reindex(idx)
f1c_a   = f1c.reindex(idx)
f1r_a   = f1r.reindex(idx)
T       = len(idx)

print(f"Common sample: {len(idx)} days | {idx[0].date()} -> {idx[-1].date()}\n")

def ret_from_pos(pos):
    with np.errstate(invalid="ignore", divide="ignore"):
        return (pos * f1c_a.diff() / f1c_a.shift(1)).replace([np.inf, -np.inf], np.nan)

mom_r, carry_r, val_r = ret_from_pos(mom_a), ret_from_pos(carry_a), ret_from_pos(val_a)
leg_ret = np.column_stack([mom_r.fillna(0).values, carry_r.fillna(0).values, val_r.fillna(0).values])
vol_np  = vol_a.values

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

print("=" * 118)
print("R2 -- IS-OPTIMISED VOL-REGIME WEIGHTS (Mom MA(60,115) / Carry Z-score-252d / Value V1 F12 -- ALUMINIUM), walk-forward, net 5bps")
print("Regime split = IS-period median of rolling-21d realised vol (no hand-picked threshold)")
print("=" * 118)
hdr = (f"{'OOS':7s} {'%HighVol':>8s} | {'w_Hi (M/C/V)':>16s} | {'w_Lo (M/C/V)':>16s} | {'EW':>7s} {'R2':>7s} {'Diff':>7s}")
print(hdr)
print("-" * 90)

rows = []
oos_s = IS_W
port_ew_full = (mom_a + carry_a + val_a) / 3.0

while oos_s < T:
    oos_e = min(oos_s + OOS_W, T)
    oos_yr = str(idx[oos_s].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
    is_slice = slice(oos_s - IS_W, oos_s)
    vol_is = vol_np[is_slice]
    thresh = np.nanmedian(vol_is)
    hi_is  = vol_is > thresh
    ret_is = leg_ret[is_slice]

    w_hi = opt_w(ret_is[hi_is])  if hi_is.sum()  >= 20 else np.array([1/3,1/3,1/3])
    w_lo = opt_w(ret_is[~hi_is]) if (~hi_is).sum() >= 20 else np.array([1/3,1/3,1/3])

    oos_dt = idx[oos_s:oos_e]
    hi_oos = (vol_a.reindex(oos_dt).values > thresh)
    w_oos  = np.where(hi_oos[:, None], w_hi, w_lo)
    m_oos  = mom_a.reindex(oos_dt).values
    c_oos  = carry_a.reindex(oos_dt).values
    v_oos  = val_a.reindex(oos_dt).values
    port_r2 = pd.Series(w_oos[:,0]*m_oos + w_oos[:,1]*c_oos + w_oos[:,2]*v_oos, index=oos_dt)
    port_ew = port_ew_full.reindex(oos_dt)

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

    sh_ew = win_sharpe_net(port_ew)
    sh_r2 = win_sharpe_net(port_r2)
    hi_pct = 100 * hi_oos.mean()

    print(f"{oos_yr:7s} {hi_pct:7.1f}% | {w_hi[0]:.2f}/{w_hi[1]:.2f}/{w_hi[2]:.2f}    | "
          f"{w_lo[0]:.2f}/{w_lo[1]:.2f}/{w_lo[2]:.2f}    | {sh_ew:+7.3f} {sh_r2:+7.3f} {sh_r2-sh_ew:+7.3f}")

    rows.append({"OOS": oos_yr, "EW": sh_ew, "R2": sh_r2})
    oos_s += OOS_W

df = pd.DataFrame(rows)
print("-" * 90)
avg_ew, avg_r2 = np.nanmean(df["EW"]), np.nanmean(df["R2"])
print(f"AVERAGE OOS SHARPE:  EW={avg_ew:+.3f}   R2={avg_r2:+.3f}   Diff={avg_r2-avg_ew:+.3f}")
post22 = df[df["OOS"].str.replace("*","").astype(int) >= 2022]
print(f"POST-2022 AVERAGE:   EW={np.nanmean(post22['EW']):+.3f}   R2={np.nanmean(post22['R2']):+.3f}   "
      f"Diff={np.nanmean(post22['R2'])-np.nanmean(post22['EW']):+.3f}")
