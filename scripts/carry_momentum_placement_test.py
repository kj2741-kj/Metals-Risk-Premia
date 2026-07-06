"""
carry_momentum_placement_test.py -- where does Carry Momentum (CarryMom-20d) add value?

Background: Boons & Prado (2019, Journal of Finance), "Basis-Momentum," identify momentum
in the near-term futures basis as a return predictor DISTINCT from both the basis level
(carry) and price momentum -- priced, maturity-specific, driven by roll returns, and
INCREASING IN VOLATILITY. That matches this project's own finding that Copper's Carry
Momentum sleeve was strongest "during repeated inventory squeezes and backwardation
episodes in 2021-2024" -- i.e. its edge looks regime-concentrated, not uniform.

Base 3-sleeve EW (current production): Momentum MA(20,45) + Carry Z-score-252d + Value V1 F8.
This script tests, walk-forward, net of 5bps, whether Carry Momentum adds value:
  (A) as an unconditional 4th equal-weight sleeve (naive 1/4 each)
  (B) as an overlay active ONLY during backwardation (term-structure regime)
  (C) as an overlay active ONLY during high realised volatility (vol regime, per Boons-Prado)

All regime splits are FIXED, simple rules (median-of-IS threshold or sign of roll_yield) --
no per-window weight optimisation -- to keep this a quick, non-overfit-prone first look
before deciding whether it's worth a fuller build.
"""
import os, sys, io, warnings
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import numpy as np
import pandas as pd

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
DATA_DIR = os.path.join(_REPO_ROOT, "data")
OUTPUTS_DIR = os.path.join(_REPO_ROOT, "outputs")
BASE = DATA_DIR
IS_W, OOS_W = 1260, 252
TC_BPS = 5

f1_df = pd.read_csv(f"{BASE}\\LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()

xls   = pd.ExcelFile(f"{BASE}\\Metals Futures Curve.csv")
cu_sh = next((s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()), xls.sheet_names[0])
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

# ── Legs (Copper production config) ────────────────────────────────────────────
mom_pos = np.sign(f1r.rolling(20).mean() - f1r.rolling(45).mean()).shift(1).fillna(0)

roll_yield = ((crv["F1"] - crv["F2"]) / crv["F1"]).replace([np.inf, -np.inf], np.nan)
zsc = ((roll_yield - roll_yield.rolling(252).mean()) / roll_yield.rolling(252).std()).replace([np.inf, -np.inf], np.nan)
carry_z_pos = np.sign(zsc).shift(1).reindex(f1c.index).fillna(0)

carry_mom_sig = roll_yield - roll_yield.shift(20)
carry_mom_pos = np.sign(carry_mom_sig).shift(1).reindex(f1c.index).fillna(0)

def v1_pos(k, N=1260, thr=0.10):
    p = crv[f"F{k}"].dropna()
    ma = p.rolling(N, min_periods=N // 2).mean()
    dev = ((p - ma) / ma).replace([np.inf, -np.inf], np.nan).dropna()
    sig = np.where(dev < -thr, 1.0, np.where(dev > thr, -1.0, 0.0))
    return pd.Series(sig, index=dev.index).shift(1).fillna(0)

val_pos = v1_pos(8)

regime_bd = (roll_yield.shift(1) > 0).reindex(f1c.index).fillna(False)
daily_ret = f1c.pct_change()
vol21 = daily_ret.rolling(21).std().shift(1)

# ── Align ───────────────────────────────────────────────────────────────────────
idx = f1c.index
for s in [mom_pos, carry_z_pos, carry_mom_pos, val_pos, vol21.dropna()]:
    idx = idx.intersection(s.index)
idx = idx.sort_values()

mom_a  = mom_pos.reindex(idx).fillna(0)
carz_a = carry_z_pos.reindex(idx).fillna(0)
carm_a = carry_mom_pos.reindex(idx).fillna(0)
val_a  = val_pos.reindex(idx).fillna(0)
bd_a   = regime_bd.reindex(idx).fillna(False)
vol_a  = vol21.reindex(idx)
f1c_a  = f1c.reindex(idx)
f1r_a  = f1r.reindex(idx)
T      = len(idx)

print(f"Common sample: {len(idx)} days | {idx[0].date()} -> {idx[-1].date()}\n")

# ── Portfolio variants ──────────────────────────────────────────────────────────
port_base3 = (mom_a + carz_a + val_a) / 3.0
port_4sleeve = (mom_a + carz_a + val_a + carm_a) / 4.0

def win_sharpe_net(pos_slice, oos_dt):
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

print("=" * 130)
print("CARRY MOMENTUM PLACEMENT TEST -- walk-forward, net 5bps  (base = Mom MA20,45 + Carry Zscore + Value V1F8, EW)")
print("=" * 130)
hdr = f"{'OOS':7s} {'Base3':>8s} {'4Sleeve':>8s} {'BD-Overlay':>11s} {'HiVol-Overlay':>14s}  |  {'%BD':>5s} {'%HiVol':>7s}"
print(hdr)
print("-" * 90)

rows = []
oos_s = IS_W
while oos_s < T:
    oos_e = min(oos_s + OOS_W, T)
    oos_yr = str(idx[oos_s].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
    is_slice = slice(oos_s - IS_W, oos_s)
    vol_thresh = np.nanmedian(vol_a.values[is_slice])

    oos_dt = idx[oos_s:oos_e]
    bd_oos = bd_a.reindex(oos_dt)
    hivol_oos = vol_a.reindex(oos_dt) > vol_thresh

    p_base3 = port_base3.reindex(oos_dt)
    p_4sl   = port_4sleeve.reindex(oos_dt)

    # Backwardation-overlay: 4-way blend when BD, base-3 EW when contango
    p_bd = pd.Series(np.where(bd_oos.values,
                               port_4sleeve.reindex(oos_dt).values,
                               port_base3.reindex(oos_dt).values), index=oos_dt)

    # High-vol overlay: 4-way blend when high-vol, base-3 EW when low-vol
    p_hv = pd.Series(np.where(hivol_oos.values,
                              port_4sleeve.reindex(oos_dt).values,
                              port_base3.reindex(oos_dt).values), index=oos_dt)

    sh_base3 = win_sharpe_net(p_base3, oos_dt)
    sh_4sl   = win_sharpe_net(p_4sl, oos_dt)
    sh_bd    = win_sharpe_net(p_bd, oos_dt)
    sh_hv    = win_sharpe_net(p_hv, oos_dt)

    print(f"{oos_yr:7s} {sh_base3:+8.3f} {sh_4sl:+8.3f} {sh_bd:+11.3f} {sh_hv:+14.3f}  |  "
          f"{100*bd_oos.mean():4.1f}% {100*hivol_oos.mean():6.1f}%")

    rows.append({"OOS": oos_yr, "Base3": sh_base3, "4Sleeve": sh_4sl, "BDOverlay": sh_bd, "HiVolOverlay": sh_hv})
    oos_s += OOS_W

df = pd.DataFrame(rows)
print("-" * 90)
for col in ["Base3", "4Sleeve", "BDOverlay", "HiVolOverlay"]:
    avg = np.nanmean(df[col])
    post22 = np.nanmean(df[df["OOS"].str.replace("*","").astype(int) >= 2022][col])
    print(f"{col:14s}  AVG={avg:+.3f}   POST-2022={post22:+.3f}")
