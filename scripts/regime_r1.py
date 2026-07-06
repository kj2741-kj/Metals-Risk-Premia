"""
regime_r1.py — R1 Term-Structure Regime Portfolio (Copper)

Regime is directly observed, not estimated: sign of the front-end roll yield
(F1-F2)/F1, shifted one day (same no-look-ahead convention as every other
signal in this project).

  Backwardation (roll_yield > 0): tightness -> upweight Momentum + Carry
  Contango      (roll_yield <= 0): surplus  -> upweight Value

Weights are pre-specified (not IS-fit), so there is nothing to overfit:
  Backwardation: Mom 0.40 / Carry 0.40 / Value 0.20
  Contango:      Mom 0.20 / Carry 0.20 / Value 0.60

Benchmark: static EW (1/3 each), same three live dashboard legs
  Copper: Mom MA(35,43) | Carry CarryMom-20d | Value V1 F8 5yr -- all shift-1.

Reports full-sample Sharpe and walk-forward OOS Sharpe (IS=5yr, OOS=1yr),
gross and net of 5bps round-trip TC on regime-driven turnover.
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

# ── Load F1 continuous + raw ────────────────────────────────────────────────────
f1_df = pd.read_csv(f"{BASE}\\LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()

# ── Load futures curve (Copper LME sheet) ──────────────────────────────────────
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
print(f"F1: {len(f1r)} rows | Curve: {len(crv)} rows | {crv.index[0].date()} -> {crv.index[-1].date()}\n")

# ── Live dashboard legs (Copper, all shift-1) ──────────────────────────────────
mom_pos = np.sign(f1r.rolling(35).mean() - f1r.rolling(43).mean()).shift(1).fillna(0)

roll_yield = ((crv["F1"] - crv["F2"]) / crv["F1"]).replace([np.inf, -np.inf], np.nan)
carry_sig  = roll_yield - roll_yield.shift(20)
carry_pos  = np.sign(carry_sig).shift(1).reindex(f1c.index).fillna(0)

def v1_pos(k, N=1260, thr=0.10):
    p   = crv[f"F{k}"].dropna()
    ma  = p.rolling(N, min_periods=N // 2).mean()
    dev = ((p - ma) / ma).replace([np.inf, -np.inf], np.nan).dropna()
    sig = np.where(dev < -thr, 1.0, np.where(dev > thr, -1.0, 0.0))
    return pd.Series(sig, index=dev.index).shift(1).fillna(0)

val_pos = v1_pos(8)

# ── R1 regime: sign of roll_yield itself, shift-1 (no look-ahead) ─────────────
regime_bd = (roll_yield.shift(1) > 0).reindex(f1c.index).fillna(False)   # True = backwardation

W_BD  = np.array([0.40, 0.40, 0.20])   # Mom, Carry, Value in backwardation
W_CTG = np.array([0.20, 0.20, 0.60])   # Mom, Carry, Value in contango

# ── Align ───────────────────────────────────────────────────────────────────
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

w_mom   = np.where(bd_a, W_BD[0], W_CTG[0])
w_carry = np.where(bd_a, W_BD[1], W_CTG[1])
w_val   = np.where(bd_a, W_BD[2], W_CTG[2])

port_r1 = pd.Series(w_mom * mom_a.values + w_carry * carry_a.values + w_val * val_a.values, index=idx)
port_ew = (mom_a + carry_a + val_a) / 3.0

print(f"Common sample: {len(idx)} days | {idx[0].date()} -> {idx[-1].date()}")
print(f"Backwardation days: {bd_a.sum()} ({100*bd_a.mean():.1f}%)  |  Contango days: {(~bd_a).sum()} ({100*(1-bd_a.mean()):.1f}%)\n")

# ── Sharpe helpers (gross + net, TC on F1_raw at position changes) ────────────
def sharpe(pos, gross=True):
    pnl = pos * f1c_a.diff()
    if not gross:
        chg = pos.diff().abs(); chg.iloc[0] = abs(pos.iloc[0])
        pnl = pnl - chg * (TC_BPS / 10000.0 / 2.0) * f1r_a
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pnl / f1c_a.shift(1)).replace([np.inf, -np.inf], np.nan)
    act = ret[pos != 0].dropna()
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def maxdd(pos):
    pnl = pos * f1c_a.diff()
    cum = pnl.cumsum()
    dd = cum - cum.cummax()
    return float(dd.min())

print("=" * 70)
print("FULL-SAMPLE COMPARISON")
print("=" * 70)
for lbl, pos in [("Static EW (1/3 each)", port_ew), ("R1 Regime-Weighted", port_r1)]:
    print(f"{lbl:24s}  gross={sharpe(pos,True):+.3f}  net5bps={sharpe(pos,False):+.3f}  "
          f"maxDD=${maxdd(pos):,.0f}/MT")
print()

# ── Walk-forward OOS (no re-fitting; weights are fixed a priori) ─────────────
print("=" * 100)
print("WALK-FORWARD OOS (IS=5yr / OOS=1yr windows, weights fixed a priori, net 5bps)")
print("=" * 100)
print(f"{'OOS':7s}  {'EW':>8s}  {'R1 Regime':>10s}  {'Diff':>8s}  {'%BD in window':>14s}")
print("-" * 100)
T = len(idx); rows = []
oos_s = IS_W
while oos_s < T:
    oos_e = min(oos_s + OOS_W, T)
    oos_yr = str(idx[oos_s].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
    ew_oos = port_ew.iloc[oos_s:oos_e]
    r1_oos = port_r1.iloc[oos_s:oos_e]
    bd_pct = 100 * bd_a.iloc[oos_s:oos_e].mean()

    def win_sharpe(pos_slice):
        oos_dt = idx[oos_s:oos_e]
        pos_slice = pos_slice.set_axis(oos_dt)
        c = f1c_a.reindex(oos_dt); r = f1r_a.reindex(oos_dt)
        pnl = pos_slice * c.diff()
        chg = pos_slice.diff().abs(); chg.iloc[0] = abs(pos_slice.iloc[0])
        pnl = pnl - chg * (TC_BPS / 10000.0 / 2.0) * r
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = (pnl / c.shift(1)).replace([np.inf, -np.inf], np.nan)
        act = ret[pos_slice != 0].dropna()
        if len(act) < 20: return np.nan
        sd = act.std(ddof=1)
        return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

    sh_ew = win_sharpe(ew_oos)
    sh_r1 = win_sharpe(r1_oos)
    diff = sh_r1 - sh_ew if not (np.isnan(sh_ew) or np.isnan(sh_r1)) else np.nan
    print(f"{oos_yr:7s}  {sh_ew:+8.3f}  {sh_r1:+10.3f}  {diff:+8.3f}  {bd_pct:13.1f}%")
    rows.append({"OOS": oos_yr, "EW": sh_ew, "R1": sh_r1})
    oos_s += OOS_W

df = pd.DataFrame(rows)
avg_ew = np.nanmean(df["EW"]); avg_r1 = np.nanmean(df["R1"])
print("-" * 100)
print(f"AVERAGE OOS SHARPE:  EW={avg_ew:+.3f}   R1 Regime={avg_r1:+.3f}   Diff={avg_r1-avg_ew:+.3f}")
post22 = df[df["OOS"].str.replace("*", "").astype(int) >= 2022]
print(f"POST-2022 AVERAGE:   EW={np.nanmean(post22['EW']):+.3f}   R1 Regime={np.nanmean(post22['R1']):+.3f}")
