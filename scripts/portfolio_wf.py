"""
portfolio_wf.py — Portfolio Walk-Forward: Methods A, B, C

Method A  EW Fixed        : (Mom + Carry + Val_F8_5yr) / 3, no re-opt
Method B  EW Val-Opt IS   : per window scan V1 F1–F15 × {1,3,5,7,10}yr, pick best IS, EW weights
Method C  Wt-Opt IS       : per window max-Sharpe QP on (Mom, Carry, Val_F8_5yr) in IS, apply OOS

IS = 5yr (1260d)  |  OOS = 1yr (252d)  |  Lag-1 (Mom, Val)  |  Same-Day (Carry)  |  0 TC gross
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

# ── Load data ──────────────────────────────────────────────────────────────────
f1_df = pd.read_csv(f"{BASE}\\LME_Copper_Rolling_F1_v2.csv",
                    parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()
print(f"F1: {len(f1r)} rows  |  {f1r.index[0].date()} → {f1r.index[-1].date()}")

xls     = pd.ExcelFile(f"{BASE}\\Metals Futures Curve.csv")
cu_sh   = next((s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()),
               next((s for s in xls.sheet_names if "copper" in s.lower()), xls.sheet_names[0]))
df_raw  = pd.read_excel(xls, sheet_name=cu_sh, header=None, nrows=5)
hr = [i for i in range(min(4, len(df_raw)))
      if any(kw in " ".join(str(v).lower() for v in df_raw.iloc[i].values if pd.notna(v))
             for kw in ["date", "f1", "f2", "price"])]
df_c = pd.read_excel(xls, sheet_name=cu_sh, header=hr if hr else [0,1])
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
print(f"Curve: {list(crv.columns)[:6]}... {len(crv)} rows\n")

# ── Helpers ────────────────────────────────────────────────────────────────────
def sharpe(ret: pd.Series, pos: pd.Series = None) -> float:
    act = ret[pos != 0].dropna() if pos is not None else ret.dropna()
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def oos_sharpe_from_pos(pos_slice: pd.Series, f1c_slice: pd.Series) -> float:
    idx = pos_slice.index.intersection(f1c_slice.index)
    p = pos_slice.reindex(idx); c = f1c_slice.reindex(idx)
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (p * c.diff() / c.shift(1)).replace([np.inf, -np.inf], np.nan)
    return sharpe(ret, p)

def max_sharpe_weights(ret_mat: np.ndarray) -> np.ndarray:
    """Max-Sharpe QP on (T,3) return matrix → w in [0,1], sum=1."""
    def neg_sh(w):
        r = ret_mat @ w
        act = r[r != 0]
        if len(act) < 20: return 0.0
        sd = act.std(ddof=1)
        return -(act.mean() / sd * np.sqrt(252)) if sd > 0 else 0.0
    best_val, best_w = np.inf, np.array([1/3, 1/3, 1/3])
    starts = [[1/3,1/3,1/3],[0.7,0.15,0.15],[0.15,0.7,0.15],[0.15,0.15,0.7],
              [0.5,0.5,0],[0.5,0,0.5],[0,0.5,0.5],[1,0,0],[0,1,0],[0,0,1]]
    for w0 in starts:
        try:
            r = minimize(neg_sh, w0, method="SLSQP",
                         bounds=[(0,1)]*3,
                         constraints=[{"type":"eq","fun":lambda w: w.sum()-1}],
                         options={"ftol":1e-9,"maxiter":500})
            if r.fun < best_val: best_val, best_w = r.fun, r.x
        except: pass
    w = np.clip(best_w, 0, 1); return w / w.sum()

# ── Build full-history signal positions ───────────────────────────────────────
print("Building signal positions (full history)...")

# Momentum MA(35,43) Lag-1
mom_pos = np.sign(f1r.rolling(35).mean() - f1r.rolling(43).mean()).shift(1).fillna(0)

# Carry (F1-F2)/F1 Same-Day
if "F1" in crv.columns and "F2" in crv.columns:
    cr_raw   = ((crv["F1"] - crv["F2"]) / crv["F1"]).replace([np.inf, -np.inf], np.nan).dropna()
    carry_pos = np.sign(cr_raw).reindex(f1c.index).fillna(0)
else:
    raise RuntimeError("F1 or F2 missing from curve data")

# Value V1 F8 5yr Lag-1 ±10% (fixed signal for A and C)
def v1_pos(k: int, N: int, thr: float = 0.10) -> pd.Series:
    col = f"F{k}"
    if col not in crv.columns: return pd.Series(dtype=float)
    p  = crv[col].dropna()
    ma = p.rolling(N, min_periods=N//2).mean()
    dev = ((p - ma) / ma).replace([np.inf, -np.inf], np.nan).dropna()
    sig = np.where(dev < -thr, 1.0, np.where(dev > thr, -1.0, 0.0))
    return pd.Series(sig, index=dev.index).shift(1).fillna(0)

val_fixed_pos = v1_pos(8, 1260)   # V1 F8 5yr — the IS-selected best

# Precompute all V1 variants for Method B grid (F1-F15 × 5 lookbacks)
print("Pre-computing V1 signal grid for Method B...")
VAL_GRID = {}
for k in range(1, 16):
    for N in [252, 756, 1260, 1764, 2520]:
        p = v1_pos(k, N)
        if not p.empty: VAL_GRID[(k, N)] = p

# Align common index across all three base signals
common_idx = f1c.index
for s in [mom_pos, carry_pos, val_fixed_pos]:
    common_idx = common_idx.intersection(s.index)
common_idx = common_idx.sort_values()
print(f"Common index: {len(common_idx)} days  |  {common_idx[0].date()} → {common_idx[-1].date()}\n")

mom_a   = mom_pos.reindex(common_idx).fillna(0)
carry_a = carry_pos.reindex(common_idx).fillna(0)
val_a   = val_fixed_pos.reindex(common_idx).fillna(0)
f1c_a   = f1c.reindex(common_idx)

T     = len(common_idx)
dates = common_idx

def ret_from_pos(pos: pd.Series) -> pd.Series:
    with np.errstate(invalid="ignore", divide="ignore"):
        return (pos * f1c_a.diff() / f1c_a.shift(1)).replace([np.inf, -np.inf], np.nan)

# ── Walk-forward ───────────────────────────────────────────────────────────────
print("=" * 110)
print("PORTFOLIO WALK-FORWARD  (IS=5yr, OOS=1yr, 0 TC, gross)  [Methods A / B / C + Individual signals]")
print("=" * 110)

hdr = (f"{'OOS':7s}  {'IS Period':12s}  "
       f"{'Mom':>7s}  {'Carry':>7s}  {'ValFix':>7s}  "
       f"{'[A] EW-Fix':>11s}  "
       f"{'[B] EW-ValOpt':>14s}  {'B: Best Val':>14s}  "
       f"{'[C] Wt-Opt':>11s}  {'C: Weights (M/C/V)':>22s}")
print(hdr)
print("-" * 110)

rows = []
oos_s = IS_W
while oos_s < T:
    oos_e   = min(oos_s + OOS_W, T)
    is_s    = oos_s - IS_W
    oos_yr  = str(dates[oos_s].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
    is_lbl  = f"{dates[is_s].year}–{dates[oos_s-1].year}"
    oos_dt  = dates[oos_s:oos_e]

    m_oos  = mom_a.iloc[oos_s:oos_e].set_axis(oos_dt)
    c_oos  = carry_a.iloc[oos_s:oos_e].set_axis(oos_dt)
    v_oos  = val_a.iloc[oos_s:oos_e].set_axis(oos_dt)
    f_oos  = f1c_a.iloc[oos_s:oos_e].set_axis(oos_dt)

    # Individual OOS
    sh_mom   = oos_sharpe_from_pos(m_oos, f_oos)
    sh_carry = oos_sharpe_from_pos(c_oos, f_oos)
    sh_val   = oos_sharpe_from_pos(v_oos, f_oos)

    # ── Method A: EW Fixed ─────────────────────────────────────────────────────
    port_a = (m_oos + c_oos + v_oos) / 3.0
    sh_a   = oos_sharpe_from_pos(port_a, f_oos)

    # ── Method B: EW Val-Opt (IS scan) ────────────────────────────────────────
    is_f   = f1c_a.iloc[is_s:oos_s]
    best_b_sh, best_b_key, best_b_pos_oos = -np.inf, (8, 1260), v_oos
    for (k, N), full_pos in VAL_GRID.items():
        p_is = full_pos.reindex(common_idx[is_s:oos_s]).fillna(0)
        if p_is.empty or (p_is == 0).all(): continue
        with np.errstate(invalid="ignore", divide="ignore"):
            r = (p_is * is_f.diff() / is_f.shift(1)).replace([np.inf, -np.inf], np.nan)
        s = sharpe(r, p_is)
        if not np.isnan(s) and s > best_b_sh:
            best_b_sh, best_b_key = s, (k, N)
    # Use OOS position of the IS-selected best variant
    best_b_pos_full = VAL_GRID.get(best_b_key, val_fixed_pos)
    b_val_oos = best_b_pos_full.reindex(oos_dt).fillna(0)
    port_b    = (m_oos + c_oos + b_val_oos) / 3.0
    sh_b      = oos_sharpe_from_pos(port_b, f_oos)
    b_lbl     = f"F{best_b_key[0]} {best_b_key[1]//252}yr"

    # ── Method C: Max-Sharpe weight optimizer on IS period ────────────────────
    def is_ret(pos_full: pd.Series) -> np.ndarray:
        p = pos_full.reindex(common_idx[is_s:oos_s]).fillna(0)
        with np.errstate(invalid="ignore", divide="ignore"):
            r = (p * is_f.diff() / is_f.shift(1)).replace([np.inf, -np.inf], np.nan)
        return r.fillna(0).values

    ret_mat = np.column_stack([is_ret(mom_a), is_ret(carry_a), is_ret(val_a)])
    w_opt   = max_sharpe_weights(ret_mat)
    port_c  = w_opt[0]*m_oos + w_opt[1]*c_oos + w_opt[2]*v_oos
    sh_c    = oos_sharpe_from_pos(port_c, f_oos)
    w_lbl   = f"{w_opt[0]:.2f}/{w_opt[1]:.2f}/{w_opt[2]:.2f}"

    def F(x): return f"{x:+.3f}" if x is not None and not np.isnan(x) else "   —  "

    line = (f"{oos_yr:7s}  {is_lbl:12s}  "
            f"{F(sh_mom):>7s}  {F(sh_carry):>7s}  {F(sh_val):>7s}  "
            f"{F(sh_a):>11s}  "
            f"{F(sh_b):>14s}  {b_lbl:>14s}  "
            f"{F(sh_c):>11s}  {w_lbl:>22s}")
    print(line)

    rows.append({
        "OOS": oos_yr, "IS": is_lbl,
        "Mom": sh_mom, "Carry": sh_carry, "ValFix": sh_val,
        "A_EW_Fix": sh_a,
        "B_EW_ValOpt": sh_b, "B_BestVal": b_lbl,
        "C_WtOpt": sh_c, "C_wM": w_opt[0], "C_wC": w_opt[1], "C_wV": w_opt[2],
    })
    oos_s += OOS_W

df = pd.DataFrame(rows)

def avg_sh(col): return np.nanmean([r for r in df[col] if r is not None and not np.isnan(r)])
def pct_pos(col): return 100 * np.nanmean([1 if (r is not None and not np.isnan(r) and r > 0) else 0
                                            for r in df[col]])
def post22(col):
    vals = [r for r, y in zip(df[col], df["OOS"])
            if r is not None and not np.isnan(r) and int(y.rstrip("*")) >= 2022]
    return np.nanmean(vals) if vals else np.nan

print("-" * 110)

def sumline(label, col):
    return (f"{label:20s}  {'':12s}  "
            f"{avg_sh('Mom'):>7.3f}  {avg_sh('Carry'):>7.3f}  {avg_sh('ValFix'):>7.3f}  "
            f"{avg_sh('A_EW_Fix'):>11.3f}  "
            f"{avg_sh('B_EW_ValOpt'):>14.3f}  {'':>14s}  "
            f"{avg_sh('C_WtOpt'):>11.3f}")

print(f"{'AVG OOS Sharpe':20s}  {'':12s}  "
      f"{avg_sh('Mom'):>7.3f}  {avg_sh('Carry'):>7.3f}  {avg_sh('ValFix'):>7.3f}  "
      f"{avg_sh('A_EW_Fix'):>11.3f}  "
      f"{avg_sh('B_EW_ValOpt'):>14.3f}  {'':>14s}  "
      f"{avg_sh('C_WtOpt'):>11.3f}")

print(f"{'% Positive Windows':20s}  {'':12s}  "
      f"{pct_pos('Mom'):>6.1f}%  {pct_pos('Carry'):>6.1f}%  {pct_pos('ValFix'):>6.1f}%  "
      f"{pct_pos('A_EW_Fix'):>10.1f}%  "
      f"{pct_pos('B_EW_ValOpt'):>13.1f}%  {'':>14s}  "
      f"{pct_pos('C_WtOpt'):>10.1f}%")

print(f"{'Avg Post-2022':20s}  {'':12s}  "
      f"{post22('Mom'):>7.3f}  {post22('Carry'):>7.3f}  {post22('ValFix'):>7.3f}  "
      f"{post22('A_EW_Fix'):>11.3f}  "
      f"{post22('B_EW_ValOpt'):>14.3f}  {'':>14s}  "
      f"{post22('C_WtOpt'):>11.3f}")

print("=" * 110)
print()

# ── Method C weight analysis ───────────────────────────────────────────────────
print("METHOD C — IS-Optimised Weights Per Window")
print(f"{'OOS':7s}  {'IS Period':12s}  {'w_Mom':>7s}  {'w_Carry':>8s}  {'w_Val':>7s}  {'IS selection notes':>30s}")
print("-" * 75)
for _, r in df.iterrows():
    dom = "Mom" if r.C_wM >= 0.5 else ("Carry" if r.C_wC >= 0.5 else ("Val" if r.C_wV >= 0.5 else "Mixed"))
    print(f"{r.OOS:7s}  {r.IS:12s}  {r.C_wM:>7.3f}  {r.C_wC:>8.3f}  {r.C_wV:>7.3f}  dominant: {dom}")

print()
avg_wm = np.nanmean(df["C_wM"]); avg_wc = np.nanmean(df["C_wC"]); avg_wv = np.nanmean(df["C_wV"])
print(f"Average weights across all windows:  w_Mom={avg_wm:.3f}  w_Carry={avg_wc:.3f}  w_Val={avg_wv:.3f}")
print()

# ── Method B best-val selection analysis ──────────────────────────────────────
print("METHOD B — IS-Selected Value Variant Per Window")
print(f"{'OOS':7s}  {'IS Period':12s}  {'IS-Best Val':>14s}  {'OOS Val Sharpe':>16s}  {'Note':>20s}")
print("-" * 75)
for _, r in df.iterrows():
    note = "= IS-opt" if r.B_BestVal == "F8 5yr" else "different from F8 5yr"
    print(f"{r.OOS:7s}  {r.IS:12s}  {r.B_BestVal:>14s}  {r.B_EW_ValOpt:>16.3f}  {note:>20s}")

print()

# ── Summary comparison ────────────────────────────────────────────────────────
print("=" * 70)
print("SUMMARY COMPARISON")
print("=" * 70)
methods = [
    ("Momentum (standalone)", "Mom"),
    ("Carry (standalone)",    "Carry"),
    ("Value F8 5yr (standalone)", "ValFix"),
    ("A: EW Fixed (1/3 each)", "A_EW_Fix"),
    ("B: EW + IS-opt Value",   "B_EW_ValOpt"),
    ("C: IS-opt Weights",      "C_WtOpt"),
]
print(f"{'Method':32s}  {'Avg OOS':>9s}  {'% Pos Win':>10s}  {'Post-2022':>10s}")
print("-" * 70)
for lbl, col in methods:
    n_windows = sum(1 for r in df[col] if r is not None and not np.isnan(r))
    print(f"{lbl:32s}  {avg_sh(col):>+9.3f}  {pct_pos(col):>9.1f}%  {post22(col):>+10.3f}")

print("=" * 70)
print()
print("Notes:")
print("  A: Fixed params throughout — MA(35,43), (F1-F2)/F1, V1 F8 5yr ±10%.")
print("  B: Mom+Carry fixed; Value re-selected per IS window from V1 F1-F15 x {1,3,5,7,10}yr.")
print("  C: EW-fixed signals; weights re-optimised per IS window via max-Sharpe QP.")
print("  All entries Lag-1 except Carry (Same-Day). Gross, 0 TC.")
print("  Partial windows (marked *) included in averages.")
