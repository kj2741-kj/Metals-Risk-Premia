"""
bucket_weight_opt.py
Per IS window:
  1. Pick best pair from each min-span bucket (IS Sharpe scan)
       Fast   : n <= 30, span >= 5
       Medium : 31 <= n <= 60, span >= 8
       Slow   : n >  60, span >= 15
  2. Run max-Sharpe weight optimizer on those 3 IS-selected pairs
  3. Apply optimised weights OOS for 1 year → roll

Compare vs:
  - MA(35,43) fixed               (+0.492)
  - Min-span EW (equal weights)   (+0.000)
  - Fixed anchors IS-opt weights  (+0.442)
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

IS, OOS, MAX_N = 1260, 252, 126
BATCH = 800

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

def max_sharpe_weights(ret_mat_np):
    """Max-Sharpe QP on (T,3) IS return matrix. Returns w summing to 1, all >= 0."""
    def neg_sh(w):
        combined = ret_mat_np @ w
        act = combined[combined != 0]
        if len(act) < 20: return 0.0
        sd = act.std(ddof=1)
        return -(act.mean() / sd * np.sqrt(252)) if sd > 0 else 0.0

    best_val, best_w = np.inf, np.array([1/3, 1/3, 1/3])
    for w0 in [[1/3,1/3,1/3],[0.8,0.1,0.1],[0.1,0.8,0.1],[0.1,0.1,0.8],
               [0.5,0.5,0.0],[0.5,0.0,0.5],[0.0,0.5,0.5],
               [1.0,0.0,0.0],[0.0,1.0,0.0],[0.0,0.0,1.0]]:
        try:
            res = minimize(neg_sh, w0, method="SLSQP",
                           bounds=[(0,1)]*3,
                           constraints=[{"type":"eq","fun":lambda w: w.sum()-1}],
                           options={"ftol":1e-9,"maxiter":500})
            if res.fun < best_val:
                best_val, best_w = res.fun, res.x
        except: pass
    w = np.clip(best_w, 0, 1)
    return w / w.sum()

def F(x): return f"{x:+.3f}" if (x is not None and not np.isnan(x)) else "  —"

# ── Pre-compute SMA matrix ─────────────────────────────────────────────────────
T = len(f1r); dates = f1r.index; f1c_np = f1c.values.astype(float)
print("Pre-computing SMA matrix...", flush=True)
sma = np.column_stack([f1r.rolling(k+1).mean().values for k in range(MAX_N)])

delta_np = np.empty(T); delta_np[0] = np.nan; delta_np[1:] = np.diff(f1c_np)
prev_np  = np.empty(T); prev_np[0]  = np.nan; prev_np[1:]  = f1c_np[:-1]

pairs = [(m+1, n+1) for m in range(MAX_N) for n in range(m+1, MAX_N)]
m_arr = np.array([p[0] for p in pairs])
n_arr = np.array([p[1] for p in pairs])
spans = n_arr - m_arr
m_idx = m_arr - 1
n_idx = n_arr - 1
NP    = len(pairs)

# ── Min-span bucket masks ──────────────────────────────────────────────────────
BUCKETS = {
    "Fast   (n≤30,  span≥5) ": (n_arr <= 30) & (spans >= 5),
    "Medium (31≤n≤60,span≥8)": (n_arr >= 31) & (n_arr <= 60) & (spans >= 8),
    "Slow   (n>60,  span≥15)": (n_arr >  60) & (spans >= 15),
}
BKEYS = list(BUCKETS.keys())

print("\nMin-span buckets:")
for name, mask in BUCKETS.items():
    ss = spans[mask]
    print(f"  {name}: {mask.sum():4d} pairs  span=[{ss.min()}–{ss.max()}]")

# ── Pre-compute anchor MA(35,43) position for benchmark ───────────────────────
ma35_pos_full = L1(np.sign(f1r.rolling(35).mean() - f1r.rolling(43).mean()))

# ── Walk-forward ───────────────────────────────────────────────────────────────
rows = []
oos_s = IS

while oos_s < T:
    oos_e     = min(oos_s + OOS, T)
    is_s      = oos_s - IS
    oos_dates = dates[oos_s:oos_e]
    oos_len   = len(oos_dates)
    oos_yr    = str(dates[oos_s].year) + ("*" if oos_len < OOS else "")
    is_lbl    = f"{dates[is_s].year}–{dates[oos_s-1].year}"

    sma_is    = sma[is_s:oos_s]
    d_is      = delta_np[is_s:oos_s]
    p_is      = prev_np[is_s:oos_s]
    base_mask = np.isfinite(d_is) & (p_is > 0)

    # ── IS scan: Sharpe for all pairs ─────────────────────────────────────────
    all_sh = np.full(NP, np.nan)
    for b0 in range(0, NP, BATCH):
        b1   = min(b0+BATCH, NP)
        mb   = m_idx[b0:b1]; nb = n_idx[b0:b1]
        diff = sma_is[:, mb] - sma_is[:, nb]
        sig  = np.sign(diff)
        pos  = np.zeros_like(sig)
        pos[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
        act  = (pos != 0) & base_mask[:, None]
        with np.errstate(invalid="ignore", divide="ignore"):
            ret = np.where(act, pos * d_is[:, None] / p_is[:, None], np.nan)
        mu  = np.nanmean(ret, axis=0)
        sd  = np.nanstd(ret, axis=0, ddof=1)
        n_a = np.sum(act, axis=0)
        all_sh[b0:b1] = np.where((sd>0)&(n_a>=20), mu/sd*np.sqrt(252), np.nan)

    # ── Pick best pair per bucket ──────────────────────────────────────────────
    f1r_w   = f1r.iloc[is_s:oos_e]
    f1c_oos = f1c.reindex(oos_dates)
    delta_w = f1c.iloc[is_s:oos_e].diff()
    prev_w  = f1c.iloc[is_s:oos_e].shift(1)

    selected_pairs = []
    is_ret_list    = []
    oos_pos_list   = []

    for bname, bmask in BUCKETS.items():
        b_sh = np.where(bmask, all_sh, np.nan)
        if np.all(np.isnan(b_sh)): continue
        bi   = int(np.nanargmax(b_sh))
        pm, pn = pairs[bi]
        selected_pairs.append((pm, pn, all_sh[bi], bname.strip()))

        # IS return series for this pair
        pos_full = L1(np.sign(f1r_w.rolling(pm).mean() - f1r_w.rolling(pn).mean()))
        with np.errstate(invalid="ignore", divide="ignore"):
            ret_is = (pos_full * delta_w / prev_w).replace([np.inf,-np.inf], np.nan)
        is_ret_list.append(ret_is.values)

        # OOS position for this pair
        oos_pos = pos_full.iloc[-oos_len:].set_axis(oos_dates)
        oos_pos_list.append(oos_pos)

    if len(selected_pairs) < 3:
        oos_s += OOS; continue

    # ── IS weight optimisation on the 3 selected pairs ────────────────────────
    is_ret_mat = np.column_stack(is_ret_list)
    is_ret_mat = np.nan_to_num(is_ret_mat, nan=0.0)
    w_opt = max_sharpe_weights(is_ret_mat)

    # ── OOS: equal-weight vs optimised-weight combined position ───────────────
    oos_ret_mat = pd.concat(oos_pos_list, axis=1)

    # equal-weight
    eq_combined  = oos_ret_mat.mean(axis=1)
    eq_oos_ret   = (eq_combined * f1c_oos.diff() / f1c_oos.shift(1)).replace([np.inf,-np.inf], np.nan)
    eq_sh        = sharpe_from_ret(eq_oos_ret)

    # optimised weight
    opt_combined_pos = pd.Series(oos_ret_mat.values @ w_opt, index=oos_dates)
    opt_oos_ret      = (opt_combined_pos * f1c_oos.diff() / f1c_oos.shift(1)).replace([np.inf,-np.inf], np.nan)
    opt_sh           = sharpe_from_ret(opt_oos_ret)

    # MA(35,43) benchmark
    ma35_sh = sharpe_pos(ma35_pos_full.reindex(oos_dates), f1c_oos)

    sp = selected_pairs
    row = {
        "OOS":       oos_yr,
        "IS window": is_lbl,
        "MA(35,43)": F(ma35_sh),
        "EW (minspan)": F(eq_sh),
        "Opt weights":  F(opt_sh),
        "w_fast": f"{w_opt[0]:.2f}",
        "w_med":  f"{w_opt[1]:.2f}",
        "w_slow": f"{w_opt[2]:.2f}",
        "Fast":   f"MA({sp[0][0]},{sp[0][1]}) IS={sp[0][2]:.2f}",
        "Medium": f"MA({sp[1][0]},{sp[1][1]}) IS={sp[1][2]:.2f}",
        "Slow":   f"MA({sp[2][0]},{sp[2][1]}) IS={sp[2][2]:.2f}",
    }
    rows.append(row)
    print(f"  {oos_yr} ({is_lbl}):  MA35={row['MA(35,43)']}  "
          f"EW={row['EW (minspan)']}  OPT={row['Opt weights']}  "
          f"w=({w_opt[0]:.2f},{w_opt[1]:.2f},{w_opt[2]:.2f})  "
          f"| F:{row['Fast']}  M:{row['Medium']}  S:{row['Slow']}", flush=True)
    oos_s += OOS

df = pd.DataFrame(rows)

pd.set_option("display.width", 240)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

def avg_col(col):
    vals = []
    for v in df[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print("\n" + "="*200)
print("BUCKET SELECTION + WEIGHT OPTIMISATION  (IS=5yr, OOS=1yr, Lag-1)")
print("Step 1: Best pair per min-span bucket from IS scan")
print("Step 2: Max-Sharpe weight optimiser on those 3 IS-selected pairs")
print("="*200)
cols = ["OOS","IS window","MA(35,43)","EW (minspan)","Opt weights",
        "w_fast","w_med","w_slow","Fast","Medium","Slow"]
print(df[cols].to_string(index=False))

print(f"\nAVERAGE OOS SHARPE:")
print(f"  MA(35,43) fixed              : {avg_col('MA(35,43)'):+.3f}  ← benchmark")
print(f"  Min-span EW (equal weights)  : {avg_col('EW (minspan)'):+.3f}")
print(f"  Min-span + Opt weights       : {avg_col('Opt weights'):+.3f}")
print(f"  Fixed anchors + IS-opt       : +0.442  (from previous run)")

post22 = df[df["OOS"].str.replace("*","").str.strip().astype(int) >= 2022].copy()
print(f"\nPOST-2022 AVERAGE:")
for col in ["MA(35,43)", "EW (minspan)", "Opt weights"]:
    print(f"  {col:30s}: {avg_col.__wrapped__(post22, col) if hasattr(avg_col,'__wrapped__') else (lambda c: np.nanmean([float(str(v).split()[0].strip().lstrip('+')) for v in post22[c] if str(v).replace('.','').replace('-','').replace('+','').strip().replace('—','').isdigit() or (str(v).split()[0].lstrip('+-').replace('.','').isdigit())]))(col):+.3f}")

# Cleaner post-2022
print(f"\nPOST-2022 AVERAGE (2022–2024):")
for col in ["MA(35,43)", "EW (minspan)", "Opt weights"]:
    vals = []
    for v in post22[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    print(f"  {col:30s}: {np.nanmean(vals):+.3f}" if vals else f"  {col}: —")

# Summary of weight assignments
print(f"\nWEIGHT ASSIGNMENTS ACROSS WINDOWS:")
print(f"  {'OOS':6s}  {'w_fast':7s}  {'w_med':6s}  {'w_slow':7s}  Fast pair      Medium pair    Slow pair")
for _, r in df.iterrows():
    if r["OOS"] == "2025*": continue
    print(f"  {r['OOS']:6s}  {r['w_fast']:7s}  {r['w_med']:6s}  {r['w_slow']:7s}  "
          f"{r['Fast']:15s}  {r['Medium']:15s}  {r['Slow']:15s}")
