"""
value_oos.py — OOS Validation for Value Signals (V1 and V2)

V1 MA Reversion: deviation = (Fk - MA_N) / MA_N; ±10% threshold → −1/0/+1, Lag-1
V2 Baz-Granger:  reversal  = F1_raw[t−N] − F1_raw[t]; sign → ±1, Lag-1

Three analyses:
  A. IS/OOS Hard Split (2016 cutoff):
     - V1: IS-optimal params → apply OOS
     - V2: compare 3yr/5yr/7yr (10yr can only start 2016 — no IS period possible)
  B. Walk-Forward Fixed Params (IS=5yr, OOS=1yr):
     - V1 fixed at F8 5yr (IS-selected from full data)
     - V2 fixed at BG 10yr (limited to windows from 2016)
  C. Walk-Forward Re-Optimization (V1 only):
     - Per IS window: scan F1-F15 × {1yr,3yr,5yr} → apply OOS
     - Compare vs fixed F8 5yr
"""

import os, sys, io, warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
DATA_DIR = os.path.join(_REPO_ROOT, "data")
OUTPUTS_DIR = os.path.join(_REPO_ROOT, "outputs")
BASE = DATA_DIR

# ── Load data ──────────────────────────────────────────────────────────────────
f1_df = pd.read_csv(os.path.join(BASE, "LME_Copper_Rolling_F1_v2.csv"),
                    parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()
print(f"F1 data: {len(f1r)} rows  |  {f1r.index[0].date()} -> {f1r.index[-1].date()}")

# Load curve
curve_path = os.path.join(BASE, "Metals Futures Curve.csv")
xls = pd.ExcelFile(curve_path)
cu_sheet = next((s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()), xls.sheet_names[0])
df_raw = pd.read_excel(xls, sheet_name=cu_sheet, header=None, nrows=5)
header_rows = []
for i in range(min(4, len(df_raw))):
    row_vals = [str(v).strip().lower() for v in df_raw.iloc[i].values if pd.notna(v)]
    if any(kw in " ".join(row_vals) for kw in ["date", "f1", "f2", "price"]):
        header_rows.append(i)
if len(header_rows) >= 2:
    df_c = pd.read_excel(xls, sheet_name=cu_sheet, header=header_rows)
elif len(header_rows) == 1:
    df_c = pd.read_excel(xls, sheet_name=cu_sheet, header=header_rows[0])
else:
    df_c = pd.read_excel(xls, sheet_name=cu_sheet, header=[0, 1])

if isinstance(df_c.columns, pd.MultiIndex):
    new_cols = []
    for col_tuple in df_c.columns:
        parts = [str(p).strip() for p in col_tuple
                 if pd.notna(p) and "Unnamed" not in str(p) and str(p).strip()]
        new_cols.append("_".join(parts) if parts else str(col_tuple))
    df_c.columns = new_cols
df_c.columns = [str(c).strip() for c in df_c.columns]
date_col = [c for c in df_c.columns if "date" in c.lower()]
if date_col:
    df_c = df_c.rename(columns={date_col[0]: "Date"})
df_c["Date"] = pd.to_datetime(df_c["Date"], errors="coerce")
df_c = df_c.dropna(subset=["Date"]).set_index("Date").sort_index()
df_c.index = df_c.index.normalize()

curve_prices = {}
for col in df_c.columns:
    col_lower = col.lower().replace(" ", "_")
    for i in range(1, 27):
        if f"f{i}_" in col_lower and "price" in col_lower:
            curve_prices[f"F{i}"] = pd.to_numeric(df_c[col], errors="coerce")
            break
curve_px = pd.DataFrame(curve_prices, index=df_c.index)
print(f"Curve: F1-F{max(int(c[1:]) for c in curve_px.columns)}  |  {len(curve_px)} rows")


# ── Helpers ────────────────────────────────────────────────────────────────────
LOOKBACKS = {"1yr": 252, "3yr": 756, "5yr": 1260, "7yr": 1764, "10yr": 2520}
CONTRACTS = list(range(1, 16))   # F1 to F15
THR = 0.10                        # ±10% threshold for V1

def sharpe(ret: pd.Series) -> float:
    act = ret[ret != 0].dropna()
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def v1_signal(k: int, N: int, price_df: pd.DataFrame) -> pd.Series:
    """V1 MA Reversion position series (Lag-1)."""
    col = f"F{k}"
    if col not in price_df.columns: return pd.Series(dtype=float)
    price = price_df[col].dropna()
    if len(price) < max(N // 2, 60): return pd.Series(dtype=float)
    ma  = price.rolling(N, min_periods=max(N // 2, 60)).mean()
    dev = (price - ma) / ma.replace(0, np.nan)
    dev = dev.replace([np.inf, -np.inf], np.nan).dropna()
    sig = np.where(dev.values < -THR, 1.0, np.where(dev.values > THR, -1.0, 0.0))
    return pd.Series(sig, index=dev.index).shift(1).fillna(0)

def v2_signal(N: int) -> pd.Series:
    """V2 BG Reversal position series (Lag-1)."""
    rev = f1r.shift(N) - f1r
    rev = rev.replace([np.inf, -np.inf], np.nan).dropna()
    return np.sign(rev).shift(1).fillna(0)

def compute_sharpe_on_period(pos: pd.Series, start=None, end=None) -> float:
    """Sharpe on a sub-period [start, end)."""
    idx = pos.index.intersection(f1c.index)
    if start: idx = idx[idx >= pd.Timestamp(start)]
    if end:   idx = idx[idx <  pd.Timestamp(end)]
    if len(idx) < 20: return np.nan
    p = pos.reindex(idx).fillna(0); c = f1c.reindex(idx)
    pnl = p * c.diff()
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pnl / c.shift(1)).replace([np.inf, -np.inf], np.nan)
    return sharpe(ret)

def ann_ret_on_period(pos: pd.Series, start=None, end=None) -> float:
    idx = pos.index.intersection(f1c.index)
    if start: idx = idx[idx >= pd.Timestamp(start)]
    if end:   idx = idx[idx <  pd.Timestamp(end)]
    if len(idx) < 20: return np.nan
    p = pos.reindex(idx).fillna(0); c = f1c.reindex(idx)
    pnl = p * c.diff()
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pnl / c.shift(1)).replace([np.inf, -np.inf], np.nan)
    return float(ret.dropna().mean() * 252 * 100)

def F(x):
    return f"{x:+.3f}" if (x is not None and not np.isnan(x)) else "     -"


# ══════════════════════════════════════════════════════════════════════════════
# FULL-PERIOD IS REFERENCE (from value_signals.py)
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*80)
print("FULL-PERIOD IS REFERENCE (all data 2006-2025, no holdout)")
print("="*80)
print("\nV1 MA Reversion — all contracts F1-F15 x lookback (Lag-1, ±10% threshold):")
v1_full_results = {}
for N_lbl, N in LOOKBACKS.items():
    row = []
    for k in CONTRACTS:
        pos = v1_signal(k, N, curve_px)
        sh  = compute_sharpe_on_period(pos)
        v1_full_results[(k, N_lbl)] = sh
        row.append(f"F{k:2d}:{sh:+.3f}" if not np.isnan(sh) else f"F{k:2d}:  -  ")
    best_k = max(CONTRACTS, key=lambda k: v1_full_results.get((k, N_lbl), np.nan) or np.nan)
    best_sh = v1_full_results.get((best_k, N_lbl), np.nan)
    print(f"  {N_lbl:5s}: best=F{best_k} {F(best_sh)}  | {' '.join(row[:8])}")

print("\nV2 BG Reversal — all lookbacks (Lag-1):")
v2_full_results = {}
for N_lbl, N in LOOKBACKS.items():
    pos = v2_signal(N)
    sh  = compute_sharpe_on_period(pos)
    ar  = ann_ret_on_period(pos)
    v2_full_results[N_lbl] = sh
    first_valid = f1r.index[f1r.index >= f1r.shift(N).first_valid_index()][0] if f1r.shift(N).dropna().shape[0] > 0 else None
    first_str = first_valid.strftime("%Y-%m") if first_valid is not None else "N/A"
    print(f"  {N_lbl:5s}: Sharpe={F(sh)}  Ann Ret={ar:+.1f}%  [signal valid from {first_str}]")


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS A: IS / OOS HARD SPLIT
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*80)
print("ANALYSIS A: IS / OOS HARD SPLIT")
print("IS = 2006-2015  |  OOS = 2016-2025")
print("="*80)

IS_END_A  = "2016-01-01"
OOS_END_A = None

# V1 Hard Split
print("\n--- V1: IS scan -> apply IS-selected params OOS ---")
v1_is_results = {}
for N_lbl, N in LOOKBACKS.items():
    for k in CONTRACTS:
        pos = v1_signal(k, N, curve_px)
        sh  = compute_sharpe_on_period(pos, end=IS_END_A)
        v1_is_results[(k, N_lbl)] = sh

best_k_a  = max([(k, l) for k, l in v1_is_results], key=lambda kl: v1_is_results.get(kl, np.nan) or np.nan)
best_sh_is = v1_is_results[best_k_a]
best_pos_a = v1_signal(best_k_a[0], LOOKBACKS[best_k_a[1]], curve_px)
best_sh_oos = compute_sharpe_on_period(best_pos_a, start=IS_END_A)
best_ar_oos = ann_ret_on_period(best_pos_a, start=IS_END_A)

# Fixed F8 5yr OOS for comparison
pos_f8_5yr = v1_signal(8, 1260, curve_px)
sh_f8_is   = compute_sharpe_on_period(pos_f8_5yr, end=IS_END_A)
sh_f8_oos  = compute_sharpe_on_period(pos_f8_5yr, start=IS_END_A)
ar_f8_oos  = ann_ret_on_period(pos_f8_5yr, start=IS_END_A)

# F12 5yr (Mark's reference)
pos_f12_5yr = v1_signal(12, 1260, curve_px)
sh_f12_is   = compute_sharpe_on_period(pos_f12_5yr, end=IS_END_A)
sh_f12_oos  = compute_sharpe_on_period(pos_f12_5yr, start=IS_END_A)

print(f"\n  IS-optimal (from IS scan):    F{best_k_a[0]} / {best_k_a[1]}")
print(f"    IS Sharpe (2006-2015) :     {F(best_sh_is)}")
print(f"    OOS Sharpe (2016-2025):     {F(best_sh_oos)}  Ann Ret {best_ar_oos:+.1f}%")
print(f"\n  Fixed F8 5yr (full-IS best):  ")
print(f"    IS Sharpe (2006-2015) :     {F(sh_f8_is)}")
print(f"    OOS Sharpe (2016-2025):     {F(sh_f8_oos)}  Ann Ret {ar_f8_oos:+.1f}%")
print(f"\n  Fixed F12 5yr (Mark's ref):   ")
print(f"    IS Sharpe (2006-2015) :     {F(sh_f12_is)}")
print(f"    OOS Sharpe (2016-2025):     {F(sh_f12_oos)}  Ann Ret {ann_ret_on_period(pos_f12_5yr, start=IS_END_A):+.1f}%")

# Top 5 by IS for context
print(f"\n  Top 5 V1 configs by IS Sharpe (2006-2015):")
sorted_is = sorted(v1_is_results.items(), key=lambda x: x[1] if not np.isnan(x[1] or np.nan) else -99, reverse=True)
for (k, lbl), sh_is in sorted_is[:5]:
    pos_tmp = v1_signal(k, LOOKBACKS[lbl], curve_px)
    sh_oos = compute_sharpe_on_period(pos_tmp, start=IS_END_A)
    print(f"    F{k:2d} {lbl:5s}: IS={F(sh_is)}  OOS={F(sh_oos)}")

# V2 Hard Split (note: 10yr is problematic)
print("\n--- V2: lookback comparison (IS cannot include 10yr — explained below) ---")
print("  NOTE: V2 10yr signal first valid Jan 2016 = IS end date -> no IS period for 10yr")
print("        Evaluating 3yr/5yr/7yr which have IS data; 10yr shown OOS-only\n")
for N_lbl, N in LOOKBACKS.items():
    pos = v2_signal(N)
    sh_is  = compute_sharpe_on_period(pos, end=IS_END_A)
    sh_oos = compute_sharpe_on_period(pos, start=IS_END_A)
    ar_oos = ann_ret_on_period(pos, start=IS_END_A)
    if N_lbl == "10yr":
        print(f"  V2 {N_lbl}: IS=  N/A (no data)   OOS={F(sh_oos)}  Ann Ret {ar_oos:+.1f}%  [only 9yr of signal]")
    else:
        print(f"  V2 {N_lbl}: IS={F(sh_is)}             OOS={F(sh_oos)}  Ann Ret {ar_oos:+.1f}%")


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS B: WALK-FORWARD FIXED PARAMS (IS=5yr, OOS=1yr)
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*80)
print("ANALYSIS B: WALK-FORWARD — FIXED PARAMS (IS=5yr rolling, OOS=1yr)")
print("No re-optimization: tests whether IS-selected params degrade in unseen windows")
print("="*80)

IS_W, OOS_W = 1260, 252

def walkforward_fixed(pos_full: pd.Series, label: str) -> dict:
    """Apply fixed-param position series walk-forward; compute OOS Sharpe per window."""
    idx   = pos_full.index.intersection(f1c.index)
    pos_f = pos_full.reindex(idx).fillna(0)
    f1c_f = f1c.reindex(idx)
    T     = len(idx)
    rows  = []
    oos_s = IS_W
    while oos_s < T:
        oos_e    = min(oos_s + OOS_W, T)
        if (oos_e - oos_s) < OOS_W // 2: break
        oos_dt   = idx[oos_s:oos_e]
        yr       = str(oos_dt[0].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
        pos_oos  = pos_f.iloc[oos_s:oos_e]
        f1c_oos  = f1c_f.iloc[oos_s:oos_e]
        pnl      = pos_oos * f1c_oos.diff()
        with np.errstate(invalid="ignore", divide="ignore"):
            ret  = (pnl / f1c_oos.shift(1)).replace([np.inf, -np.inf], np.nan)
        sh = sharpe(ret)
        rows.append({"OOS_yr": yr, "Sharpe": sh})
        oos_s += OOS_W
    return rows

print("\n--- V1 Fixed F8 5yr ---")
rows_v1_f8 = walkforward_fixed(pos_f8_5yr, "V1 F8 5yr")
sh_vals = [r["Sharpe"] for r in rows_v1_f8 if not np.isnan(r["Sharpe"] or np.nan)]
print(f"  Windows: {len(rows_v1_f8)}  |  Avg OOS Sharpe: {np.nanmean(sh_vals):+.3f}  |  Full-IS Sharpe: {compute_sharpe_on_period(pos_f8_5yr):+.3f}")
print("  " + "  ".join([f"{r['OOS_yr']}: {F(r['Sharpe'])}" for r in rows_v1_f8]))
pos_v = sh_vals
print(f"  Positive OOS windows: {sum(1 for v in pos_v if v > 0)}/{len(pos_v)}")

print("\n--- V1 Fixed F12 5yr (Mark's reference) ---")
rows_v1_f12 = walkforward_fixed(pos_f12_5yr, "V1 F12 5yr")
sh_vals_f12 = [r["Sharpe"] for r in rows_v1_f12 if not np.isnan(r["Sharpe"] or np.nan)]
print(f"  Windows: {len(rows_v1_f12)}  |  Avg OOS Sharpe: {np.nanmean(sh_vals_f12):+.3f}  |  Full-IS Sharpe: {compute_sharpe_on_period(pos_f12_5yr):+.3f}")
print("  " + "  ".join([f"{r['OOS_yr']}: {F(r['Sharpe'])}" for r in rows_v1_f12]))

print("\n--- V2 Fixed BG 10yr (windows from 2016 when signal first valid) ---")
pos_v2_10 = v2_signal(2520)
rows_v2_10 = walkforward_fixed(pos_v2_10, "V2 BG 10yr")
# Only count windows where actual OOS data has valid signal
rows_v2_10_valid = [r for r in rows_v2_10 if r["OOS_yr"] >= "2016"]
sh_vals_v2 = [r["Sharpe"] for r in rows_v2_10_valid if not np.isnan(r["Sharpe"] or np.nan)]
print(f"  Full walk-forward (some early windows may have no 10yr signal):")
print("  " + "  ".join([f"{r['OOS_yr']}: {F(r['Sharpe'])}" for r in rows_v2_10]))
print(f"  Windows from 2016+ (valid signal): {len(rows_v2_10_valid)}  |  Avg OOS: {np.nanmean(sh_vals_v2):+.3f}")

print("\n--- V2 Fixed BG 3yr (testable across all windows) ---")
pos_v2_3 = v2_signal(756)
rows_v2_3 = walkforward_fixed(pos_v2_3, "V2 BG 3yr")
sh_vals_v2_3 = [r["Sharpe"] for r in rows_v2_3 if not np.isnan(r["Sharpe"] or np.nan)]
print(f"  Windows: {len(rows_v2_3)}  |  Avg OOS Sharpe: {np.nanmean(sh_vals_v2_3):+.3f}  |  Full-IS Sharpe: {compute_sharpe_on_period(pos_v2_3):+.3f}")
print("  " + "  ".join([f"{r['OOS_yr']}: {F(r['Sharpe'])}" for r in rows_v2_3]))


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS C: WALK-FORWARD RE-OPTIMIZATION (V1 only)
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*80)
print("ANALYSIS C: WALK-FORWARD RE-OPTIMIZATION — V1 only (IS=5yr, OOS=1yr)")
print("Per IS window: scan F1-F15 x {1yr,3yr,5yr} -> apply IS-optimal params OOS")
print("Lookback capped at IS window length (5yr = 1260d); 7yr/10yr excluded")
print("="*80)

idx_c1 = pos_f8_5yr.index.intersection(f1c.index)
T_c1 = len(idx_c1)
dates_c1 = idx_c1

WF_LOOKBACKS = {"1yr": 252, "3yr": 756, "5yr": 1260}  # capped by IS window

rows_c = []
oos_s = IS_W
while oos_s < T_c1:
    oos_e   = min(oos_s + OOS_W, T_c1)
    if (oos_e - oos_s) < OOS_W // 2: break
    is_s    = oos_s - IS_W
    oos_dt  = dates_c1[oos_s:oos_e]
    yr      = str(oos_dt[0].year) + ("*" if (oos_e - oos_s) < OOS_W else "")
    is_end_dt  = dates_c1[oos_s]
    is_start_dt = dates_c1[is_s]

    # IS scan: find best (k, N) on IS window
    best_is_sh, best_k, best_N_lbl = -np.inf, 8, "5yr"
    for N_lbl, N in WF_LOOKBACKS.items():
        for k in CONTRACTS:
            pos = v1_signal(k, N, curve_px)
            sh  = compute_sharpe_on_period(pos, start=str(is_start_dt.date()), end=str(is_end_dt.date()))
            if not np.isnan(sh) and sh > best_is_sh:
                best_is_sh, best_k, best_N_lbl = sh, k, N_lbl

    # Apply IS-selected params OOS
    pos_oos_full = v1_signal(best_k, WF_LOOKBACKS[best_N_lbl], curve_px)
    sh_oos = compute_sharpe_on_period(pos_oos_full, start=str(oos_dt[0].date()),
                                      end=str(oos_dt[-1].date()) if oos_e < T_c1 else None)

    # Fixed F8 5yr OOS for this same window (benchmark)
    sh_oos_f8 = compute_sharpe_on_period(pos_f8_5yr, start=str(oos_dt[0].date()),
                                          end=str(oos_dt[-1].date()) if oos_e < T_c1 else None)

    rows_c.append({
        "OOS_yr": yr,
        "IS_best_k": best_k,
        "IS_best_lb": best_N_lbl,
        "IS_Sharpe":  best_is_sh,
        "OOS_ReOpt":  sh_oos,
        "OOS_F8_5yr": sh_oos_f8,
    })
    oos_s += OOS_W

print(f"\n{'OOS_yr':8s}  {'IS-Best':10s}  {'IS_Sh':8s}  {'OOS_ReOpt':10s}  {'OOS_F8_5yr':10s}")
print("-"*55)
for r in rows_c:
    print(f"  {r['OOS_yr']:6s}  F{r['IS_best_k']:2d}/{r['IS_best_lb']:4s}  "
          f"{F(r['IS_Sharpe'])}  {F(r['OOS_ReOpt'])}    {F(r['OOS_F8_5yr'])}")

reopt_sh = [r["OOS_ReOpt"]  for r in rows_c if not np.isnan(r["OOS_ReOpt"]  or np.nan)]
fixed_sh = [r["OOS_F8_5yr"] for r in rows_c if not np.isnan(r["OOS_F8_5yr"] or np.nan)]
print(f"\n  Avg OOS Sharpe — Re-Optimized per window : {np.nanmean(reopt_sh):+.3f}")
print(f"  Avg OOS Sharpe — Fixed F8 5yr            : {np.nanmean(fixed_sh):+.3f}")
print(f"  Positive OOS windows — Re-Opt  : {sum(1 for v in reopt_sh if v > 0)}/{len(reopt_sh)}")
print(f"  Positive OOS windows — Fixed   : {sum(1 for v in fixed_sh if v > 0)}/{len(fixed_sh)}")

# Which contract/lookback was selected most often?
from collections import Counter
sel_counts = Counter(f"F{r['IS_best_k']}/{r['IS_best_lb']}" for r in rows_c)
print(f"\n  IS-selected param frequency across windows:")
for combo, cnt in sel_counts.most_common():
    print(f"    {combo}: {cnt} windows")


# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*80)
print("SUMMARY — OOS REALITY CHECK")
print("="*80)
print(f"\nV1 MA Reversion:")
print(f"  Full-IS Sharpe (F8  5yr): {F(compute_sharpe_on_period(pos_f8_5yr))}")
print(f"  Full-IS Sharpe (F12 5yr): {F(compute_sharpe_on_period(pos_f12_5yr))}")
print(f"  Hard-Split OOS (F8  5yr, 2016-2025): {F(sh_f8_oos)}")
print(f"  Hard-Split OOS (F12 5yr, 2016-2025): {F(sh_f12_oos)}")
print(f"  Walk-Forward fixed F8  5yr (avg OOS): {F(np.nanmean([r['Sharpe'] for r in rows_v1_f8 if not np.isnan(r['Sharpe'] or np.nan)]))}")
print(f"  Walk-Forward fixed F12 5yr (avg OOS): {F(np.nanmean([r['Sharpe'] for r in rows_v1_f12 if not np.isnan(r['Sharpe'] or np.nan)]))}")
print(f"  Walk-Forward re-opt (avg OOS)        : {F(np.nanmean(reopt_sh))}")

print(f"\nV2 Baz-Granger Reversal:")
print(f"  Full-IS Sharpe (BG 10yr): {F(compute_sharpe_on_period(pos_v2_10))}  [signal only valid from Jan 2016]")
print(f"  Full-IS Sharpe (BG  3yr): {F(compute_sharpe_on_period(pos_v2_3))}")
print(f"  Hard-Split OOS 2016-2025 (BG 10yr): {F(compute_sharpe_on_period(pos_v2_10, start=IS_END_A))}  <- same as full-IS (no IS period)")
print(f"  Hard-Split OOS 2016-2025 (BG  3yr): {F(compute_sharpe_on_period(pos_v2_3,  start=IS_END_A))}")
print(f"  Walk-Forward avg OOS (BG 10yr, 2016+): {F(np.nanmean(sh_vals_v2))}")
print(f"  Walk-Forward avg OOS (BG  3yr, all) : {F(np.nanmean(sh_vals_v2_3))}")

print("\nDone.")
