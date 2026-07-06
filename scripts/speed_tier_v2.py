"""
speed_tier_v2.py
Walk-forward: 5yr IS picks best MA pair per speed category, EW-combines, OOS 1yr.
Categories by span (n - m):
  Fast   : span  1 – 30
  Medium : span 31 – 60
  Slow   : span 61 – 125
"""
import warnings, sys, io, os
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import numpy as np
import pandas as pd

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

def L1(sig):
    return sig.shift(1).fillna(0)

def sharpe_ser(pos, f1c_s):
    idx = pos.index.intersection(f1c_s.index)
    pos = pos.reindex(idx); c = f1c_s.reindex(idx)
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pos * c.diff() / c.shift(1)).replace([np.inf, -np.inf], np.nan)
    act = ret[pos != 0].dropna()
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def F(x):
    return f"{x:+.3f}" if (x is not None and not np.isnan(x)) else "  —"

# ── Pre-compute ────────────────────────────────────────────────────────────────
T = len(f1r); dates = f1r.index; f1c_np = f1c.values.astype(float)
print("Pre-computing SMA matrix...", flush=True)
sma = np.column_stack([f1r.rolling(k + 1).mean().values for k in range(MAX_N)])

delta_np = np.empty(T); delta_np[0] = np.nan; delta_np[1:] = np.diff(f1c_np)
prev_np  = np.empty(T); prev_np[0]  = np.nan; prev_np[1:]  = f1c_np[:-1]

pairs = [(m + 1, n + 1) for m in range(MAX_N) for n in range(m + 1, MAX_N)]
m_idx = np.array([p[0] - 1 for p in pairs])
n_idx = np.array([p[1] - 1 for p in pairs])
spans = np.array([p[1] - p[0] for p in pairs])
NP    = len(pairs)

# ── Bucket assignment ──────────────────────────────────────────────────────────
BUCKETS = {
    "Fast   (span  1–30)": (spans >= 1)  & (spans <= 30),
    "Medium (span 31–60)": (spans >= 31) & (spans <= 60),
    "Slow   (span 61+)  ": (spans >= 61),
}
BUCKET_KEYS = list(BUCKETS.keys())

print(f"\nSpeed buckets:")
for name, mask in BUCKETS.items():
    lo, hi = spans[mask].min(), spans[mask].max()
    print(f"  {name}: {mask.sum():4d} pairs  (span {lo}–{hi})")

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

    # IS scan — one pass
    all_sh = np.full(NP, np.nan)
    for b0 in range(0, NP, BATCH):
        b1   = min(b0 + BATCH, NP)
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
        all_sh[b0:b1] = np.where((sd > 0) & (n_a >= 20), mu / sd * np.sqrt(252), np.nan)

    # OOS base
    f1r_w   = f1r.iloc[is_s:oos_e]
    f1c_oos = f1c.reindex(oos_dates)

    # MA(35,43) fixed
    ma35_pos = L1(np.sign(f1r_w.rolling(35).mean() - f1r_w.rolling(43).mean()))
    ma35_sh  = sharpe_ser(ma35_pos.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)

    # Best from each bucket → equal-weight combination
    selected_pairs = []
    bucket_bests   = {}
    for bname, bmask in BUCKETS.items():
        b_sh = np.where(bmask, all_sh, np.nan)
        if np.all(np.isnan(b_sh)):
            continue
        best_i = int(np.nanargmax(b_sh))
        selected_pairs.append(pairs[best_i])
        bucket_bests[bname] = (pairs[best_i], all_sh[best_i])

    combined = pd.Series(0.0, index=f1r_w.index)
    for pm, pn in selected_pairs:
        combined += L1(np.sign(f1r_w.rolling(pm).mean() - f1r_w.rolling(pn).mean()))
    combined /= len(selected_pairs)

    comb_oos = combined.iloc[-oos_len:].set_axis(oos_dates)
    comb_sh  = sharpe_ser(comb_oos, f1c_oos)

    row = {
        "OOS":       oos_yr,
        "IS window": is_lbl,
        "MA(35,43)": F(ma35_sh),
        "F/M/S EW":  F(comb_sh),
        "Fast pair":   f"({bucket_bests.get(BUCKET_KEYS[0], ((0,0),0))[0][0]},{bucket_bests.get(BUCKET_KEYS[0], ((0,0),0))[0][1]}) IS={bucket_bests.get(BUCKET_KEYS[0], ((0,0),0))[1]:.2f}",
        "Med pair":    f"({bucket_bests.get(BUCKET_KEYS[1], ((0,0),0))[0][0]},{bucket_bests.get(BUCKET_KEYS[1], ((0,0),0))[0][1]}) IS={bucket_bests.get(BUCKET_KEYS[1], ((0,0),0))[1]:.2f}",
        "Slow pair":   f"({bucket_bests.get(BUCKET_KEYS[2], ((0,0),0))[0][0]},{bucket_bests.get(BUCKET_KEYS[2], ((0,0),0))[0][1]}) IS={bucket_bests.get(BUCKET_KEYS[2], ((0,0),0))[1]:.2f}",
    }
    rows.append(row)
    print(f"  {oos_yr} ({is_lbl}):  MA35={row['MA(35,43)']}  FMS={row['F/M/S EW']}  "
          f"| F:{row['Fast pair']}  M:{row['Med pair']}  S:{row['Slow pair']}", flush=True)
    oos_s += OOS

df = pd.DataFrame(rows)

pd.set_option("display.width", 220)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

print("\n" + "="*130)
print("FAST / MEDIUM / SLOW TIER PORTFOLIO  (IS=5yr, OOS=1yr, Lag-1, EW)")
print("Best IS pair per bucket → equal-weight 3-signal combined position")
print("="*130)
print(df[["OOS","IS window","MA(35,43)","F/M/S EW","Fast pair","Med pair","Slow pair"]].to_string(index=False))

def avg_col(col):
    vals = []
    for v in df[col].dropna():
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print(f"\nAVERAGE OOS SHARPE  (~2010–2024):")
print(f"  MA(35,43) fixed  : {avg_col('MA(35,43)'):+.3f}")
print(f"  F/M/S EW         : {avg_col('F/M/S EW'):+.3f}")

# Post-2022
post22 = df[df["OOS"].str.replace("*","").str.strip().astype(int) >= 2022].copy()
print("\n" + "="*80)
print("POST-2022")
print("="*80)
print(post22[["OOS","IS window","MA(35,43)","F/M/S EW","Fast pair","Med pair","Slow pair"]].to_string(index=False))

for col in ["MA(35,43)", "F/M/S EW"]:
    vals = []
    for v in post22[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    print(f"  {col:15s} avg 2022+: {np.nanmean(vals):+.3f}" if vals else f"  {col}: —")
