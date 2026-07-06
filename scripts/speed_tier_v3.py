"""
speed_tier_v3.py
Walk-forward: 5yr IS picks best MA pair per speed bucket (by parameter range),
EW-combines 3 tiers, OOS 1yr, rolls forward.

Buckets by n (slow MA period):
  Fast   : n <= 30            → both m and n <= 30  e.g. MA(5,15), MA(20,30)
  Medium : 31 <= n <= 60      → e.g. MA(10,40), MA(35,43), MA(20,55)
  Slow   : n >  60            → e.g. MA(40,80), MA(63,100), MA(50,120)
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

# ── Pre-compute SMA matrix ─────────────────────────────────────────────────────
T = len(f1r); dates = f1r.index; f1c_np = f1c.values.astype(float)
print("Pre-computing SMA matrix...", flush=True)
sma = np.column_stack([f1r.rolling(k + 1).mean().values for k in range(MAX_N)])

delta_np = np.empty(T); delta_np[0] = np.nan; delta_np[1:] = np.diff(f1c_np)
prev_np  = np.empty(T); prev_np[0]  = np.nan; prev_np[1:]  = f1c_np[:-1]

pairs = [(m + 1, n + 1) for m in range(MAX_N) for n in range(m + 1, MAX_N)]
m_arr = np.array([p[0] for p in pairs])
n_arr = np.array([p[1] for p in pairs])
m_idx = m_arr - 1
n_idx = n_arr - 1
NP    = len(pairs)

# ── Bucket definitions (by n, the slow MA period) ─────────────────────────────
BUCKETS = {
    "Fast   (n ≤ 30)":  n_arr <= 30,
    "Medium (31≤n≤60)": (n_arr >= 31) & (n_arr <= 60),
    "Slow   (n > 60)":  n_arr > 60,
}
BKEYS = list(BUCKETS.keys())

print("\nSpeed buckets (by slow MA period n):")
for name, mask in BUCKETS.items():
    ms = m_arr[mask]; ns = n_arr[mask]
    print(f"  {name}: {mask.sum():4d} pairs  "
          f"m=[{ms.min()}–{ms.max()}]  n=[{ns.min()}–{ns.max()}]")
print(f"  Example pairs — Fast: MA(10,25) MA(20,30) | "
      f"Medium: MA(35,43) MA(20,55) | Slow: MA(40,80) MA(63,100)")

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

    # IS scan — single pass over all pairs
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

    # OOS data
    f1r_w   = f1r.iloc[is_s:oos_e]
    f1c_oos = f1c.reindex(oos_dates)

    # MA(35,43) fixed benchmark
    ma35_pos = L1(np.sign(f1r_w.rolling(35).mean() - f1r_w.rolling(43).mean()))
    ma35_sh  = sharpe_ser(ma35_pos.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)

    # Best from each bucket → equal-weight 3-signal combined position
    selected = {}
    combined = pd.Series(0.0, index=f1r_w.index)
    for bname, bmask in BUCKETS.items():
        b_sh = np.where(bmask, all_sh, np.nan)
        if np.all(np.isnan(b_sh)):
            continue
        bi   = int(np.nanargmax(b_sh))
        pm, pn = pairs[bi]
        selected[bname] = (pm, pn, all_sh[bi])
        combined += L1(np.sign(f1r_w.rolling(pm).mean() - f1r_w.rolling(pn).mean()))
    combined /= len(selected)

    comb_oos = combined.iloc[-oos_len:].set_axis(oos_dates)
    comb_sh  = sharpe_ser(comb_oos, f1c_oos)

    def pair_str(bk):
        if bk not in selected: return "—"
        pm, pn, sh = selected[bk]
        return f"MA({pm},{pn}) IS={sh:.2f}"

    row = {
        "OOS":       oos_yr,
        "IS window": is_lbl,
        "MA(35,43)": F(ma35_sh),
        "F/M/S EW":  F(comb_sh),
        "Fast":      pair_str(BKEYS[0]),
        "Medium":    pair_str(BKEYS[1]),
        "Slow":      pair_str(BKEYS[2]),
    }
    rows.append(row)
    print(f"  {oos_yr} ({is_lbl}):  MA35={row['MA(35,43)']}  FMS={row['F/M/S EW']}  "
          f"| F: {row['Fast']}  M: {row['Medium']}  S: {row['Slow']}", flush=True)
    oos_s += OOS

df = pd.DataFrame(rows)

pd.set_option("display.width", 220)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

print("\n" + "="*160)
print("FAST / MEDIUM / SLOW PORTFOLIO  (IS=5yr, OOS=1yr, Lag-1, EW)")
print("Buckets by n (slow MA period): Fast n≤30 | Medium 31≤n≤60 | Slow n>60")
print("="*160)
print(df.to_string(index=False))

def avg_col(col):
    vals = []
    for v in df[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print(f"\nAVERAGE OOS SHARPE  (~2010–2024):")
print(f"  MA(35,43) fixed : {avg_col('MA(35,43)'):+.3f}")
print(f"  F/M/S EW        : {avg_col('F/M/S EW'):+.3f}")

# Post-2022
post22 = df[df["OOS"].str.replace("*","").str.strip().astype(int) >= 2022].copy()
print("\n" + "="*100)
print("POST-2022")
print("="*100)
print(post22.to_string(index=False))
print()
for col in ["MA(35,43)", "F/M/S EW"]:
    vals = []
    for v in post22[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    print(f"  {col:15s} avg 2022+: {np.nanmean(vals):+.3f}" if vals else f"  {col}: —")
