"""
anchor_vs_minspan.py
Compares three approaches on the same walk-forward windows (IS=5yr, OOS=1yr):

  1. MA(35,43) fixed         — single fixed parameter, no optimization
  2. Fixed anchors EW        — MA(10,25) + MA(35,43) + MA(63,100), zero optimization
  3. Min-span F/M/S EW       — IS picks best per bucket with minimum span constraint:
                               Fast   n≤30  span≥5
                               Medium 31≤n≤60  span≥8
                               Slow   n>60   span≥15
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

# Fixed anchors
ANCHORS = [(10, 25), (35, 43), (63, 100)]

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
m_arr = np.array([p[0] for p in pairs])
n_arr = np.array([p[1] for p in pairs])
spans = n_arr - m_arr
m_idx = m_arr - 1
n_idx = n_arr - 1
NP    = len(pairs)

# ── Min-span bucket masks ──────────────────────────────────────────────────────
MS_BUCKETS = {
    "Fast   (n≤30,  span≥5) ": (n_arr <= 30)              & (spans >= 5),
    "Medium (31≤n≤60,span≥8)": (n_arr >= 31) & (n_arr <= 60) & (spans >= 8),
    "Slow   (n>60,  span≥15)": (n_arr > 60)               & (spans >= 15),
}
MS_KEYS = list(MS_BUCKETS.keys())

print("\nMin-span buckets:")
for name, mask in MS_BUCKETS.items():
    ms = m_arr[mask]; ns = n_arr[mask]; ss = spans[mask]
    print(f"  {name}: {mask.sum():4d} pairs  "
          f"span=[{ss.min()}–{ss.max()}]  "
          f"e.g. MA({ms[len(ms)//2]},{ns[len(ns)//2]})")

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

    f1r_w   = f1r.iloc[is_s:oos_e]
    f1c_oos = f1c.reindex(oos_dates)

    # ── 1. MA(35,43) fixed ─────────────────────────────────────────────────────
    ma35_pos = L1(np.sign(f1r_w.rolling(35).mean() - f1r_w.rolling(43).mean()))
    ma35_sh  = sharpe_ser(ma35_pos.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)

    # ── 2. Fixed anchors — no IS scan needed ───────────────────────────────────
    anc_combined = pd.Series(0.0, index=f1r_w.index)
    for am, an in ANCHORS:
        anc_combined += L1(np.sign(f1r_w.rolling(am).mean() - f1r_w.rolling(an).mean()))
    anc_combined /= len(ANCHORS)
    anc_sh = sharpe_ser(anc_combined.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)

    # ── 3. Min-span IS optimisation ────────────────────────────────────────────
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

    ms_selected = {}
    ms_combined = pd.Series(0.0, index=f1r_w.index)
    for bname, bmask in MS_BUCKETS.items():
        b_sh = np.where(bmask, all_sh, np.nan)
        if np.all(np.isnan(b_sh)):
            continue
        bi = int(np.nanargmax(b_sh))
        pm, pn = pairs[bi]
        ms_selected[bname] = (pm, pn, all_sh[bi])
        ms_combined += L1(np.sign(f1r_w.rolling(pm).mean() - f1r_w.rolling(pn).mean()))
    ms_combined /= len(ms_selected)

    ms_oos = ms_combined.iloc[-oos_len:].set_axis(oos_dates)
    ms_sh  = sharpe_ser(ms_oos, f1c_oos)

    def ps(bk):
        if bk not in ms_selected: return "—"
        pm, pn, sh = ms_selected[bk]
        return f"MA({pm},{pn}) sp={pn-pm} IS={sh:.2f}"

    row = {
        "OOS":           oos_yr,
        "IS window":     is_lbl,
        "MA(35,43)":     F(ma35_sh),
        "Fixed anchors": F(anc_sh),
        "Min-span F/M/S":F(ms_sh),
        "F picked":      ps(MS_KEYS[0]),
        "M picked":      ps(MS_KEYS[1]),
        "S picked":      ps(MS_KEYS[2]),
    }
    rows.append(row)
    print(f"  {oos_yr} ({is_lbl}):  MA35={row['MA(35,43)']}  "
          f"Anchors={row['Fixed anchors']}  MinSpan={row['Min-span F/M/S']}  "
          f"| F:{row['F picked']}  M:{row['M picked']}  S:{row['S picked']}", flush=True)
    oos_s += OOS

df = pd.DataFrame(rows)

pd.set_option("display.width", 240)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

print("\n" + "="*180)
print("FIXED ANCHORS vs MIN-SPAN F/M/S vs MA(35,43)  (IS=5yr, OOS=1yr, Lag-1)")
print("Fixed anchors: MA(10,25) + MA(35,43) + MA(63,100), equal-weight, zero optimization")
print("Min-span:      IS picks best in Fast(n≤30,sp≥5) / Med(31≤n≤60,sp≥8) / Slow(n>60,sp≥15)")
print("="*180)
cols = ["OOS","IS window","MA(35,43)","Fixed anchors","Min-span F/M/S",
        "F picked","M picked","S picked"]
print(df[cols].to_string(index=False))

def avg_col(col):
    vals = []
    for v in df[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print(f"\nAVERAGE OOS SHARPE  (~2010–2024):")
print(f"  MA(35,43) fixed    : {avg_col('MA(35,43)'):+.3f}")
print(f"  Fixed anchors EW   : {avg_col('Fixed anchors'):+.3f}")
print(f"  Min-span F/M/S EW  : {avg_col('Min-span F/M/S'):+.3f}")

post22 = df[df["OOS"].str.replace("*","").str.strip().astype(int) >= 2022].copy()
print("\n" + "="*90)
print("POST-2022")
print("="*90)
print(post22[["OOS","IS window","MA(35,43)","Fixed anchors","Min-span F/M/S",
              "F picked","M picked","S picked"]].to_string(index=False))
print()
for col in ["MA(35,43)", "Fixed anchors", "Min-span F/M/S"]:
    vals = []
    for v in post22[col]:
        s = str(v).split()[0].strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    print(f"  {col:20s} avg 2022+: {np.nanmean(vals):+.3f}" if vals else f"  {col}: —")

# Per-year correlation between Fixed anchors and MA(35,43)
print("\n" + "="*60)
print("YEAR-BY-YEAR: Fixed anchors vs MA(35,43)")
print("="*60)
for _, r in df.iterrows():
    if r["OOS"] == "2025*": continue
    ma  = str(r["MA(35,43)"]).lstrip("+")
    anc = str(r["Fixed anchors"]).lstrip("+")
    try:
        diff = float(anc) - float(ma)
        flag = " ▲" if diff > 0.1 else (" ▼" if diff < -0.1 else "  ")
        print(f"  {r['OOS']}:  MA35={r['MA(35,43)']}  Anchors={r['Fixed anchors']}  diff={diff:+.3f}{flag}")
    except:
        pass
