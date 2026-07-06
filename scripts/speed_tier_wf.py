"""
speed_tier_wf.py
Walk-forward: 5yr IS window picks best MA pair per speed tier (bucketed by span n-m),
equal-weights K tiers, evaluates OOS for 1yr, rolls forward.
Also reports MA(35,43) fixed OOS and post-2022 slice.
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
TIER_LIST = [3, 5]

# ── Helpers ────────────────────────────────────────────────────────────────────
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
print(f"Pre-computing {MAX_N} SMA series...", flush=True)
sma = np.column_stack([f1r.rolling(k + 1).mean().values for k in range(MAX_N)])

delta_np = np.empty(T); delta_np[0] = np.nan; delta_np[1:] = np.diff(f1c_np)
prev_np  = np.empty(T); prev_np[0]  = np.nan; prev_np[1:]  = f1c_np[:-1]

pairs  = [(m + 1, n + 1) for m in range(MAX_N) for n in range(m + 1, MAX_N)]
spans  = np.array([p[1] - p[0] for p in pairs])
m_idx  = np.array([p[0] - 1 for p in pairs])
n_idx  = np.array([p[1] - 1 for p in pairs])
NP     = len(pairs)
print(f"{NP} pairs | batches of {BATCH}")

# Bucket assignment: divide span range evenly into K buckets
def make_buckets(K, spans, MAX_N):
    max_span = MAX_N - 1  # 125
    return np.minimum((spans.astype(float) * K / (max_span + 1)).astype(int), K - 1)

bucket_maps = {K: make_buckets(K, spans, MAX_N) for K in TIER_LIST}

# Print bucket boundaries for reference
for K in TIER_LIST:
    bk = bucket_maps[K]
    print(f"\nK={K} speed buckets (by span n-m):")
    for k in range(K):
        mask = bk == k
        lo, hi = spans[mask].min(), spans[mask].max()
        print(f"  Tier {k+1}: span {lo:3d}–{hi:3d}  ({mask.sum()} pairs)")

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

    # ── IS scan: compute Sharpe for all pairs ──────────────────────────────────
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
        mu   = np.nanmean(ret, axis=0)
        sd   = np.nanstd(ret,  axis=0, ddof=1)
        n_a  = np.sum(act, axis=0)
        all_sh[b0:b1] = np.where((sd > 0) & (n_a >= 20), mu / sd * np.sqrt(252), np.nan)

    # ── OOS base data ──────────────────────────────────────────────────────────
    f1r_w   = f1r.iloc[is_s:oos_e]
    f1c_oos = f1c.reindex(oos_dates)

    # MA(35,43) fixed
    ma35_pos = L1(np.sign(f1r_w.rolling(35).mean() - f1r_w.rolling(43).mean()))
    ma35_sh  = sharpe_ser(ma35_pos.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)

    row = {"OOS": oos_yr, "IS window": is_lbl, "MA(35,43)": F(ma35_sh)}

    # ── Speed-tier portfolios ──────────────────────────────────────────────────
    for K in TIER_LIST:
        bk = bucket_maps[K]
        selected = []
        for k in range(K):
            mask  = bk == k
            b_sh  = np.where(mask, all_sh, np.nan)
            if np.all(np.isnan(b_sh)):
                continue
            best_i = int(np.nanargmax(b_sh))
            selected.append(pairs[best_i])

        combined = pd.Series(0.0, index=f1r_w.index)
        for pm, pn in selected:
            combined += L1(np.sign(f1r_w.rolling(pm).mean() - f1r_w.rolling(pn).mean()))
        combined /= len(selected)

        comb_oos = combined.iloc[-oos_len:].set_axis(oos_dates)
        comb_sh  = sharpe_ser(comb_oos, f1c_oos)

        pair_str = " / ".join([f"({p[0]},{p[1]})" for p in selected])
        row[f"K={K} Sharpe"] = F(comb_sh)
        row[f"K={K} pairs"]  = pair_str

    rows.append(row)
    print(f"  {oos_yr} ({is_lbl}):  MA35={row['MA(35,43)']}  "
          f"K3={row.get('K=3 Sharpe','—')}  K5={row.get('K=5 Sharpe','—')}", flush=True)
    oos_s += OOS

df = pd.DataFrame(rows)

# ── Print main table ───────────────────────────────────────────────────────────
pd.set_option("display.width", 220); pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

cols_main = ["OOS", "IS window", "MA(35,43)", "K=3 Sharpe", "K=3 pairs", "K=5 Sharpe", "K=5 pairs"]
print("\n" + "="*180)
print("SPEED-TIER PORTFOLIO — Walk-Forward  (IS=5yr rolling, OOS=1yr, Lag-1)")
print("Tier selection: best IS Sharpe within each span bucket. EW combined in OOS.")
print("="*180)
print(df[cols_main].to_string(index=False))

def avg_col(col):
    vals = []
    for v in df[col]:
        try: vals.append(float(str(v).strip().replace("—","").replace("+","")))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print(f"\nAVERAGE OOS SHARPE  (all windows ~2011–2025):")
print(f"  MA(35,43) fixed  : {avg_col('MA(35,43)'):+.3f}")
print(f"  K=3 speed tiers  : {avg_col('K=3 Sharpe'):+.3f}")
print(f"  K=5 speed tiers  : {avg_col('K=5 Sharpe'):+.3f}")

# ── Post-2022 slice ────────────────────────────────────────────────────────────
post22 = df[df["OOS"].str.replace("*", "").str.strip().astype(int) >= 2022].copy()
print("\n" + "="*80)
print("MA(35,43) WALK-FORWARD — 2022 onwards only")
print("="*80)
print(post22[["OOS", "IS window", "MA(35,43)"]].to_string(index=False))

p22_vals = []
for v in post22["MA(35,43)"]:
    try: p22_vals.append(float(str(v).strip().replace("+","")))
    except: pass
print(f"\nMA(35,43) avg OOS Sharpe 2022+: {np.nanmean(p22_vals):+.3f}")
print(f"MA(35,43) avg OOS Sharpe full : {avg_col('MA(35,43)'):+.3f}")
