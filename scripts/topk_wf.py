"""
topk_wf.py
Walk-forward: 5yr IS window ranks all 7875 MA pairs by Sharpe,
picks top K, equal-weights them OOS for 1yr, rolls forward.
K = 3 and K = 5.
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
K_LIST = [3, 5]

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
print(f"{NP} pairs  |  {len(K_LIST)} portfolio variants (K={K_LIST})", flush=True)

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

    # IS scan — one pass for all pairs
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

    # IS-best single pair (WF-standard)
    best_i   = int(np.nanargmax(all_sh)) if not np.all(np.isnan(all_sh)) else 0
    best_m, best_n = pairs[best_i]
    wf_pos   = L1(np.sign(f1r_w.rolling(best_m).mean() - f1r_w.rolling(best_n).mean()))
    wf_sh    = sharpe_ser(wf_pos.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)

    row = {
        "OOS":          oos_yr,
        "IS window":    is_lbl,
        "WF best pair": f"({best_m},{best_n})",
        "WF OOS":       F(wf_sh),
        "MA(35,43)":    F(ma35_sh),
    }

    # Top-K portfolios
    valid_idx = np.where(np.isfinite(all_sh))[0]
    sorted_idx = valid_idx[np.argsort(all_sh[valid_idx])[::-1]]  # descending

    for K in K_LIST:
        top_k = sorted_idx[:K]
        selected = [pairs[i] for i in top_k]
        sel_spans = [spans[i] for i in top_k]
        sel_sh_is = [all_sh[i] for i in top_k]

        combined = pd.Series(0.0, index=f1r_w.index)
        for pm, pn in selected:
            combined += L1(np.sign(f1r_w.rolling(pm).mean() - f1r_w.rolling(pn).mean()))
        combined /= K

        comb_oos = combined.iloc[-oos_len:].set_axis(oos_dates)
        comb_sh  = sharpe_ser(comb_oos, f1c_oos)

        pair_str  = " / ".join([f"({p[0]},{p[1]})" for p in selected])
        span_str  = " / ".join([str(s) for s in sel_spans])
        is_sh_str = " / ".join([f"{s:.2f}" for s in sel_sh_is])

        row[f"Top{K} OOS Sharpe"] = F(comb_sh)
        row[f"Top{K} pairs"]      = pair_str
        row[f"Top{K} spans"]      = span_str
        row[f"Top{K} IS Sharpes"] = is_sh_str

    rows.append(row)
    print(f"  {oos_yr} ({is_lbl}): "
          f"WF={row['WF OOS']}  MA35={row['MA(35,43)']}  "
          f"Top3={row.get('Top3 OOS Sharpe','—')}  Top5={row.get('Top5 OOS Sharpe','—')}", flush=True)
    oos_s += OOS

df = pd.DataFrame(rows)

pd.set_option("display.width", 220)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

# ── Summary table ──────────────────────────────────────────────────────────────
print("\n" + "="*160)
print("TOP-K MA PORTFOLIO — Walk-Forward  (IS=5yr, OOS=1yr, Lag-1, EW)")
print("="*160)
cols_compact = ["OOS", "IS window", "WF OOS", "MA(35,43)",
                "Top3 OOS Sharpe", "Top5 OOS Sharpe"]
print(df[cols_compact].to_string(index=False))

def avg_col(col):
    vals = []
    for v in df[col].dropna():
        s = str(v).strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print(f"\nAVERAGE OOS SHARPE  (~2010–2024):")
print(f"  WF best single pair : {avg_col('WF OOS'):+.3f}   (re-optimizes every year → overfitting)")
print(f"  MA(35,43) fixed     : {avg_col('MA(35,43)'):+.3f}   (fixed parameter, no optimization)")
print(f"  Top-3 EW portfolio  : {avg_col('Top3 OOS Sharpe'):+.3f}")
print(f"  Top-5 EW portfolio  : {avg_col('Top5 OOS Sharpe'):+.3f}")

# ── Detail: which pairs were selected ─────────────────────────────────────────
print("\n" + "="*160)
print("TOP-3 SELECTED PAIRS  (IS Sharpes in brackets) + average span")
print("="*160)
for _, r in df.iterrows():
    if r["OOS"] == "2025*": continue
    print(f"  {r['OOS']} ({r['IS window']}): {r.get('Top3 pairs','—')}")
    print(f"         spans: {r.get('Top3 spans','—')}   IS Sharpes: {r.get('Top3 IS Sharpes','—')}")

# ── Post-2022 ──────────────────────────────────────────────────────────────────
post22 = df[df["OOS"].str.replace("*", "").str.strip().astype(int) >= 2022].copy()
print("\n" + "="*80)
print("POST-2022 COMPARISON")
print("="*80)
print(post22[["OOS", "IS window", "MA(35,43)", "Top3 OOS Sharpe", "Top5 OOS Sharpe"]].to_string(index=False))
print(f"\nAvg OOS Sharpe 2022+:")
print(f"  MA(35,43) : {avg_col('MA(35,43)') if True else '':+.3f}")

# recompute post22 avgs
for col in ["MA(35,43)", "Top3 OOS Sharpe", "Top5 OOS Sharpe"]:
    vals = []
    for v in post22[col]:
        s = str(v).strip().lstrip("+")
        try: vals.append(float(s))
        except: pass
    label = col.replace(" OOS Sharpe","").replace("MA(35,43)","MA(35,43)")
    print(f"  {label:20s}: {np.nanmean(vals):+.3f}" if vals else f"  {label}: —")
