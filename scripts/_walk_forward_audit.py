"""_walk_forward_audit.py — not in git, for review only"""
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

# ── Load data ──────────────────────────────────────────────────────────────
f1_df = pd.read_csv(f"{BASE}\\LME_Copper_Rolling_F1_v2.csv",
                    parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()
print(f"F1 data: {len(f1r)} rows | {f1r.index[0].date()} → {f1r.index[-1].date()}")

carry_ok = False
try:
    xls  = pd.ExcelFile(f"{BASE}\\Metals Futures Curve.csv")
    df_c = pd.read_excel(xls, sheet_name="Copper LME", header=[0,1,2])
    nc   = []
    for col in df_c.columns:
        parts = [str(p).strip() for p in col
                 if "Unnamed" not in str(p) and str(p).strip()]
        nc.append("_".join(parts) if parts else str(col))
    df_c.columns = nc
    date_col = next(c for c in df_c.columns if "date" in c.lower())
    df_c = (df_c.rename(columns={date_col: "Date"})
               .assign(Date=lambda d: pd.to_datetime(d["Date"], errors="coerce"))
               .dropna(subset=["Date"])
               .set_index("Date")
               .sort_index())
    df_c.index = df_c.index.normalize()
    f2_col = next((c for c in df_c.columns if "f2_price" in c.lower()), None)
    if f2_col:
        f2 = pd.to_numeric(df_c[f2_col], errors="coerce")
        carry_ok = True
        print(f"F2 loaded ({f2_col}): {f2.dropna().count()} values")
    else:
        print(f"F2 col not found. Cols: {[c for c in df_c.columns[:10]]}")
        f2 = None
except Exception as e:
    print(f"Curve load failed: {e}")
    f2 = None

# ── Signal helpers ─────────────────────────────────────────────────────────
def L1(sig):
    return sig.shift(1).fillna(0)

def sharpe_ser(pos, f1c_s):
    idx = pos.index.intersection(f1c_s.index)
    pos = pos.reindex(idx); f1c_ = f1c_s.reindex(idx)
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pos * f1c_.diff() / f1c_.shift(1)).replace([np.inf,-np.inf], np.nan)
    act = ret[pos != 0].dropna()
    if len(act) < 20: return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def cta_sig(f1r_s):
    vol = f1r_s.rolling(63).std()
    us  = []
    for s, l in [(8,24),(16,48),(32,96)]:
        x = (f1r_s.ewm(com=s-1, adjust=False).mean()
           - f1r_s.ewm(com=l-1, adjust=False).mean())
        y = x / vol
        z = y / y.rolling(252).std()
        u = z * np.exp(-z.values**2 / 4) / 0.8579
        us.append(pd.Series(u if isinstance(u, np.ndarray) else u.values,
                             index=f1r_s.index))
    return np.sign(pd.concat(us, axis=1).mean(axis=1))

def carry_pos_series(f1r_s, f2_s):
    idx = f1r_s.index.intersection(f2_s.index)
    raw = (f1r_s.reindex(idx) - f2_s.reindex(idx)) / f1r_s.reindex(idx)
    return np.sign(raw).fillna(0)   # same-day, no L1

# ── Walk-forward (IS=5yr rolling, OOS=1yr) ────────────────────────────────
def walk_forward(f1r, f1c, f2, IS=1260, OOS=252, MAX_N=126):
    T = len(f1r)
    dates = f1r.index
    f1c_np = f1c.values.astype(float)

    print(f"\nPre-computing {MAX_N} SMA series...")
    sma = np.column_stack([f1r.rolling(k+1).mean().values for k in range(MAX_N)])

    delta_np = np.empty(T); delta_np[0] = np.nan; delta_np[1:] = np.diff(f1c_np)
    prev_np  = np.empty(T); prev_np[0]  = np.nan; prev_np[1:]  = f1c_np[:-1]

    pairs  = [(m+1, n+1) for m in range(MAX_N) for n in range(m+1, MAX_N)]
    m_idx  = np.array([p[0]-1 for p in pairs])
    n_idx  = np.array([p[1]-1 for p in pairs])
    BATCH  = 800
    print(f"Scanning {len(pairs)} pairs per window in batches of {BATCH}...")

    rows = []
    oos_s = IS

    while oos_s < T:
        oos_e     = min(oos_s + OOS, T)
        is_s      = oos_s - IS
        oos_dates = dates[oos_s:oos_e]
        oos_len   = len(oos_dates)
        oos_yr    = str(dates[oos_s].year) + ("*" if oos_len < OOS else "")
        is_lbl    = f"{dates[is_s].year}–{dates[oos_s-1].year}"

        # Vectorised IS scan (batched)
        sma_is    = sma[is_s:oos_s]
        d_is      = delta_np[is_s:oos_s]
        p_is      = prev_np[is_s:oos_s]
        base_mask = np.isfinite(d_is) & (p_is > 0)

        best_sh  = -np.inf
        best_m, best_n = 35, 43

        for b0 in range(0, len(pairs), BATCH):
            b1   = min(b0+BATCH, len(pairs))
            mb   = m_idx[b0:b1]; nb = n_idx[b0:b1]
            diff = sma_is[:, mb] - sma_is[:, nb]
            sig  = np.sign(diff)
            pos  = np.zeros_like(sig)
            pos[1:] = np.where(np.isfinite(sig[:-1]), sig[:-1], 0.0)
            act  = (pos != 0) & base_mask[:, None]
            with np.errstate(invalid="ignore", divide="ignore"):
                ret = np.where(act, pos * d_is[:,None] / p_is[:,None], np.nan)
            mu   = np.nanmean(ret, axis=0)
            sd   = np.nanstd(ret,  axis=0, ddof=1)
            n_a  = np.sum(act, axis=0)
            sh   = np.where((sd>0)&(n_a>=20), mu/sd*np.sqrt(252), np.nan)
            bi   = int(np.nanargmax(sh)) if not np.all(np.isnan(sh)) else -1
            if bi >= 0 and sh[bi] > best_sh:
                best_sh = sh[bi]
                best_m, best_n = pairs[b0+bi]

        # OOS evaluation — compute over IS+OOS so warmup is correct, then slice
        f1r_w = f1r.iloc[is_s:oos_e]
        f1c_w = f1c.iloc[is_s:oos_e]

        wf_full   = L1(np.sign(f1r_w.rolling(best_m).mean()
                               - f1r_w.rolling(best_n).mean()))
        ma35_full = L1(np.sign(f1r_w.rolling(35).mean()
                               - f1r_w.rolling(43).mean()))
        cta_full  = L1(cta_sig(f1r_w))

        f1c_oos   = f1c.reindex(oos_dates)
        wf_sh     = sharpe_ser(wf_full.iloc[-oos_len:].set_axis(oos_dates),   f1c_oos)
        ma35_sh   = sharpe_ser(ma35_full.iloc[-oos_len:].set_axis(oos_dates), f1c_oos)
        cta_sh    = sharpe_ser(cta_full.iloc[-oos_len:].set_axis(oos_dates),  f1c_oos)

        if carry_ok:
            idx_c  = oos_dates.intersection(f2.index)
            car_pos = carry_pos_series(f1r.reindex(idx_c), f2.reindex(idx_c)) if len(idx_c)>=20 else None
            car_sh  = sharpe_ser(car_pos, f1c.reindex(idx_c)) if car_pos is not None else np.nan
        else:
            car_sh = np.nan

        # Value V2 10yr — compute from full f1r so shift(2520) has history
        v2_sig  = np.sign(f1r.shift(2520) - f1r)
        v2_pos  = L1(v2_sig).reindex(oos_dates)
        v2_sh   = sharpe_ser(v2_pos, f1c_oos)

        def F(x): return f"{x:+.3f}" if (x is not None and not np.isnan(x)) else "   —"

        rows.append({
            "OOS"          : oos_yr,
            "IS window"    : is_lbl,
            "WF (m,n)"     : f"({best_m},{best_n})",
            "IS Sharpe"    : F(best_sh),
            "WF OOS"       : F(wf_sh),
            "MA(35,43) OOS": F(ma35_sh),
            "CTA Paper OOS": F(cta_sh),
            "Carry V1 OOS" : F(car_sh),
            "Val V2 10y OOS": F(v2_sh),
        })

        print(f"  {oos_yr}: WF ({best_m:3d},{best_n:3d}) IS={F(best_sh)} "
              f"| WF-OOS={F(wf_sh)}  MA35-OOS={F(ma35_sh)}  CTA-OOS={F(cta_sh)}")
        oos_s += OOS

    return pd.DataFrame(rows)


# ── Sub-period table (full dataset signals, sliced by date) ───────────────
def sub_period_table(f1r, f1c, f2, carry_ok):
    ma_p  = L1(np.sign(f1r.rolling(35).mean() - f1r.rolling(43).mean()))
    cta_p = L1(cta_sig(f1r))
    v2_p  = L1(np.sign(f1r.shift(2520) - f1r))
    if carry_ok:
        car_p = carry_pos_series(f1r, f2)
    def F(x): return f"{x:+.3f}" if (x is not None and not np.isnan(x)) else "   —"
    def sh(pos, s=None, e=None, f1c_=None):
        f1c_ = f1c_ if f1c_ is not None else f1c
        p = pos.copy(); c = f1c_.copy()
        if s: p=p[p.index>=s]; c=c[c.index>=s]
        if e: p=p[p.index<=e]; c=c[c.index<=e]
        return sharpe_ser(p, c)
    rows = []
    periods = [
        ("Full (2006–2025)", None, None),
        ("Pre-2022 (2006–2021)", None, "2021-12-31"),
        ("Post-2022 (2022–2025)", "2022-01-01", None),
    ]
    for yr in range(2011, 2026):
        periods.append((str(yr), f"{yr}-01-01", f"{yr}-12-31"))
    for lbl, s, e in periods:
        car = sh(car_p, s, e, f1c.reindex(car_p.index)) if carry_ok else np.nan
        rows.append({
            "Period"      : lbl,
            "MA(35,43)"  : F(sh(ma_p,  s, e)),
            "CTA Paper"  : F(sh(cta_p, s, e)),
            "Carry V1 SD": F(car),
            "Val V2 10yr": F(sh(v2_p,  s, e)),
        })
    return pd.DataFrame(rows)


# ── Run ───────────────────────────────────────────────────────────────────
pd.set_option("display.width", 130)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)

wf = walk_forward(f1r, f1c, f2, IS=1260, OOS=252, MAX_N=126)

print("\n" + "="*115)
print("TABLE 1 — Walk-Forward MA Crossover  (IS=5yr rolling, OOS=1yr)   * = partial year")
print("="*115)
print(wf.to_string(index=False))

def avg(col):
    vals = []
    for v in wf[col]:
        try: vals.append(float(str(v).strip().lstrip("+")))
        except: pass
    return np.nanmean(vals) if vals else np.nan

print(f"\nMEAN ANNUAL OOS SHARPE  (avg over {len(wf)} windows):")
print(f"  Walk-forward best (m,n) each year : {avg('WF OOS'):+.3f}")
print(f"  Fixed MA(35,43)                   : {avg('MA(35,43) OOS'):+.3f}")
print(f"  CTA Paper                         : {avg('CTA Paper OOS'):+.3f}")
print(f"  Carry V1 Same-Day                 : {avg('Carry V1 OOS'):+.3f}")
print(f"  Value V2 10yr                     : {avg('Val V2 10y OOS'):+.3f}")

sp = sub_period_table(f1r, f1c, f2 if carry_ok else None, carry_ok)

print("\n" + "="*75)
print("TABLE 2 — Sub-Period Performance  (full-data signals, active-only Sharpe)")
print("  Carry = Same-Day  |  All others = Lag-1")
print("="*75)
print(sp.to_string(index=False))
print(f"\n* 2025 through {f1r.index[-1].date()}")
