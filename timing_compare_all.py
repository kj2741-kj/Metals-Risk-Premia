"""
timing_compare_all.py
=====================
For every variant across all 3 strategy tabs (Momentum, Carry, Value) and all
versions, compute BOTH timing conventions and report Sharpe / MaxDD / AnnRet.

Timing convention (project standard, post look-ahead fix — matches app.py _carry_cum_pnl):
  Same-Day = position(t) = sign(signal(t-1))  -> sig.shift(1)  [trade at signal close, earn t->t+1]
  Lag-1    = position(t) = sign(signal(t-2))  -> sig.shift(2)  [trade next close, earn t+1->t+2]

Output: timing_compare_all.csv (full matrix) + condensed summary printed.
All metrics are GROSS plus a Net-5bps Sharpe column.
"""
import os, sys, io, warnings
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np, pandas as pd

BASE = r"C:\Users\Kartavya\Metals Risk Premia"

# ── Load data (same as generate_tradebooks.py) ────────────────────────────────
f1_df = pd.read_csv(f"{BASE}/LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()

xls = pd.ExcelFile(f"{BASE}/Metals Futures Curve.csv")
cu_sh = next(s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower())
df_raw = pd.read_excel(xls, sheet_name=cu_sh, header=None, nrows=5)
hr = [i for i in range(min(4, len(df_raw)))
      if any(kw in " ".join(str(v).lower() for v in df_raw.iloc[i].values if pd.notna(v))
             for kw in ["date","f1","f2","price"])]
df_c = pd.read_excel(xls, sheet_name=cu_sh, header=hr if hr else [0,1])
if isinstance(df_c.columns, pd.MultiIndex):
    df_c.columns = ["_".join(str(p) for p in t if pd.notna(p) and "Unnamed" not in str(p) and str(p).strip()) or str(t)
                    for t in df_c.columns]
df_c.columns = [str(c).strip() for c in df_c.columns]
dc = [c for c in df_c.columns if "date" in c.lower()]
if dc: df_c = df_c.rename(columns={dc[0]: "Date"})
df_c["Date"] = pd.to_datetime(df_c["Date"], errors="coerce")
df_c = df_c.dropna(subset=["Date"]).set_index("Date").sort_index()
df_c.index = df_c.index.normalize()
crv = {}
for col in df_c.columns:
    cl = col.lower().replace(" ", "_")
    for i in range(1, 27):
        if f"f{i}_" in cl and "price" in cl:
            crv[f"F{i}"] = pd.to_numeric(df_c[col], errors="coerce"); break
crv = pd.DataFrame(crv, index=df_c.index)

cash_cu = pd.DataFrame()
try:
    xls_cash = pd.ExcelFile(f"{BASE}/Metals Cash and 3M.xlsx")
    cu_cs = next((s for s in xls_cash.sheet_names if "copper" in s.lower()), None)
    if cu_cs:
        df_cash = pd.read_excel(xls_cash, sheet_name=cu_cs, header=[0,1,2])
        nc = []
        for c in df_cash.columns:
            parts = [str(p).strip() for p in c if str(p).strip() and "Unnamed" not in str(p)]
            nc.append("_".join(parts) if parts else str(c))
        df_cash.columns = nc
        dc2 = [c for c in df_cash.columns if "date" in c.lower()]
        if dc2: df_cash = df_cash.rename(columns={dc2[0]: "Date"})
        df_cash["Date"] = pd.to_datetime(df_cash["Date"], errors="coerce")
        df_cash = df_cash.dropna(subset=["Date"]).set_index("Date").sort_index()
        df_cash.index = df_cash.index.normalize()
        res = pd.DataFrame(index=df_cash.index)
        for col in df_cash.columns:
            cl = col.lower()
            if "spread" not in cl and "cash" in cl and "price" in cl and "cash_price" not in res.columns:
                res["cash_price"] = pd.to_numeric(df_cash[col], errors="coerce")
            elif "spread" not in cl and "3m" in cl and "price" in cl and "3m_price" not in res.columns:
                res["3m_price"] = pd.to_numeric(df_cash[col], errors="coerce")
        cash_cu = res
except Exception as e:
    print(f"Cash load failed: {e}")

# ── Metric engine ─────────────────────────────────────────────────────────────
def metrics(sig: pd.Series, same_day: bool, tc_bps=5):
    """sig = raw discrete signal (+1/0/-1 or fractional). Returns (sharpe, net5_sharpe, annret%, maxdd_usd)."""
    sig = sig.replace([np.inf, -np.inf], np.nan)
    pos = (sig.shift(1) if same_day else sig.shift(2)).fillna(0)
    idx = pos.index.intersection(f1c.index)
    if len(idx) < 30: return (np.nan, np.nan, np.nan, np.nan)
    pos = pos.reindex(idx); f = f1c.reindex(idx)
    gp = pos * f.diff()
    chg = pos.diff().abs(); chg.iloc[0] = abs(pos.iloc[0])
    tc = chg * (tc_bps/10000.0/2.0) * f
    npnl = gp - tc
    with np.errstate(invalid="ignore", divide="ignore"):
        gr = (gp / f.shift(1)).replace([np.inf,-np.inf], np.nan)
        nr = (npnl / f.shift(1)).replace([np.inf,-np.inf], np.nan)
    act_g = gr[pos != 0].dropna(); act_n = nr[pos != 0].dropna()
    sh  = float(act_g.mean()/act_g.std(ddof=1)*np.sqrt(252)) if len(act_g) > 20 else np.nan
    nsh = float(act_n.mean()/act_n.std(ddof=1)*np.sqrt(252)) if len(act_n) > 20 else np.nan
    ar  = float(gr.dropna().mean()*252*100)
    cum = gp.cumsum(); dd = float((cum - cum.cummax()).min())
    return (sh, nsh, ar, dd)

def ew(s, com): return s.ewm(com=com, adjust=False).mean()

# ── Build raw signal series per variant ───────────────────────────────────────
def ma_sig(s, l): return np.sign(f1r.rolling(s).mean() - f1r.rolling(l).mean())

def cta_sig():
    pv = f1r.rolling(63).std(); us = []
    for sk, lk in [(8,24),(16,48),(32,96)]:
        x = ew(f1r, sk-1) - ew(f1r, lk-1); y = x/pv
        z = y / y.rolling(252).std()
        with np.errstate(invalid="ignore"):
            u = z*np.exp(-z**2/4)/0.89
        us.append(u.values)
    return pd.Series(np.sign(np.nanmean(np.stack(us, axis=1), axis=1)), index=f1r.index)

def anchors_ew_sig():
    return (ma_sig(10,25).fillna(0) + ma_sig(35,43).fillna(0) + ma_sig(63,100).fillna(0)) / 3.0

def carry_v1(kind):
    if kind == "f1f2": raw = (crv["F1"]-crv["F2"])/crv["F1"]
    elif kind == "f1f3": raw = (crv["F1"]-crv["F3"])/crv["F1"]
    else:
        cp = cash_cu["cash_price"]; tm = cash_cu["3m_price"]
        i = cp.dropna().index.intersection(tm.dropna().index)
        raw = (cp.reindex(i)-tm.reindex(i))/cp.reindex(i)
    return np.sign(raw.replace([np.inf,-np.inf], np.nan))

def carry_v2(j, k): return np.sign(((crv[f"F{j}"]-crv[f"F{k}"])/crv[f"F{k}"]).replace([np.inf,-np.inf], np.nan))

def carry_v3():
    base = ((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf], np.nan)
    z = (base - base.rolling(252).mean())/base.rolling(252).std()
    return np.sign(z.replace([np.inf,-np.inf], np.nan))

def carry_v4(h):
    base = ((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf], np.nan)
    return np.sign((base - base.shift(h)).replace([np.inf,-np.inf], np.nan))

def value_v1(k, N, thr=0.10):
    col = f"F{k}"
    if col not in crv.columns: return None
    p = crv[col].dropna(); ma = p.rolling(N, min_periods=N//2).mean()
    dev = ((p-ma)/ma).replace([np.inf,-np.inf], np.nan)
    return pd.Series(np.where(dev < -thr, 1.0, np.where(dev > thr, -1.0, 0.0)), index=dev.index)

def value_v2(N): return np.sign((f1r.shift(N)-f1r).replace([np.inf,-np.inf], np.nan))

LB = {"1yr":252, "3yr":756, "5yr":1260, "7yr":1764, "10yr":2520}

# ── Enumerate every variant ───────────────────────────────────────────────────
rows = []
def add(tab, version, variant, sig):
    if sig is None or (hasattr(sig, "empty") and sig.empty): return
    for tname, sd in [("Same-Day", True), ("Lag-1", False)]:
        sh, nsh, ar, dd = metrics(sig, sd)
        rows.append(dict(Tab=tab, Version=version, Variant=variant, Timing=tname,
                         Sharpe=round(sh,3) if sh==sh else None,
                         Net5_Sharpe=round(nsh,3) if nsh==nsh else None,
                         AnnRet_pct=round(ar,2) if ar==ar else None,
                         MaxDD_USD=round(dd,1) if dd==dd else None))

# MOMENTUM
add("Momentum", "MA Crossover", "MA(35,43)", ma_sig(35,43))
add("Momentum", "MA Crossover", "MA(10,25)", ma_sig(10,25))
add("Momentum", "MA Crossover", "MA(63,100)", ma_sig(63,100))
add("Momentum", "CTA", "Baz-Granger 3-timescale", cta_sig())
add("Momentum", "Anchors", "Anchors EW (10/25+35/43+63/100)", anchors_ew_sig())

# CARRY
add("Carry", "V1 Roll Yield", "(F1-F2)/F1", carry_v1("f1f2"))
add("Carry", "V1 Roll Yield", "(F1-F3)/F1", carry_v1("f1f3"))
if not cash_cu.empty:
    add("Carry", "V1 Roll Yield", "(Cash-3M)/Cash", carry_v1("cash3m"))
for j,k in [(3,15),(4,16),(5,17),(6,18),(7,19),(8,20),(9,21),(10,22),(11,23),(12,24)]:
    if f"F{j}" in crv.columns and f"F{k}" in crv.columns:
        add("Carry", "V2 Long Slope", f"F{j}-F{k}", carry_v2(j,k))
add("Carry", "V3 Z-score", "Z-score 252d", carry_v3())
add("Carry", "V4 Carry Mom", "20-day Δcarry", carry_v4(20))
add("Carry", "V4 Carry Mom", "60-day Δcarry", carry_v4(60))

# VALUE
for k in range(1, 16):
    for lb, N in LB.items():
        add("Value", f"V1 MA Reversion {lb}", f"F{k}", value_v1(k, N))
for lb, N in LB.items():
    add("Value", f"V2 BG Reversal", lb, value_v2(N))

df = pd.DataFrame(rows)
out = f"{BASE}/timing_compare_all.csv"
df.to_csv(out, index=False)
print(f"Wrote {out}  ({len(df)} rows, {len(df)//2} variants × 2 timings)\n")

# ── Condensed summary: headline variants + best-per-group ─────────────────────
pd.set_option("display.width", 200, "display.max_columns", 20, "display.max_rows", 300)

def show(title, sub):
    print(f"\n{'='*78}\n{title}\n{'='*78}")
    print(sub.to_string(index=False))

# Momentum + Carry: show all
show("MOMENTUM — all variants, both timings",
     df[df.Tab=="Momentum"][["Version","Variant","Timing","Sharpe","Net5_Sharpe","AnnRet_pct","MaxDD_USD"]])
show("CARRY — all variants, both timings",
     df[df.Tab=="Carry"][["Version","Variant","Timing","Sharpe","Net5_Sharpe","AnnRet_pct","MaxDD_USD"]])

# Value: best contract per (version, timing) by Sharpe + the F8/F12 references
val = df[df.Tab=="Value"].copy()
best = (val.sort_values("Sharpe", ascending=False)
           .groupby(["Version","Timing"], as_index=False).first())
show("VALUE — best contract per lookback × timing (by gross Sharpe)",
     best.sort_values(["Version","Timing"])[["Version","Variant","Timing","Sharpe","Net5_Sharpe","AnnRet_pct","MaxDD_USD"]])
show("VALUE — references F8 & F12 (5yr) both timings",
     val[(val.Version=="V1 MA Reversion 5yr") & (val.Variant.isin(["F8","F12"]))]
        [["Version","Variant","Timing","Sharpe","Net5_Sharpe","AnnRet_pct","MaxDD_USD"]])
