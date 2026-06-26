"""
portfolio_timing_compare.py
===========================
EW portfolio (Mom+Carry+Value)/3 across leg-variant combinations, under BOTH
timing conventions (Same-Day=shift1, Lag-1=shift2). Reports Sharpe/Net5/AnnRet/MaxDD.
Identifies the best-of-three combination.
"""
import os, sys, io, warnings
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np, pandas as pd
from itertools import product

BASE = r"C:\Users\Kartavya\Metals Risk Premia"

f1_df = pd.read_csv(f"{BASE}/LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index(); f1c = f1_df["F1_continuous"].sort_index()

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

# ── Signal builders ───────────────────────────────────────────────────────────
def ew(s, com): return s.ewm(com=com, adjust=False).mean()
def ma_sig(s, l): return np.sign(f1r.rolling(s).mean() - f1r.rolling(l).mean())
def cta_sig():
    pv = f1r.rolling(63).std(); us = []
    for sk, lk in [(8,24),(16,48),(32,96)]:
        x = ew(f1r, sk-1) - ew(f1r, lk-1); y = x/pv; z = y / y.rolling(252).std()
        with np.errstate(invalid="ignore"): us.append((z*np.exp(-z**2/4)/0.89).values)
    return pd.Series(np.sign(np.nanmean(np.stack(us, axis=1), axis=1)), index=f1r.index)
def anchors_ew(): return (ma_sig(10,25).fillna(0)+ma_sig(35,43).fillna(0)+ma_sig(63,100).fillna(0))/3.0
def carry_v4(h):
    base = ((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf], np.nan)
    return np.sign((base - base.shift(h)).replace([np.inf,-np.inf], np.nan))
def carry_v3():
    base = ((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf], np.nan)
    return np.sign(((base-base.rolling(252).mean())/base.rolling(252).std()).replace([np.inf,-np.inf], np.nan))
def carry_v1(): return np.sign(((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf], np.nan))
def value_v1(k, N, thr=0.10):
    p = crv[f"F{k}"].dropna(); ma = p.rolling(N, min_periods=N//2).mean()
    dev = ((p-ma)/ma).replace([np.inf,-np.inf], np.nan)
    return pd.Series(np.where(dev<-thr,1.0,np.where(dev>thr,-1.0,0.0)), index=dev.index)
def value_v2(N): return np.sign((f1r.shift(N)-f1r).replace([np.inf,-np.inf], np.nan))

# ── Leg variant menus ─────────────────────────────────────────────────────────
MOM = {"MA(35,43)": ma_sig(35,43), "Anchors_EW": anchors_ew(), "CTA": cta_sig()}
CAR = {"CarryMom20d": carry_v4(20), "V3_Zscore": carry_v3(), "V1_level": carry_v1()}
VAL = {"V2_BG_10yr": value_v2(2520), "V1_F8_5yr": value_v1(8,1260), "V1_F7_5yr": value_v1(7,1260)}

def metrics_port(m, c, v, same_day, tc_bps=5):
    idx = m.index.intersection(c.index).intersection(v.index).intersection(f1c.index)
    port = (m.reindex(idx).fillna(0)+c.reindex(idx).fillna(0)+v.reindex(idx).fillna(0))/3.0
    pos = (port.shift(1) if same_day else port.shift(2)).fillna(0)
    f = f1c.reindex(idx); gp = pos*f.diff()
    chg = pos.diff().abs(); chg.iloc[0]=abs(pos.iloc[0]); tc = chg*(tc_bps/10000.0/2.0)*f
    with np.errstate(invalid="ignore", divide="ignore"):
        gr = (gp/f.shift(1)).replace([np.inf,-np.inf], np.nan)
        nr = ((gp-tc)/f.shift(1)).replace([np.inf,-np.inf], np.nan)
    ag = gr[pos!=0].dropna(); an = nr[pos!=0].dropna()
    sh = float(ag.mean()/ag.std(ddof=1)*np.sqrt(252)) if len(ag)>20 else np.nan
    nsh = float(an.mean()/an.std(ddof=1)*np.sqrt(252)) if len(an)>20 else np.nan
    ar = float(gr.dropna().mean()*252*100)
    cum = gp.cumsum(); dd = float((cum-cum.cummax()).min())
    return sh, nsh, ar, dd

rows = []
for (mn, ms), (cn, cs), (vn, vs) in product(MOM.items(), CAR.items(), VAL.items()):
    for tname, sd in [("Same-Day", True), ("Lag-1", False)]:
        sh, nsh, ar, dd = metrics_port(ms, cs, vs, sd)
        rows.append(dict(Mom=mn, Carry=cn, Value=vn, Timing=tname,
                         Sharpe=round(sh,3), Net5=round(nsh,3),
                         AnnRet_pct=round(ar,2), MaxDD_USD=round(dd,1)))
df = pd.DataFrame(rows)
df.to_csv(f"{BASE}/portfolio_timing_compare.csv", index=False)
print(f"Wrote portfolio_timing_compare.csv ({len(df)} rows = {len(df)//2} combos × 2 timings)\n")

pd.set_option("display.width", 220, "display.max_rows", 200)
best3 = ("MA(35,43)","CarryMom20d","V2_BG_10yr")
print("="*92); print("BEST-OF-THREE  (MA(35,43) + CarryMom20d + V2_BG_10yr)"); print("="*92)
print(df[(df.Mom==best3[0])&(df.Carry==best3[1])&(df.Value==best3[2])].to_string(index=False))

print("\n"+"="*92); print("ALL COMBINATIONS — Same-Day, sorted by Sharpe"); print("="*92)
print(df[df.Timing=="Same-Day"].sort_values("Sharpe", ascending=False)
        [["Mom","Carry","Value","Sharpe","Net5","AnnRet_pct","MaxDD_USD"]].to_string(index=False))

print("\n"+"="*92); print("ALL COMBINATIONS — Lag-1, sorted by Sharpe"); print("="*92)
print(df[df.Timing=="Lag-1"].sort_values("Sharpe", ascending=False)
        [["Mom","Carry","Value","Sharpe","Net5","AnnRet_pct","MaxDD_USD"]].to_string(index=False))
