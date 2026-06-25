"""
carry_full_compare.py
Full apples-to-apples comparison of every carry signal on the dashboard
+ the carry-change ('v4 raw') and new 20d/60d carry-momentum candidates.
All positions = signal.shift(1) (Same-Day execution, no look-ahead).
Reports gross / net5 / net10 / net20, turnover, OOS net5, and a shift2
latency check (microstructure tell: a real edge decays gracefully).
"""
import pandas as pd, numpy as np, warnings, sys, io
warnings.filterwarnings("ignore"); sys.stdout=io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE=r"C:\Users\Kartavya\Metals Risk Premia"
f=pd.read_csv(f"{BASE}/LME_Copper_Rolling_F1_v2.csv",parse_dates=["Date"]).set_index("Date"); f.index=f.index.normalize()
f1c=f["F1_continuous"].sort_index(); f1r=f["F1_raw"].sort_index()
xls=pd.ExcelFile(f"{BASE}/Metals Futures Curve.csv")
sh=[s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()][0]
rw=pd.read_excel(xls,sh,header=None,nrows=4)
hr=[i for i in range(4) if any(k in " ".join(str(v).lower() for v in rw.iloc[i] if pd.notna(v)) for k in ["date","f1","price"])]
d=pd.read_excel(xls,sh,header=hr if hr else [0,1])
if isinstance(d.columns,pd.MultiIndex): d.columns=["_".join(str(p) for p in t if pd.notna(p) and "Unnamed" not in str(p)) for t in d.columns]
d.columns=[str(c).strip() for c in d.columns]; dc=[c for c in d.columns if "date" in c.lower()][0]
d[dc]=pd.to_datetime(d[dc],errors="coerce"); d=d.dropna(subset=[dc]).set_index(dc).sort_index(); d.index=d.index.normalize()
cols={}
for c in d.columns:
    cl=c.lower().replace(" ","_")
    for i in range(1,25):
        if f"f{i}_" in cl and "price" in cl: cols[f"F{i}"]=pd.to_numeric(d[c],errors="coerce"); break
crv=pd.DataFrame(cols); idx=f1c.index; diff=f1c.diff()

def sh_tc(pos,tc,shift=1):
    p=pos.reindex(idx).shift(shift).fillna(0); pnl=p*diff
    if tc>0: chg=p.diff().abs(); chg.iloc[0]=abs(p.iloc[0]); pnl=pnl-chg*(tc/10000/2)*f1c
    r=(pnl/f1c.shift(1)).replace([np.inf,-np.inf],np.nan); a=r[p!=0].dropna()
    return float(a.mean()/a.std(ddof=1)*np.sqrt(252)) if len(a)>20 else np.nan
def flips(pos): p=pos.reindex(idx).shift(1).fillna(0); return int((p.diff().abs()>1e-9).sum())
def oos5(pos):
    pa=pos.reindex(idx).shift(1).fillna(0); T=len(idx); o=1260; shs=[]
    while o<T:
        e=min(o+252,T); oi=idx[o:e]; p=pa.reindex(oi); fo=f1c.reindex(oi); pnl=p*fo.diff()
        chg=p.diff().abs(); chg.iloc[0]=abs(p.iloc[0]); pnl=pnl-chg*(5/10000/2)*fo
        r=(pnl/fo.shift(1)).replace([np.inf,-np.inf],np.nan); a=r[p!=0].dropna()
        if len(a)>5: shs.append(a.mean()/a.std(ddof=1)*np.sqrt(252))
        o+=252
    return np.nanmean(shs)

carry=((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf],np.nan)
def zc(w): return ((carry-carry.rolling(w).mean())/carry.rolling(w).std()).replace([np.inf,-np.inf],np.nan)

V={}
V["V1 (F1-F2)/F1 level"]   = np.sign(carry)
V["V1 (F1-F3)/F1 level"]   = np.sign(((crv["F1"]-crv["F3"])/crv["F1"]).replace([np.inf,-np.inf],np.nan))
V["V2 Slope (F3-F15)"]     = np.sign(((crv["F3"]-crv["F15"])/crv["F15"]).replace([np.inf,-np.inf],np.nan)) if "F15" in crv else None
V["V2 Slope (F6-F18)"]     = np.sign(((crv["F6"]-crv["F18"])/crv["F18"]).replace([np.inf,-np.inf],np.nan)) if "F18" in crv else None
V["V3 Z-score sign 252d"]  = np.sign(zc(252))
V["V3 Z deadband |z|>0.5"] = pd.Series(np.where(zc(252)>0.5,1,np.where(zc(252)<-0.5,-1,0)),index=carry.index)
V["Carry-change raw (1d)"] = np.sign(carry.diff())
V["Carry-mom 20d"]         = np.sign(carry-carry.shift(20))
V["Carry-mom 60d"]         = np.sign(carry-carry.shift(60))
V["Carry-mom 20d z-dband"] = pd.Series(np.where((carry-carry.shift(20))>0,1,np.where((carry-carry.shift(20))<0,-1,0)),index=carry.index)

print(f"{'Signal':24s}{'gross':>8s}{'net5':>8s}{'net10':>8s}{'net20':>8s}{'OOS5':>8s}{'flips':>7s}{'shift2_n5':>10s}")
print("-"*81)
for k,v in V.items():
    if v is None: continue
    v=pd.Series(v,index=carry.index) if not isinstance(v,pd.Series) else v
    print(f"{k:24s}{sh_tc(v,0):>+8.3f}{sh_tc(v,5):>+8.3f}{sh_tc(v,10):>+8.3f}{sh_tc(v,20):>+8.3f}"
          f"{oos5(v):>+8.3f}{flips(v):>7d}{sh_tc(v,5,2):>+10.3f}")
print("\nshift2_n5 = net5 Sharpe if executed one extra day later (latency/microstructure check).")
print("A robust edge decays gracefully shift1->shift2; a bounce artifact collapses.")
