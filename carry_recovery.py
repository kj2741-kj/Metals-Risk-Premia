"""
carry_recovery.py
Can we legitimately beat the honest carry Sharpe (~0.10) WITHOUT look-ahead?
All positions are signal.shift(1) (Same-Day execution) or shift(2) (Lag-1) -> no look-ahead.
Reports gross + net 5bps, turnover, and a walk-forward OOS check on the winners.
"""
import pandas as pd, numpy as np, warnings, sys, io
warnings.filterwarnings("ignore"); sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = r"C:\Users\Kartavya\Metals Risk Premia"
f = pd.read_csv(f"{BASE}/LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date"); f.index=f.index.normalize()
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
    for i in range(1,16):
        if f"f{i}_" in cl and "price" in cl: cols[f"F{i}"]=pd.to_numeric(d[c],errors="coerce"); break
crv=pd.DataFrame(cols)
idx=f1c.index; diff=f1c.diff()

def perf(pos, tc=5, shift=1):
    p=pos.reindex(idx).shift(shift).fillna(0)
    pnl=p*diff
    if tc>0:
        chg=p.diff().abs(); chg.iloc[0]=abs(p.iloc[0]); pnl=pnl-chg*(tc/10000/2)*f1c
    r=(pnl/f1c.shift(1)).replace([np.inf,-np.inf],np.nan); a=r[p!=0].dropna()
    sh=float(a.mean()/a.std(ddof=1)*np.sqrt(252)) if len(a)>20 else np.nan
    flips=int((p.diff().abs()>1e-9).sum()); pct=100*float((p!=0).sum())/len(p)
    return sh, flips, pct

carry=((crv["F1"]-crv["F2"])/crv["F1"]).replace([np.inf,-np.inf],np.nan)
mom=np.sign(f1r.rolling(35).mean()-f1r.rolling(43).mean())

def z(s,w): return ((s-s.rolling(w).mean())/s.rolling(w).std()).replace([np.inf,-np.inf],np.nan)

variants={}
variants["1. Level sign (honest baseline)"]      = np.sign(carry)
variants["2. Z-score 252d sign"]                 = np.sign(z(carry,252))
variants["3. Z-score 126d sign"]                 = np.sign(z(carry,126))
variants["4. Carry-change sign (raw)"]           = np.sign(carry.diff())
variants["5. Carry-change EWMA5 smoothed"]       = np.sign(carry.diff().ewm(span=5).mean())
variants["6. Carry-change EWMA10 smoothed"]      = np.sign(carry.diff().ewm(span=10).mean())
variants["7. tanh(z252) continuous weight"]      = np.tanh(z(carry,252))
variants["8. Level, deadband |z|>0.5"]           = np.where(z(carry,252)>0.5,1,np.where(z(carry,252)<-0.5,-1,0))
variants["9. Longer carry (F1-F3)/F1 sign"]      = np.sign(((crv["F1"]-crv["F3"])/crv["F1"]).replace([np.inf,-np.inf],np.nan))
variants["10. Curve slope (F1-F6)/F1 sign"]      = np.sign(((crv["F1"]-crv["F6"])/crv["F1"]).replace([np.inf,-np.inf],np.nan)) if "F6" in crv else None
variants["11. Carry & Momentum AGREE"]           = pd.Series(np.where((np.sign(carry)>0)&(mom>0),1,np.where((np.sign(carry)<0)&(mom<0),-1,0)),index=carry.index)
variants["12. Carry-change & Mom AGREE"]         = pd.Series(np.where((np.sign(carry.diff())>0)&(mom>0),1,np.where((np.sign(carry.diff())<0)&(mom<0),-1,0)),index=carry.index)

print(f"{'Variant':40s}{'gross':>8s}{'net5':>8s}{'flips':>7s}{'%in':>6s}")
print("-"*69)
results={}
for k,v in variants.items():
    if v is None: continue
    v=pd.Series(v,index=carry.index) if not isinstance(v,pd.Series) else v
    g,_,_=perf(v,0,1); n,fl,pc=perf(v,5,1)
    results[k]=(g,n,fl,pc,v)
    print(f"{k:40s}{g:>+8.3f}{n:>+8.3f}{fl:>7d}{pc:>5.0f}%")

# --- OOS walk-forward on the most promising legitimate variants ---
print("\nWalk-forward OOS (IS=1260, OOS=252) on top legitimate variants, net 5bps:")
def oos(v):
    p_all=v.reindex(idx).shift(1).fillna(0); T=len(idx); o=1260; shs=[]
    while o<T:
        e=min(o+252,T); oi=idx[o:e]; p=p_all.reindex(oi); fo=f1c.reindex(oi)
        pnl=p*fo.diff(); chg=p.diff().abs(); chg.iloc[0]=abs(p.iloc[0]); pnl=pnl-chg*(5/10000/2)*fo
        r=(pnl/fo.shift(1)).replace([np.inf,-np.inf],np.nan); a=r[p!=0].dropna()
        if len(a)>5: shs.append(a.mean()/a.std(ddof=1)*np.sqrt(252))
        o+=252
    return np.nanmean(shs), sum(1 for x in shs if x>0), len(shs)
for k in ["1. Level sign (honest baseline)","2. Z-score 252d sign","6. Carry-change EWMA10 smoothed",
          "11. Carry & Momentum AGREE","12. Carry-change & Mom AGREE"]:
    if k in results:
        m,pp,nn=oos(results[k][4]); print(f"  {k:40s} OOS avg={m:+.3f}  ({pp}/{nn} +)")
