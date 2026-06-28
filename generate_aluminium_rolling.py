"""
generate_aluminium_rolling.py
Builds LME_Aluminium_Rolling_F1_v2.csv (Date, F1_raw, F1_continuous) using the SAME
ratio back-adjusted 4-phase roll-stitch as the copper v2 file. End-anchored so the
F1_continuous level equals F1_raw on the last date (matches copper v2 convention);
returns are scale-invariant so this does not change any return-based metric.
Validated separately: return corr vs copper v2 = 0.99939.
"""
import numpy as np, pandas as pd
BASE=r"C:\Users\Kartavya\Metals Risk Premia"; CURVE="Metals Futures Curve.csv"; CAL="expiry_calendars_20260526.xlsx"
raw=pd.read_excel(f"{BASE}/{CURVE}",sheet_name="ALuminium LME",header=None)
dts=pd.to_datetime(raw.iloc[3:,0],errors="coerce")
f1=pd.to_numeric(raw.iloc[3:,1],errors="coerce"); f2=pd.to_numeric(raw.iloc[3:,4],errors="coerce")
px=pd.DataFrame({"Date":dts.values,"F1_raw":f1.values,"F2_raw":f2.values}).dropna(subset=["Date"]).set_index("Date").sort_index(); px.index=px.index.normalize()
cal=pd.read_excel(f"{BASE}/{CAL}",sheet_name="LA - LME Aluminium")
cal["e"]=pd.to_datetime(cal["FUT_DLV_DT_LAST"],errors="coerce"); cal["r"]=pd.to_datetime(cal["LAST_TRADEABLE_DT"],errors="coerce")
cal=cal.dropna(subset=["r","e"]).sort_values("r").reset_index(drop=True)
rs=set(cal["r"].dt.normalize()); r2r={x["r"].normalize():x for _,x in cal.iterrows()}
td=px.index.normalize(); f1a=px["F1_raw"].values; f2a=px["F2_raw"].values; n=len(td)
fc=np.full(n,np.nan); inf2=False; p2e=None; pf2=np.nan; pf1=np.nan
for i,d in enumerate(td):
    f1v=f1a[i]; f2v=f2a[i]
    if d in rs:
        row=r2r[d]
        fc[i]=f1v if (i==0 or np.isnan(fc[i-1])) else (fc[i-1]*(f1v/pf1) if (not np.isnan(f1v) and not np.isnan(pf1) and pf1!=0) else fc[i-1])
        p2e=(row["e"].normalize()+pd.offsets.BDay(1)).normalize(); inf2=True; pf2=f2v; pf1=f1v; continue
    if inf2 and d<=p2e:
        fp=fc[i-1] if i>0 else np.nan
        fc[i]=fp*(f2v/pf2) if (not np.isnan(f2v) and not np.isnan(pf2) and pf2!=0 and not np.isnan(fp)) else fp; pf2=f2v; continue
    if inf2 and d>p2e:
        fp=fc[i-1] if i>0 else np.nan
        fc[i]=fp*(f1v/pf2) if (not np.isnan(f1v) and not np.isnan(pf2) and pf2!=0 and not np.isnan(fp)) else fp; pf1=f1v; inf2=False; continue
    fc[i]=f1v if (i==0 or np.isnan(fc[i-1])) else (fc[i-1]*(f1v/pf1) if (not np.isnan(f1v) and not np.isnan(pf1) and pf1!=0) else fc[i-1]); pf1=f1v
f1c=pd.Series(fc,index=px.index)
# end-anchor: scale so last F1_continuous == last F1_raw (returns unchanged)
scale=float(px["F1_raw"].dropna().iloc[-1])/float(f1c.dropna().iloc[-1]); f1c=f1c*scale
out=pd.DataFrame({"Date":px.index.strftime("%Y-%m-%d"),"F1_raw":px["F1_raw"].round(4).values,"F1_continuous":f1c.round(4).values})
path=f"{BASE}/LME_Aluminium_Rolling_F1_v2.csv"; out.to_csv(path,index=False)
print(f"Saved {path}  ({len(out)} rows)  F1_raw last={px['F1_raw'].dropna().iloc[-1]:.0f}  F1_cont last={f1c.dropna().iloc[-1]:.0f}")
