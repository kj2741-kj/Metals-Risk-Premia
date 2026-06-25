"""
carry_diagnostics.py
Decompose the carry Same-Day "edge" and test whether any LEGITIMATE
(no look-ahead, no IS-fit) variant captures part of it.
Convention everywhere: pnl(t) = pos(t-?) * [F1c(t) - F1c(t-1)]
  - position set at close of day D earns from D->D+1  => signal.shift(1) * diff  (CORRECT)
"""
import pandas as pd, numpy as np, warnings, sys, io
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = r"C:\Users\Kartavya\Metals Risk Premia"

f = pd.read_csv(f"{BASE}/LME_Copper_Rolling_F1_v2.csv", parse_dates=["Date"]).set_index("Date")
f.index = f.index.normalize()
f1c = f["F1_continuous"].sort_index(); f1r = f["F1_raw"].sort_index()

xls = pd.ExcelFile(f"{BASE}/Metals Futures Curve.csv")
sh = [s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()][0]
raw = pd.read_excel(xls, sh, header=None, nrows=4)
hr = [i for i in range(4) if any(k in " ".join(str(v).lower() for v in raw.iloc[i] if pd.notna(v))
                                 for k in ["date","f1","price"])]
d = pd.read_excel(xls, sh, header=hr if hr else [0,1])
if isinstance(d.columns, pd.MultiIndex):
    d.columns = ["_".join(str(p) for p in t if pd.notna(p) and "Unnamed" not in str(p)) for t in d.columns]
d.columns = [str(c).strip() for c in d.columns]
dcol = [c for c in d.columns if "date" in c.lower()][0]
d[dcol] = pd.to_datetime(d[dcol], errors="coerce")
d = d.dropna(subset=[dcol]).set_index(dcol).sort_index(); d.index = d.index.normalize()
cols = {}
for c in d.columns:
    cl = c.lower().replace(" ", "_")
    for i in range(1, 5):
        if f"f{i}_" in cl and "price" in cl: cols[f"F{i}"] = pd.to_numeric(d[c], errors="coerce"); break
crv = pd.DataFrame(cols)

carry = ((crv["F1"] - crv["F2"]) / crv["F1"]).replace([np.inf, -np.inf], np.nan)
ret_idx = f1c.index
diff = f1c.diff()

def perf(pos, label):
    pos = pos.reindex(ret_idx).fillna(0)
    pnl = pos * diff
    r = (pnl / f1c.shift(1)).replace([np.inf, -np.inf], np.nan)
    act = r[pos != 0].dropna()
    sh = float(act.mean()/act.std(ddof=1)*np.sqrt(252)) if len(act) > 20 else np.nan
    ann = float(r.dropna().mean()*252*100)
    cum = pnl.cumsum(); dd = float((cum-cum.cummax()).min())
    flips = int((pos.diff().abs() > 1e-9).sum())
    pct_in = 100*float((pos != 0).sum())/len(pos)
    return dict(label=label, sharpe=sh, ann=ann, dd=dd, flips=flips, pct_in=pct_in)

rows = []
c = np.sign(carry)

# --- baseline timing ---
rows.append(perf(c,                 "A. Same-Day level  (LOOK-AHEAD)"))
rows.append(perf(c.shift(1),        "B. Lag-1 level     (CORRECT, = honest carry)"))
rows.append(perf(c.shift(2),        "C. Lag-2 level     (robustness)"))

# --- decomposition: does same-day sign predict the NEXT day? ---
# (sign(carry_t) earning t->t+1 is exactly Lag-1; already = B)
# contemporaneous-only contribution:
contemp = (np.sign(carry).reindex(ret_idx).fillna(0) * diff)
predict = (np.sign(carry).shift(1).reindex(ret_idx).fillna(0) * diff)
print("CONTEMPORANEOUS vs PREDICTIVE split of Same-Day PnL")
print(f"  Same-Day total cum PnL : {contemp.cumsum().iloc[-1]:>12,.0f}")
print(f"  Lag-1 (predictive) PnL : {predict.cumsum().iloc[-1]:>12,.0f}")
print(f"  Pure look-ahead piece  : {contemp.cumsum().iloc[-1]-predict.cumsum().iloc[-1]:>12,.0f}\n")

# --- LEGITIMATE attempts to recover edge (all use info <= t-1) ---
# 1. carry change / steepening momentum
dcarry = carry.diff()
rows.append(perf(np.sign(dcarry).shift(1),          "D. Lag-1 carry CHANGE (steepening)"))
# 2. carry z-score (252d), level
z = ((carry - carry.rolling(252).mean())/carry.rolling(252).std()).replace([np.inf,-np.inf],np.nan)
rows.append(perf(np.sign(z).shift(1),               "E. Lag-1 carry Z-score 252d level"))
# 3. multi-day holding of correct carry (premium harvest) — hold N days
for N in [5, 10, 21, 63]:
    held = c.shift(1).reindex(ret_idx).rolling(N).mean()  # smoothed/persistent position
    rows.append(perf(np.sign(held),                 f"F{N}. Lag-1 level, {N}d-smoothed hold"))
# 4. carry + momentum agreement (overlay)
mom = np.sign(f1r.rolling(35).mean()-f1r.rolling(43).mean()).shift(1)
agree = np.where((c.shift(1)>0)&(mom>0),1,np.where((c.shift(1)<0)&(mom<0),-1,0))
rows.append(perf(pd.Series(agree,index=carry.index), "G. Carry & Momentum AGREE only (Lag-1)"))
# 5. carry weekly rebalance (slow premium): resample signal to weekly, ffill
wk = c.shift(1).resample("W-FRI").last().reindex(ret_idx).ffill()
rows.append(perf(wk,                                "H. Lag-1 level, weekly rebalance"))
# 6. continuous tanh weight on z (no hard flips), lag-1
wt = np.tanh(z).shift(1)
rows.append(perf(wt,                                "I. Lag-1 tanh(z) continuous weight"))

res = pd.DataFrame(rows)
pd.set_option("display.width",160)
print(res.to_string(index=False, formatters={
    "sharpe":lambda x:f"{x:+.3f}", "ann":lambda x:f"{x:+.1f}%",
    "dd":lambda x:f"{x:,.0f}", "pct_in":lambda x:f"{x:.0f}%"}))
