"""
portfolio.py — EW Portfolio Construction: Momentum + Carry + Value

Signals (per CLAUDE.md best-signal choices):
  Momentum : MA(35,43) Lag-1 on F1_raw         → ±1
  Carry    : V1 (F1-F2)/F1 Same-Day            → ±1
  Value A  : V1 F8 5yr Lag-1 (±10% threshold)  → −1/0/+1
  Value B  : V2 Baz-Granger 10yr Lag-1         → ±1

EW portfolio: Port = (1/3)×Mom + (1/3)×Carry + (1/3)×Value
Outputs results to portfolio_output/ directory.
"""

import os, sys, io, warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
DATA_DIR = os.path.join(_REPO_ROOT, "data")
OUTPUTS_DIR = os.path.join(_REPO_ROOT, "outputs")
BASE = _REPO_ROOT
OUT  = os.path.join(OUTPUTS_DIR, "portfolio_output")
os.makedirs(OUT, exist_ok=True)

# ── Load data ──────────────────────────────────────────────────────────────────
f1_df = pd.read_csv(os.path.join(DATA_DIR, "LME_Copper_Rolling_F1_v2.csv"),
                    parse_dates=["Date"]).set_index("Date")
f1_df.index = f1_df.index.normalize()
f1r = f1_df["F1_raw"].sort_index()
f1c = f1_df["F1_continuous"].sort_index()
print(f"F1 data: {len(f1r)} rows  |  {f1r.index[0].date()} → {f1r.index[-1].date()}")

# Load curve (Excel file disguised as CSV)
curve_path = os.path.join(DATA_DIR, "Metals Futures Curve.csv")
xls = pd.ExcelFile(curve_path)
print(f"Curve sheets: {xls.sheet_names}")

# Find copper sheet
cu_sheet = next((s for s in xls.sheet_names if "copper" in s.lower() and "lme" in s.lower()), None)
if cu_sheet is None:
    cu_sheet = next((s for s in xls.sheet_names if "copper" in s.lower()), xls.sheet_names[0])
print(f"Using curve sheet: '{cu_sheet}'")

# Parse curve prices
df_raw = pd.read_excel(xls, sheet_name=cu_sheet, header=None, nrows=5)
header_rows = []
for i in range(min(4, len(df_raw))):
    row_vals = [str(v).strip().lower() for v in df_raw.iloc[i].values if pd.notna(v)]
    if any(kw in " ".join(row_vals) for kw in ["date", "f1", "f2", "price"]):
        header_rows.append(i)

if len(header_rows) >= 2:
    df_c = pd.read_excel(xls, sheet_name=cu_sheet, header=header_rows)
elif len(header_rows) == 1:
    df_c = pd.read_excel(xls, sheet_name=cu_sheet, header=header_rows[0])
else:
    df_c = pd.read_excel(xls, sheet_name=cu_sheet, header=[0, 1])

# Flatten MultiIndex columns
if isinstance(df_c.columns, pd.MultiIndex):
    new_cols = []
    for col_tuple in df_c.columns:
        parts = [str(p).strip() for p in col_tuple
                 if pd.notna(p) and "Unnamed" not in str(p) and str(p).strip()]
        new_cols.append("_".join(parts) if parts else str(col_tuple))
    df_c.columns = new_cols
df_c.columns = [str(c).strip() for c in df_c.columns]

# Set Date index
date_col = [c for c in df_c.columns if "date" in c.lower()]
if date_col:
    df_c = df_c.rename(columns={date_col[0]: "Date"})
df_c["Date"] = pd.to_datetime(df_c["Date"], errors="coerce")
df_c = df_c.dropna(subset=["Date"]).set_index("Date").sort_index()
df_c.index = df_c.index.normalize()

# Extract F1–F15 price columns
curve_prices = {}
for col in df_c.columns:
    col_lower = col.lower().replace(" ", "_")
    for i in range(1, 27):
        if f"f{i}_" in col_lower and "price" in col_lower:
            curve_prices[f"F{i}"] = pd.to_numeric(df_c[col], errors="coerce")
            break
curve_px = pd.DataFrame(curve_prices, index=df_c.index)
print(f"Curve columns loaded: {list(curve_px.columns)}")
print(f"Curve data: {len(curve_px)} rows  |  {curve_px.index[0].date()} → {curve_px.index[-1].date()}")


# ── Helper: Sharpe from daily return series ───────────────────────────────────
def sharpe(ret: pd.Series) -> float:
    act = ret[ret != 0].dropna()
    if len(act) < 20:
        return np.nan
    sd = act.std(ddof=1)
    return float(act.mean() / sd * np.sqrt(252)) if sd > 0 else np.nan

def perf_metrics(pos: pd.Series, f1c_s: pd.Series, label: str) -> dict:
    idx = pos.index.intersection(f1c_s.index)
    p = pos.reindex(idx); c = f1c_s.reindex(idx)
    pnl = p * c.diff()
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pnl / c.shift(1)).replace([np.inf, -np.inf], np.nan)
    sh = sharpe(ret)
    ann_ret = float(ret.dropna().mean() * 252 * 100) if len(ret.dropna()) > 20 else np.nan
    cum = pnl.cumsum()
    roll_max = cum.cummax()
    dd = cum - roll_max
    max_dd = float(dd.min())
    active = (p != 0).sum()
    pct_flat = float(100 * (p == 0).sum() / len(p)) if len(p) > 0 else np.nan
    pos_days = int(active)
    return {
        "Label": label,
        "Sharpe": round(sh, 3) if not np.isnan(sh) else None,
        "Ann Ret %": round(ann_ret, 1) if not np.isnan(ann_ret) else None,
        "Max DD (USD/MT)": round(max_dd, 0),
        "% Flat": round(pct_flat, 1),
        "Active Days": pos_days,
        "Start": str(idx[0].date()),
        "End": str(idx[-1].date()),
    }


# ── SIGNAL 1: Momentum MA(35,43) Lag-1 ───────────────────────────────────────
print("\n--- Computing signals ---")
ma_diff = f1r.rolling(35).mean() - f1r.rolling(43).mean()
mom_pos = np.sign(ma_diff).shift(1).fillna(0)
mom_pos.name = "Momentum"
print(f"Momentum: {(mom_pos != 0).sum()} active days  |  current pos = {mom_pos.iloc[-1]:+.0f}")

# ── SIGNAL 2: Carry V1 (F1-F2)/F1 Same-Day ───────────────────────────────────
f1_curve = curve_px["F1"].dropna()
f2_curve = curve_px["F2"].dropna()
idx_c = f1_curve.index.intersection(f2_curve.index)
carry_raw = (f1_curve.reindex(idx_c) - f2_curve.reindex(idx_c)) / f1_curve.reindex(idx_c)
carry_raw = carry_raw.replace([np.inf, -np.inf], np.nan).dropna()
carry_pos = np.sign(carry_raw).reindex(f1c.index).fillna(0)
carry_pos.name = "Carry"
print(f"Carry:    {(carry_pos != 0).sum()} active days  |  current pos = {carry_pos.iloc[-1]:+.0f}")

# ── SIGNAL 3A: Value V1 F8 5yr Lag-1 ±10% threshold ─────────────────────────
if "F8" in curve_px.columns:
    f8 = curve_px["F8"].dropna()
    N_v1 = 1260
    ma_f8 = f8.rolling(N_v1, min_periods=630).mean()
    dev_f8 = (f8 - ma_f8) / ma_f8.replace(0, np.nan)
    dev_f8 = dev_f8.replace([np.inf, -np.inf], np.nan).dropna()
    val_sig_v1 = np.where(dev_f8.values < -0.10,  1.0,
                 np.where(dev_f8.values >  0.10, -1.0, 0.0))
    val_raw_v1 = pd.Series(val_sig_v1, index=dev_f8.index)
    val_pos_v1 = val_raw_v1.shift(1).fillna(0).reindex(f1c.index).fillna(0)
    val_pos_v1.name = "Value_V1_F8_5yr"
    print(f"Value V1: {(val_pos_v1 != 0).sum()} active days  |  current pos = {val_pos_v1.iloc[-1]:+.0f}")
else:
    print("WARNING: F8 column not found in curve data — Value V1 unavailable")
    val_pos_v1 = None

# ── SIGNAL 3B: Value V2 Baz-Granger 10yr Lag-1 ───────────────────────────────
N_v2 = 2520
rev_v2 = f1r.shift(N_v2) - f1r
rev_v2 = rev_v2.replace([np.inf, -np.inf], np.nan).dropna()
val_sig_v2 = np.sign(rev_v2.values)
val_raw_v2 = pd.Series(val_sig_v2, index=rev_v2.index)
val_pos_v2 = val_raw_v2.shift(1).fillna(0).reindex(f1c.index).fillna(0)
val_pos_v2.name = "Value_V2_BG10yr"
print(f"Value V2: {(val_pos_v2 != 0).sum()} active days  |  current pos = {val_pos_v2.iloc[-1]:+.0f}")


# ── Align all signals on common index ─────────────────────────────────────────
signals = [mom_pos, carry_pos]
if val_pos_v1 is not None:
    signals.append(val_pos_v1)
signals.append(val_pos_v2)

common = f1c.index
for s in signals:
    common = common.intersection(s.index)
common = common.sort_values()
print(f"\nCommon index: {len(common)} days  |  {common[0].date()} → {common[-1].date()}")

mom_a   = mom_pos.reindex(common).fillna(0)
carry_a = carry_pos.reindex(common).fillna(0)
val_v1_a = val_pos_v1.reindex(common).fillna(0) if val_pos_v1 is not None else None
val_v2_a = val_pos_v2.reindex(common).fillna(0)
f1c_a   = f1c.reindex(common)


# ── EW Portfolios ─────────────────────────────────────────────────────────────
# Portfolio A: Mom + Carry + Value V1 (F8 5yr)
if val_v1_a is not None:
    port_v1 = (mom_a + carry_a + val_v1_a) / 3.0
    port_v1.name = "Portfolio_EW_V1"
else:
    port_v1 = None

# Portfolio B: Mom + Carry + Value V2 (BG 10yr)
port_v2 = (mom_a + carry_a + val_v2_a) / 3.0
port_v2.name = "Portfolio_EW_V2"


# ══════════════════════════════════════════════════════════════════════════════
# PAIRWISE CORRELATIONS
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*70)
print("PAIRWISE POSITION CORRELATIONS")
print("="*70)
pos_df = pd.DataFrame({
    "Momentum": mom_a,
    "Carry":    carry_a,
})
if val_v1_a is not None:
    pos_df["Value_V1"] = val_v1_a
pos_df["Value_V2"] = val_v2_a

corr = pos_df.corr()
print(corr.round(3).to_string())

print("\n(Rule of thumb: EW Sharpe ≈ √3 × avg individual if all pairs < 0.3)")


# ══════════════════════════════════════════════════════════════════════════════
# INDIVIDUAL SIGNAL PERFORMANCE
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*70)
print("INDIVIDUAL SIGNAL PERFORMANCE (Gross, common period)")
print("="*70)
rows = []
rows.append(perf_metrics(mom_a,   f1c_a, "Momentum MA(35,43) Lag-1"))
rows.append(perf_metrics(carry_a, f1c_a, "Carry V1 (F1-F2)/F1 Same-Day"))
if val_v1_a is not None:
    rows.append(perf_metrics(val_v1_a, f1c_a, "Value V1 F8 5yr ±10% Lag-1"))
rows.append(perf_metrics(val_v2_a, f1c_a, "Value V2 BG 10yr Lag-1"))
ind_df = pd.DataFrame(rows)
print(ind_df[["Label","Sharpe","Ann Ret %","Max DD (USD/MT)","% Flat"]].to_string(index=False))


# ══════════════════════════════════════════════════════════════════════════════
# EW PORTFOLIO PERFORMANCE
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*70)
print("EW PORTFOLIO PERFORMANCE (Gross)")
print("="*70)
port_rows = []
if port_v1 is not None:
    port_rows.append(perf_metrics(port_v1, f1c_a, "EW Portfolio (Mom+Carry+V1_F8_5yr)"))
port_rows.append(perf_metrics(port_v2, f1c_a, "EW Portfolio (Mom+Carry+V2_BG10yr)"))
port_df = pd.DataFrame(port_rows)
print(port_df[["Label","Sharpe","Ann Ret %","Max DD (USD/MT)","% Flat"]].to_string(index=False))


# ══════════════════════════════════════════════════════════════════════════════
# SUB-PERIOD TABLE
# ══════════════════════════════════════════════════════════════════════════════
def sub_sharpe(pos: pd.Series, f1c_s: pd.Series, start=None, end=None) -> float:
    idx = pos.index
    if start: idx = idx[idx >= pd.Timestamp(start)]
    if end:   idx = idx[idx <  pd.Timestamp(end)]
    if len(idx) < 20: return np.nan
    p = pos.reindex(idx); c = f1c_s.reindex(idx)
    pnl = p * c.diff()
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = (pnl / c.shift(1)).replace([np.inf, -np.inf], np.nan)
    return sharpe(ret)

periods = [
    ("Full (2006–2025)", None,   None),
    ("Pre-2020",         None,   "2020-01-01"),
    ("2020–2021",        "2020-01-01", "2022-01-01"),
    ("Pre-2022",         None,   "2022-01-01"),
    ("Post-2022",        "2022-01-01", None),
]

print("\n" + "="*90)
print("SUB-PERIOD SHARPE TABLE")
print("="*90)
header = f"{'Period':<22}  {'Momentum':>10}  {'Carry':>10}  {'Val V1':>8}  {'Val V2':>8}  {'Port V1':>9}  {'Port V2':>9}"
print(header)
print("-"*90)
for label, s, e in periods:
    m_sh  = sub_sharpe(mom_a,   f1c_a, s, e)
    ca_sh = sub_sharpe(carry_a, f1c_a, s, e)
    v1_sh = sub_sharpe(val_v1_a, f1c_a, s, e) if val_v1_a is not None else np.nan
    v2_sh = sub_sharpe(val_v2_a, f1c_a, s, e)
    p1_sh = sub_sharpe(port_v1, f1c_a, s, e) if port_v1 is not None else np.nan
    p2_sh = sub_sharpe(port_v2, f1c_a, s, e)
    def F(x): return f"{x:+.3f}" if not np.isnan(x) else "     —"
    print(f"{label:<22}  {F(m_sh):>10}  {F(ca_sh):>10}  {F(v1_sh):>8}  {F(v2_sh):>8}  {F(p1_sh):>9}  {F(p2_sh):>9}")


# ══════════════════════════════════════════════════════════════════════════════
# SIGNAL AGREEMENT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
print("\n" + "="*70)
print("SIGNAL AGREEMENT (using V2 BG-10yr for Value in portfolio)")
print("="*70)

def sig_sign(s):
    return np.sign(s).astype(int)

for val_col, val_label in ([(val_v1_a, "Value V1 F8 5yr")] if val_v1_a is not None else []) + [(val_v2_a, "Value V2 BG10yr")]:
    sm = sig_sign(mom_a); sc = sig_sign(carry_a); sv = sig_sign(val_col)
    total = len(sm)
    all_agree_long  = ((sm == 1) & (sc == 1) & (sv == 1)).sum()
    all_agree_short = ((sm == -1) & (sc == -1) & (sv == -1)).sum()
    all_agree = all_agree_long + all_agree_short
    two_agree = ((sm == sc) | (sm == sv) | (sc == sv)).sum() - all_agree
    split = total - all_agree - two_agree
    print(f"\n  Value = {val_label}")
    print(f"  All 3 agree (same direction) : {all_agree:4d} days  ({100*all_agree/total:.1f}%)")
    print(f"    — All LONG  : {all_agree_long:4d} days  ({100*all_agree_long/total:.1f}%)")
    print(f"    — All SHORT : {all_agree_short:4d} days  ({100*all_agree_short/total:.1f}%)")
    print(f"  2 of 3 agree                 : {two_agree:4d} days  ({100*two_agree/total:.1f}%)")
    print(f"  Split (no majority)          : {split:4d} days  ({100*split/total:.1f}%)")

# Current positions
print("\n" + "="*70)
print("CURRENT POSITIONS (most recent date)")
print("="*70)
last_date = common[-1]
print(f"  Date          : {last_date.date()}")
print(f"  Momentum      : {mom_a.iloc[-1]:+.2f}")
print(f"  Carry         : {carry_a.iloc[-1]:+.2f}")
if val_v1_a is not None:
    print(f"  Value V1 F8   : {val_v1_a.iloc[-1]:+.2f}")
print(f"  Value V2 BG   : {val_v2_a.iloc[-1]:+.2f}")
if port_v1 is not None:
    print(f"  Portfolio V1  : {port_v1.iloc[-1]:+.4f}")
print(f"  Portfolio V2  : {port_v2.iloc[-1]:+.4f}")


# ══════════════════════════════════════════════════════════════════════════════
# SAVE OUTPUT
# ══════════════════════════════════════════════════════════════════════════════
out_df = pd.DataFrame({
    "Date":       common,
    "F1_cont":    f1c_a.values,
    "Mom_pos":    mom_a.values,
    "Carry_pos":  carry_a.values,
})
if val_v1_a is not None:
    out_df["Val_V1_pos"] = val_v1_a.values
    out_df["Port_V1_pos"] = port_v1.values
out_df["Val_V2_pos"]  = val_v2_a.values
out_df["Port_V2_pos"] = port_v2.values
out_df.to_csv(os.path.join(OUT, "portfolio_positions.csv"), index=False)
print(f"\nPositions saved → portfolio_output/portfolio_positions.csv")
print("Done.")
