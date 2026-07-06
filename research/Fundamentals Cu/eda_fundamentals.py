"""
eda_fundamentals.py — profile every file in fundamental_data/
Outputs a per-file profile + per-category roll-up to console and eda_report.csv
"""
import os, sys, io, glob, warnings
warnings.filterwarnings("ignore")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import numpy as np, pandas as pd

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fundamental_data")

def infer_freq(idx):
    if len(idx) < 3: return "?"
    d = pd.Series(idx).sort_values().diff().dropna().dt.days
    med = d.median()
    if med <= 1.5: return "B/D"      # business/daily
    if med <= 4:   return "~daily"
    if 5 <= med <= 9:  return "weekly"
    if 25 <= med <= 35: return "monthly"
    if 80 <= med <= 100: return "quarterly"
    if med > 100: return "irregular/Q+"
    return f"~{med:.0f}d"

def load_any(path):
    """Return a tidy (date-indexed) df with a single value col where possible."""
    df = pd.read_csv(path)
    # find date col
    dcol = None
    for c in df.columns:
        if str(c).lower() in ("date","dates","time"): dcol = c; break
    if dcol is None: dcol = df.columns[0]
    df[dcol] = pd.to_datetime(df[dcol], errors="coerce")
    df = df.dropna(subset=[dcol]).set_index(dcol).sort_index()
    return df

rows = []
cats = sorted([d for d in os.listdir(ROOT) if os.path.isdir(os.path.join(ROOT,d))])
for cat in cats:
    files = sorted(glob.glob(os.path.join(ROOT, cat, "*.csv")))
    for f in files:
        name = os.path.basename(f).replace(".csv","")
        try:
            df = load_any(f)
            valcols = [c for c in df.columns if df[c].dtype != "O" or pd.to_numeric(df[c],errors="coerce").notna().any()]
            vc = valcols[0] if valcols else df.columns[0]
            s = pd.to_numeric(df[vc], errors="coerce")
            idx = df.index
            n = len(df)
            nmiss = int(s.isna().sum())
            # duplicates in index
            ndup = int(idx.duplicated().sum())
            # gaps: max consecutive day gap
            gap = pd.Series(idx).sort_values().diff().dropna().dt.days
            rows.append(dict(
                category=cat, file=name, ncols=df.shape[1], nrows=n,
                start=idx.min().date() if n else None,
                end=idx.max().date() if n else None,
                freq=infer_freq(idx), n_missing=nmiss, n_dup_dates=ndup,
                max_gap_days=int(gap.max()) if len(gap) else 0,
                vmin=round(float(s.min()),3) if s.notna().any() else None,
                vmax=round(float(s.max()),3) if s.notna().any() else None,
                vmean=round(float(s.mean()),3) if s.notna().any() else None,
                has_neg=bool((s<0).any()),
                n_zero=int((s==0).sum()),
                cols=";".join(map(str, df.columns))[:60],
            ))
        except Exception as e:
            rows.append(dict(category=cat, file=name, ncols=None, nrows=None,
                             start=None, end=None, freq="ERR", n_missing=None,
                             n_dup_dates=None, max_gap_days=None, vmin=None, vmax=None,
                             vmean=None, has_neg=None, n_zero=None, cols=f"ERROR: {e}"))

rep = pd.DataFrame(rows)
rep.to_csv(os.path.join(os.path.dirname(ROOT), "eda_report.csv"), index=False)

pd.set_option("display.width", 200, "display.max_columns", 30, "display.max_colwidth", 40)
for cat in cats:
    sub = rep[rep.category==cat]
    print(f"\n{'='*120}\n### {cat}  ({len(sub)} files)\n{'='*120}")
    print(sub[["file","nrows","start","end","freq","n_missing","max_gap_days",
               "vmin","vmax","vmean","has_neg"]].to_string(index=False))

print("\n\n#### FREQUENCY DISTRIBUTION")
print(rep.freq.value_counts().to_string())
print("\n#### FILES WITH NEGATIVE VALUES")
print(rep[rep.has_neg==True][["category","file","vmin","vmax"]].to_string(index=False))
print("\n#### FILES WITH MISSING VALUES (>0)")
print(rep[rep.n_missing>0][["category","file","nrows","n_missing"]].to_string(index=False))
print("\n#### DATE COVERAGE SPREAD")
print(f"Earliest start across all: {rep.start.min()}")
print(f"Latest end across all:     {rep.end.max()}")
print(f"Saved full report -> eda_report.csv")
