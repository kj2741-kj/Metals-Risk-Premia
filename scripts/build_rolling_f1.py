"""
build_rolling_f1.py
====================
Constructs a continuous rolling F1 price series for LME Copper.

Rolling logic (return-based / panama-canal stitching):
-------------------------------------------------------
  Phase 1  – Before & ON roll_date          : F1_cont[t] = F1_cont[t-1] + ΔF1[t]
              (F1-delta tracking; no hard reset to raw price at roll day)
  Phase 2  – After roll_date through
              FUT_DLV_DT_LAST + 1 BDay inclusive : F1_cont[t] = F1_cont[t-1] + ΔF2[t]
  Phase 3  – First trading day > FUT_DLV_DT_LAST + 1 BDay ("bridge"):
              F1_cont[t] = F1_cont[t-1] + F1[t] − F2[t-1]
  Phase 4  – Subsequent days until the next roll_date:
              F1_cont[t] = F1_cont[t-1] + ΔF1[t]

Key dates (from expiry_calendars_20260526.xlsx, "LP - LME Copper" sheet):
  roll_date  = FUT_DLV_DT_LAST − 2 BDay  (= LAST_TRADEABLE_DT for all non-holiday months)
  expiry     = FUT_DLV_DT_LAST
  phase2_end = FUT_DLV_DT_LAST + 1 BDay  (Bloomberg relabels after last delivery)
  bridge_day = FUT_DLV_DT_LAST + 2 BDay

Verified against user-supplied "Correct F1_continuos" column (col I of LME_Copper_Rolling_F1.csv):
  Jan 2006 (LPF06) roll=Jan16, expiry=Jan18:
    Phase2 Jan17-19, bridge Jan20  → F1_cont(Jan20)=4671 ✓
  Feb 2006 (LPG06) roll=Feb13, expiry=Feb15:
    Phase2 Feb14-16, bridge Feb17  → F1_cont(Feb17)=4918 ✓
  Mar 2006 (LPH06) roll=Mar13, expiry=Mar15:
    Phase2 Mar14-16, bridge Mar17  → F1_cont(Mar17)=5264.5 ✓
"""

import os
import numpy as np
import pandas as pd

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
DATA_DIR = os.path.join(_REPO_ROOT, "data")
OUTPUTS_DIR = os.path.join(_REPO_ROOT, "outputs")

FUTURES_FILE  = os.path.join(DATA_DIR, "Metals Futures Curve.csv")
CALENDAR_FILE = os.path.join(DATA_DIR, "expiry_calendars_20260526.xlsx")
OUTPUT_CSV    = os.path.join(DATA_DIR, "LME_Copper_Rolling_F1.csv")


# ─────────────────────────────────────────────────────────────────────────────
# 1. Load raw price data (Copper LME sheet)
# ─────────────────────────────────────────────────────────────────────────────
def load_copper_prices() -> pd.DataFrame:
    raw   = pd.read_excel(FUTURES_FILE, sheet_name="Copper LME", header=None)
    dates = pd.to_datetime(raw.iloc[3:, 0], errors="coerce")
    f1    = pd.to_numeric(raw.iloc[3:, 1], errors="coerce")
    f2    = pd.to_numeric(raw.iloc[3:, 4], errors="coerce")

    df = pd.DataFrame({"Date": dates.values, "F1_raw": f1.values, "F2_raw": f2.values})
    df["Date"] = pd.to_datetime(df["Date"])
    df = (df.dropna(subset=["Date"])
            .set_index("Date")
            .sort_index())
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 2. Load expiry / rolling calendar for LME Copper
# ─────────────────────────────────────────────────────────────────────────────
def load_copper_calendar() -> pd.DataFrame:
    cal = pd.read_excel(CALENDAR_FILE, sheet_name="LP - LME Copper")

    # New: expiry = FUT_DLV_DT_LAST (old used FUT_DLV_DT_FIRST)
    cal["expiry_date"] = pd.to_datetime(cal["FUT_DLV_DT_LAST"], errors="coerce")

    # Roll date: use LAST_TRADEABLE_DT from the calendar — it already accounts
    # for Easter and other LME holiday adjustments (4 April months where
    # FUT_DLV_DT_LAST − 2 BDay falls on Easter Monday, a non-trading day).
    cal["roll_date"] = pd.to_datetime(cal["LAST_TRADEABLE_DT"], errors="coerce")

    cal = (cal.dropna(subset=["roll_date", "expiry_date"])
              .sort_values("roll_date")
              .reset_index(drop=True))
    return cal[["Contract", "roll_date", "expiry_date"]]


# ─────────────────────────────────────────────────────────────────────────────
# 3. Build the continuous rolling F1 series
# ─────────────────────────────────────────────────────────────────────────────
def build_continuous_f1(prices: pd.DataFrame,
                         calendar: pd.DataFrame) -> pd.DataFrame:
    """
    Returns DataFrame with columns:
        F1_raw, F2_raw, F1_continuous,
        Phase, is_roll_date, is_bridge_date, active_contract
    """
    roll_set = set(calendar["roll_date"].dt.normalize())

    roll_to_row = {
        row["roll_date"].normalize(): row
        for _, row in calendar.iterrows()
    }

    trading_dates = prices.index.normalize()
    f1_arr = prices["F1_raw"].values
    f2_arr = prices["F2_raw"].values
    n      = len(trading_dates)

    f1_cont      = np.full(n, np.nan)
    phase_labels = np.full(n, "", dtype=object)
    is_roll      = np.zeros(n, dtype=bool)
    is_bridge    = np.zeros(n, dtype=bool)
    active_cont  = np.full(n, "", dtype=object)

    # ── State ──────────────────────────────────────────────────────────────
    in_f2_phase      = False
    current_expiry   = None   # FUT_DLV_DT_LAST (normalised)
    phase2_end       = None   # current_expiry + 1 BDay (inclusive Phase-2 boundary)
    current_contract = "—"
    prev_f2          = np.nan
    prev_f1          = np.nan

    for i, d in enumerate(trading_dates):
        f1v = f1_arr[i]
        f2v = f2_arr[i]

        # ── Roll Day ─────────────────────────────────────────────────────────
        if d in roll_set:
            row = roll_to_row[d]

            # F1-delta tracking — NO hard reset to raw price.
            # On very first trading day of the series, initialise from raw.
            if i == 0 or np.isnan(f1_cont[i - 1]):
                f1_cont[i] = f1v
            else:
                if not np.isnan(f1v) and not np.isnan(prev_f1):
                    f1_cont[i] = f1_cont[i - 1] + (f1v - prev_f1)
                else:
                    f1_cont[i] = f1_cont[i - 1]

            phase_labels[i] = "F1_Direct_RollDay"
            is_roll[i]      = True
            active_cont[i]  = row["Contract"]

            # Transition to Phase 2 from the NEXT trading day
            current_expiry   = row["expiry_date"].normalize()
            phase2_end       = (current_expiry + pd.offsets.BDay(1)).normalize()
            current_contract = row["Contract"]
            in_f2_phase      = True
            prev_f2          = f2v   # F2 on roll day → first delta denominator
            prev_f1          = f1v
            continue

        # ── Phase 2: F2 tracking (through FUT_DLV_DT_LAST + 1 BDay) ─────────
        if in_f2_phase and d <= phase2_end:
            f1_prev = f1_cont[i - 1] if i > 0 else np.nan

            if not np.isnan(f2v) and not np.isnan(prev_f2) and not np.isnan(f1_prev):
                f1_cont[i] = f1_prev + (f2v - prev_f2)
            else:
                f1_cont[i] = f1_prev

            phase_labels[i] = "F2_Tracking"
            active_cont[i]  = current_contract
            prev_f2         = f2v
            continue

        # ── Phase 3: Bridge (first trading day after phase2_end) ─────────────
        if in_f2_phase and d > phase2_end:
            f1_prev = f1_cont[i - 1] if i > 0 else np.nan

            if not np.isnan(f1v) and not np.isnan(prev_f2) and not np.isnan(f1_prev):
                f1_cont[i] = f1_prev + (f1v - prev_f2)
            else:
                f1_cont[i] = f1_prev

            phase_labels[i] = "Bridge"
            is_bridge[i]    = True
            active_cont[i]  = current_contract
            prev_f1         = f1v
            in_f2_phase     = False
            continue

        # ── Phase 1 / 4: F1 tracking ──────────────────────────────────────────
        if i == 0 or np.isnan(f1_cont[i - 1]):
            f1_cont[i] = f1v
        else:
            if not np.isnan(f1v) and not np.isnan(prev_f1):
                f1_cont[i] = f1_cont[i - 1] + (f1v - prev_f1)
            else:
                f1_cont[i] = f1_cont[i - 1]

        phase_labels[i] = "F1_Tracking"
        active_cont[i]  = current_contract if current_contract else "—"
        prev_f1         = f1v

    # ── Assemble output ────────────────────────────────────────────────────────
    out = prices[["F1_raw", "F2_raw"]].copy()
    out["F1_continuous"]   = f1_cont
    out["Phase"]           = phase_labels
    out["is_roll_date"]    = is_roll
    out["is_bridge_date"]  = is_bridge
    out["active_contract"] = active_cont
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 4. Verification — print rows around roll events and compare to reference
# ─────────────────────────────────────────────────────────────────────────────
def verify_roll_events(result: pd.DataFrame, calendar: pd.DataFrame, n_events: int = 5):
    print("\n" + "=" * 80)
    print(f"VERIFICATION — first {n_events} roll events")
    print("=" * 80)
    for _, row in calendar.head(n_events + 8).iterrows():
        r = row["roll_date"]
        e = row["expiry_date"]
        if r < result.index[0]:
            continue
        if n_events <= 0:
            break
        n_events -= 1
        print(f"\n  {row['Contract']}  |  roll={r.date()}  |  expiry(DLV_LAST)={e.date()}")
        window = result.loc[
            r - pd.offsets.BDay(2) : e + pd.offsets.BDay(3),
            ["F1_raw", "F2_raw", "F1_continuous", "Phase"]
        ]
        print(window.to_string())


def compare_with_reference(result: pd.DataFrame, ref_csv: str):
    """Compare F1_continuous against the 'Correct F1_continuos' column in the reference CSV."""
    try:
        ref = pd.read_csv(ref_csv, usecols=["Date", "Correct F1_continuos"])
        ref["Date"] = pd.to_datetime(ref["Date"], format="mixed")
        ref = ref.dropna(subset=["Correct F1_continuos"]).set_index("Date")
        ref.index = ref.index.normalize()

        common = result.index.intersection(ref.index)
        diff = (result.loc[common, "F1_continuous"] - ref.loc[common, "Correct F1_continuos"]).abs()
        bad  = diff[diff > 0.01]

        print(f"\n  Reference rows compared : {len(common)}")
        print(f"  Exact matches (tol 0.01): {len(common) - len(bad)}")
        if len(bad) > 0:
            print(f"  Mismatches              : {len(bad)}")
            print(bad.head(20).to_string())
        else:
            print("  All values match the reference!")
    except Exception as exc:
        print(f"  (reference comparison skipped: {exc})")


# ─────────────────────────────────────────────────────────────────────────────
# 5. Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("Loading LME Copper price data...")
    prices = load_copper_prices()
    print(f"  {len(prices)} trading days  |  {prices.index[0].date()} -> {prices.index[-1].date()}")

    print("Loading expiry / rolling calendar (FUT_DLV_DT_LAST logic)...")
    calendar = load_copper_calendar()
    print(f"  {len(calendar)} contracts  |  {calendar['roll_date'].iloc[0].date()} -> "
          f"{calendar['roll_date'].iloc[-1].date()}")

    print("Building continuous rolling F1 series...")
    result = build_continuous_f1(prices, calendar)

    n_rolls   = result["is_roll_date"].sum()
    n_bridges = result["is_bridge_date"].sum()
    print(f"  Roll events   : {n_rolls}")
    print(f"  Bridge events : {n_bridges}")
    print(f"  Phase counts  :\n{result['Phase'].value_counts().to_string()}")

    print("\nF1_continuous summary:")
    print(result["F1_continuous"].describe().to_string())

    verify_roll_events(result, calendar, n_events=5)

    print("\nComparing against reference ('Correct F1_continuos' column)...")
    compare_with_reference(result, OUTPUT_CSV)

    # Save
    out_csv = result.reset_index()
    out_csv["Date"] = out_csv["Date"].dt.strftime("%Y-%m-%d")
    out_csv.to_csv(OUTPUT_CSV, index=False, float_format="%.4f")
    print(f"\nSaved -> {OUTPUT_CSV}  ({len(out_csv)} rows)")

    return result


if __name__ == "__main__":
    result = main()
