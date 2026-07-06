"""
rolling_continuous.py
=====================
Generic continuous rolling F1 price-series builder for metals futures.

Rolling logic (return-based stitching, verified against LME Copper data):
--------------------------------------------------------------------------
  Phase 1  — before & ON roll_date:
              F1_cont[t] = F1_cont[t-1] + ΔF1[t]   (F1-delta tracking; no hard reset)
  Phase 2  — day after roll_date through FUT_DLV_DT_LAST + 1 BDay (inclusive):
              F1_cont[t] = F1_cont[t-1] + ΔF2[t]   (track held F2 position)
  Phase 3  — first day > (FUT_DLV_DT_LAST + 1 BDay)  — "bridge":
              F1_cont[t] = F1_cont[t-1] + F1[t] − F2[t-1]
  Phase 4  — after bridge until next roll:
              F1_cont[t] = F1_cont[t-1] + ΔF1[t]

Calendar dates (from expiry_calendars_20260526.xlsx):
  roll_date   = LAST_TRADEABLE_DT   (accounts for Easter / LME holiday adjustments)
  expiry_date = FUT_DLV_DT_LAST
  phase2_end  = FUT_DLV_DT_LAST + 1 BDay
  bridge_day  = FUT_DLV_DT_LAST + 2 BDay

Usage:
------
    from rolling_continuous import get_metal_rolling_f1

    df = get_metal_rolling_f1("LP")   # LME Copper
    df = get_metal_rolling_f1("LA")   # LME Aluminium
    # etc.

    # Or step-by-step:
    prices  = load_metal_prices(FUTURES_FILE, sheet_name="Copper LME", f1_col=1, f2_col=4)
    cal     = load_metal_calendar(CALENDAR_FILE, sheet_name="LP - LME Copper")
    result  = build_rolling_f1(prices, cal)
"""

from __future__ import annotations

import os
import numpy as np
import pandas as pd

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
DATA_DIR = os.path.join(_REPO_ROOT, "data")
OUTPUTS_DIR = os.path.join(_REPO_ROOT, "outputs")

# ── Default file paths (can be overridden at call time) ───────────────────────
DEFAULT_FUTURES_FILE  = os.path.join(DATA_DIR, "Metals Futures Curve.csv")
DEFAULT_CALENDAR_FILE = os.path.join(DATA_DIR, "expiry_calendars_20260526.xlsx")

# ── Metal configuration ───────────────────────────────────────────────────────
# Extend this dict to add more metals.
# f1_col / f2_col: 0-based column index in the price sheet (after iloc).
# data_start_row: first row of actual price data (0-based, after header rows).
METAL_CONFIG: dict[str, dict] = {
    "LP": {
        "name"          : "LME Copper",
        "price_sheet"   : "Copper LME",
        "calendar_sheet": "LP - LME Copper",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "LA": {
        "name"          : "LME Aluminium",
        "price_sheet"   : "Aluminium LME",
        "calendar_sheet": "LA - LME Aluminium",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "LX": {
        "name"          : "LME Zinc",
        "price_sheet"   : "Zinc LME",
        "calendar_sheet": "LX - LME Zinc",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "LN": {
        "name"          : "LME Nickel",
        "price_sheet"   : "Nickel LME",
        "calendar_sheet": "LN - LME Nickel",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "LL": {
        "name"          : "LME Lead",
        "price_sheet"   : "Lead LME",
        "calendar_sheet": "LL - LME Lead",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "LT": {
        "name"          : "LME Tin",
        "price_sheet"   : "Tin LME",
        "calendar_sheet": "LT - Custom (LT)",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "GC": {
        "name"          : "COMEX Gold",
        "price_sheet"   : "Gold COMEX",
        "calendar_sheet": "GC - COMEX Gold",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "SI": {
        "name"          : "COMEX Silver",
        "price_sheet"   : "Silver COMEX",
        "calendar_sheet": "SI - COMEX Silver",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "PL": {
        "name"          : "COMEX Platinum",
        "price_sheet"   : "Platinum COMEX",
        "calendar_sheet": "PL - COMEX Platinum",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
    "PA": {
        "name"          : "COMEX Palladium",
        "price_sheet"   : "Palladium COMEX",
        "calendar_sheet": "PA - COMEX Palladium",
        "f1_col"        : 1,
        "f2_col"        : 4,
        "data_start_row": 3,
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# 1. Data loaders
# ─────────────────────────────────────────────────────────────────────────────

def load_metal_prices(
    filepath: str,
    sheet_name: str,
    f1_col: int = 1,
    f2_col: int = 4,
    data_start_row: int = 3,
) -> pd.DataFrame:
    """
    Load F1 and F2 raw closing prices from the futures curve workbook.

    Parameters
    ----------
    filepath       : Path to the futures curve file (xlsx or xlsx disguised as csv).
    sheet_name     : Worksheet name, e.g. "Copper LME".
    f1_col         : 0-based column index for F1 price (default 1).
    f2_col         : 0-based column index for F2 price (default 4).
    data_start_row : First row of price data, 0-based (default 3, skipping 3 header rows).

    Returns
    -------
    DataFrame indexed by Date with columns: F1_raw, F2_raw.
    """
    raw   = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    dates = pd.to_datetime(raw.iloc[data_start_row:, 0], errors="coerce")
    f1    = pd.to_numeric(raw.iloc[data_start_row:, f1_col], errors="coerce")
    f2    = pd.to_numeric(raw.iloc[data_start_row:, f2_col], errors="coerce")

    df = pd.DataFrame({"Date": dates.values, "F1_raw": f1.values, "F2_raw": f2.values})
    df["Date"] = pd.to_datetime(df["Date"])
    df = (df.dropna(subset=["Date"])
            .set_index("Date")
            .sort_index())
    return df


def load_metal_calendar(
    filepath: str,
    sheet_name: str,
) -> pd.DataFrame:
    """
    Load roll and expiry calendar for a metal.

    Uses LAST_TRADEABLE_DT as the roll date (holiday-adjusted) and
    FUT_DLV_DT_LAST as the expiry anchor.

    Returns
    -------
    DataFrame with columns: Contract, roll_date, expiry_date  (both pd.Timestamp).
    """
    cal = pd.read_excel(filepath, sheet_name=sheet_name)
    cal["roll_date"]   = pd.to_datetime(cal["LAST_TRADEABLE_DT"], errors="coerce")
    cal["expiry_date"] = pd.to_datetime(cal["FUT_DLV_DT_LAST"],   errors="coerce")
    cal = (cal.dropna(subset=["roll_date", "expiry_date"])
              .sort_values("roll_date")
              .reset_index(drop=True))
    return cal[["Contract", "roll_date", "expiry_date"]]


# ─────────────────────────────────────────────────────────────────────────────
# 2. Core rolling algorithm
# ─────────────────────────────────────────────────────────────────────────────

def build_rolling_f1(
    prices: pd.DataFrame,
    calendar: pd.DataFrame,
) -> pd.DataFrame:
    """
    Build a continuous rolling F1 price series using return-based stitching.

    Parameters
    ----------
    prices   : DataFrame indexed by Date with columns F1_raw, F2_raw.
    calendar : DataFrame with columns Contract, roll_date, expiry_date.

    Returns
    -------
    DataFrame with columns:
        F1_raw, F2_raw, F1_continuous,
        Phase, is_roll_date, is_bridge_date, active_contract
    """
    roll_set    = set(calendar["roll_date"].dt.normalize())
    roll_to_row = {r["roll_date"].normalize(): r for _, r in calendar.iterrows()}

    trading_dates = prices.index.normalize()
    f1_arr = prices["F1_raw"].values
    f2_arr = prices["F2_raw"].values
    n      = len(trading_dates)

    f1_cont      = np.full(n, np.nan)
    phase_labels = np.full(n, "", dtype=object)
    is_roll      = np.zeros(n, dtype=bool)
    is_bridge    = np.zeros(n, dtype=bool)
    active_cont  = np.full(n, "", dtype=object)

    # ── State variables ────────────────────────────────────────────────────
    in_f2_phase      = False
    current_expiry   = None    # FUT_DLV_DT_LAST (normalized)
    phase2_end       = None    # current_expiry + 1 BDay
    current_contract = "—"
    prev_f2          = np.nan
    prev_f1          = np.nan

    for i, d in enumerate(trading_dates):
        f1v = f1_arr[i]
        f2v = f2_arr[i]

        # ── Roll day ──────────────────────────────────────────────────────
        if d in roll_set:
            row = roll_to_row[d]

            # F1-delta tracking at roll day (no reset to raw price)
            if i == 0 or np.isnan(f1_cont[i - 1]):
                f1_cont[i] = f1v
            elif not np.isnan(f1v) and not np.isnan(prev_f1):
                f1_cont[i] = f1_cont[i - 1] + (f1v - prev_f1)
            else:
                f1_cont[i] = f1_cont[i - 1]

            phase_labels[i] = "F1_Direct_RollDay"
            is_roll[i]      = True
            active_cont[i]  = row["Contract"]

            current_expiry   = row["expiry_date"].normalize()
            phase2_end       = (current_expiry + pd.offsets.BDay(1)).normalize()
            current_contract = row["Contract"]
            in_f2_phase      = True
            prev_f2          = f2v
            prev_f1          = f1v
            continue

        # ── Phase 2: F2 tracking (through FUT_DLV_DT_LAST + 1 BDay) ──────
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

        # ── Phase 3: Bridge ───────────────────────────────────────────────
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

        # ── Phase 1 / 4: F1 tracking ─────────────────────────────────────
        if i == 0 or np.isnan(f1_cont[i - 1]):
            f1_cont[i] = f1v
        elif not np.isnan(f1v) and not np.isnan(prev_f1):
            f1_cont[i] = f1_cont[i - 1] + (f1v - prev_f1)
        else:
            f1_cont[i] = f1_cont[i - 1]

        phase_labels[i] = "F1_Tracking"
        active_cont[i]  = current_contract or "—"
        prev_f1         = f1v

    # ── Assemble result ────────────────────────────────────────────────────
    out = prices[["F1_raw", "F2_raw"]].copy()
    out["F1_continuous"]   = f1_cont
    out["Phase"]           = phase_labels
    out["is_roll_date"]    = is_roll
    out["is_bridge_date"]  = is_bridge
    out["active_contract"] = active_cont
    return out


# ─────────────────────────────────────────────────────────────────────────────
# 3. High-level wrapper
# ─────────────────────────────────────────────────────────────────────────────

def get_metal_rolling_f1(
    metal_code: str,
    futures_file: str  = DEFAULT_FUTURES_FILE,
    calendar_file: str = DEFAULT_CALENDAR_FILE,
    verbose: bool      = True,
) -> pd.DataFrame:
    """
    End-to-end loader + builder for any configured metal.

    Parameters
    ----------
    metal_code    : Key from METAL_CONFIG, e.g. "LP" for LME Copper.
    futures_file  : Path to the futures curve workbook.
    calendar_file : Path to the expiry calendars workbook.
    verbose       : Print progress messages if True.

    Returns
    -------
    DataFrame with columns: F1_raw, F2_raw, F1_continuous, Phase,
                            is_roll_date, is_bridge_date, active_contract.
    """
    if metal_code not in METAL_CONFIG:
        raise ValueError(
            f"Unknown metal code '{metal_code}'. "
            f"Available: {list(METAL_CONFIG.keys())}"
        )

    cfg = METAL_CONFIG[metal_code]
    if verbose:
        print(f"[{metal_code}] Loading prices from sheet '{cfg['price_sheet']}' ...")

    prices = load_metal_prices(
        filepath       = futures_file,
        sheet_name     = cfg["price_sheet"],
        f1_col         = cfg["f1_col"],
        f2_col         = cfg["f2_col"],
        data_start_row = cfg["data_start_row"],
    )

    if verbose:
        print(f"[{metal_code}] {len(prices)} trading days "
              f"({prices.index[0].date()} -> {prices.index[-1].date()})")
        print(f"[{metal_code}] Loading calendar from sheet '{cfg['calendar_sheet']}' ...")

    cal = load_metal_calendar(
        filepath   = calendar_file,
        sheet_name = cfg["calendar_sheet"],
    )

    if verbose:
        print(f"[{metal_code}] {len(cal)} contracts in calendar")
        print(f"[{metal_code}] Building continuous rolling F1 ...")

    result = build_rolling_f1(prices, cal)

    if verbose:
        rolls   = result["is_roll_date"].sum()
        bridges = result["is_bridge_date"].sum()
        print(f"[{metal_code}] Done: {rolls} roll events, {bridges} bridge events")

    return result


# ─────────────────────────────────────────────────────────────────────────────
# 4. Quick self-test when run directly
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    result = get_metal_rolling_f1("LP", verbose=True)
    print("\nF1_continuous summary:")
    print(result["F1_continuous"].describe().to_string())

    out_path = os.path.join(DATA_DIR, "LME_Copper_Rolling_F1.csv")
    out      = result.reset_index()
    out["Date"] = out["Date"].dt.strftime("%Y-%m-%d")
    out.to_csv(out_path, index=False, float_format="%.4f")
    print(f"\nSaved -> {out_path}  ({len(out)} rows)")
