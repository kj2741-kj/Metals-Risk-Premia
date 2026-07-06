"""
Spread construction for Paper 1.

All spreads expressed in $/gal.

Bloomberg unit assumptions:
  CAP, BAP, DAE  : stored as $/gal  → use directly
  PCW            : stored as $/lb   → multiply by ethylene density (lb/gal) to get $/gal
  PGP            : stored as $/lb   → multiply by propylene density (lb/gal) to get $/gal

Processing Spread = F1_Petrochemical($/gal) − F1_NGL($/gal)
"""

import pandas as pd
import numpy as np
from config import (
    START_DATE, END_DATE,
    ETHYLENE_DENSITY, PROPYLENE_DENSITY,
)

# Spread definitions: (petrochem_ticker, density, NGL_ticker)
INTRA_SPREAD_DEFS = {
    "Ethane-Ethylene":   ("PCW", ETHYLENE_DENSITY,  "CAP"),
    "Propane-Ethylene":  ("PCW", ETHYLENE_DENSITY,  "BAP"),
    "Propane-Propylene": ("PGP", PROPYLENE_DENSITY, "BAP"),
    "Butane-Ethylene":   ("PCW", ETHYLENE_DENSITY,  "DAE"),
    "Butane-Propylene":  ("PGP", PROPYLENE_DENSITY, "DAE"),
}

# Inter-spread definitions: (spread_name_1, spread_name_2)
INTER_SPREAD_DEFS = {
    "Ethane-Ethylene vs Propane-Propylene": ("Ethane-Ethylene",   "Propane-Propylene"),
    "Propane-Ethylene vs Propane-Propylene":("Propane-Ethylene",  "Propane-Propylene"),
    "Propane-Ethylene vs Butane-Ethylene":  ("Propane-Ethylene",  "Butane-Ethylene"),
    "Propane-Propylene vs Butane-Ethylene": ("Propane-Propylene", "Butane-Ethylene"),
    "Butane-Ethylene vs Butane-Propylene":  ("Butane-Ethylene",   "Butane-Propylene"),
}


def build_price_series(raw: dict) -> pd.DataFrame:
    """
    Convert raw Bloomberg prices to $/gal for all 5 commodities.
    Returns DataFrame indexed by Date with one column per ticker ($/gal).
    """
    petrochem = {
        "PCW": raw["PCW"]["F1_Price"] * ETHYLENE_DENSITY,
        "PGP": raw["PGP"]["F1_Price"] * PROPYLENE_DENSITY,
    }
    ngl = {
        "CAP": raw["CAP"]["F1_Price"],
        "BAP": raw["BAP"]["F1_Price"],
        "DAE": raw["DAE"]["F1_Price"],
    }
    prices = pd.DataFrame({**ngl, **petrochem})
    prices = prices.loc[START_DATE:END_DATE]
    prices = prices.dropna(how="all")
    return prices


def build_f2_prices(raw: dict) -> pd.DataFrame:
    """
    Return F2 prices ($/gal) for transaction-cost calculation on roll days.
    """
    petrochem = {
        "PCW": raw["PCW"]["F2_Price"] * ETHYLENE_DENSITY,
        "PGP": raw["PGP"]["F2_Price"] * PROPYLENE_DENSITY,
    }
    ngl = {
        "CAP": raw["CAP"]["F2_Price"],
        "BAP": raw["BAP"]["F2_Price"],
        "DAE": raw["DAE"]["F2_Price"],
    }
    f2 = pd.DataFrame({**ngl, **petrochem})
    f2 = f2.loc[START_DATE:END_DATE]
    return f2


def build_intra_spreads(prices: pd.DataFrame) -> pd.DataFrame:
    """Build all 5 intra-spreads ($/gal)."""
    spreads = {}
    for name, (pet_col, _, ngl_col) in INTRA_SPREAD_DEFS.items():
        s = prices[pet_col] - prices[ngl_col]
        spreads[name] = s
    df = pd.DataFrame(spreads)
    common = df.dropna(how="all").index
    return df.loc[common]


def get_leg_prices(name: str, prices: pd.DataFrame) -> tuple[str, str]:
    """
    Return (petrochem_col, ngl_col) column names in the `prices` DataFrame
    for a given intra-spread name.
    """
    pet_ticker, _, ngl_ticker = INTRA_SPREAD_DEFS[name]
    return pet_ticker, ngl_ticker


def align_all(raw: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Return (prices, f2_prices, spreads) all aligned on the common trading-day index.
    """
    prices = build_price_series(raw)
    f2     = build_f2_prices(raw)
    # Align to days where ALL 5 commodities have F1 prices
    common_idx = prices.dropna().index
    prices = prices.loc[common_idx]
    f2     = f2.reindex(common_idx)
    spreads = build_intra_spreads(prices)
    return prices, f2, spreads
