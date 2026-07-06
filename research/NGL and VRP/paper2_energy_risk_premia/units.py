"""
Unit conversions: all prices → $/bbl equivalent (Appendix E).

Raw Bloomberg units:
  CL, CO   : $/bbl          → multiply by 1.0
  XB, HO   : ¢/gal          → multiply by 42/100 = 0.42
  NG       : $/MMBtu         → multiply by 5.8  (crude energy equiv.)
  QS       : $/metric ton    → multiply by 1/7.45 (gasoil density)
  CAP/BAP/DAE : $/gal        → multiply by 42
"""

import pandas as pd
from config import INSTRUMENTS, BBL_CONVERSIONS


def to_usd_bbl(series: pd.Series, ticker: str) -> pd.Series:
    """Convert a single price series to $/bbl."""
    unit = INSTRUMENTS[ticker]["unit"]
    factor = BBL_CONVERSIONS[unit]
    return (series * factor).rename(ticker)


def convert_all(
    cont: dict[str, tuple[pd.Series, pd.Series]],
) -> dict[str, tuple[pd.Series, pd.Series]]:
    """
    Apply $/bbl conversion to every (adjusted, unadjusted) pair.
    Returns same structure with converted prices.
    """
    out = {}
    for ticker, (adj, unadj) in cont.items():
        out[ticker] = (to_usd_bbl(adj, ticker), to_usd_bbl(unadj, ticker))
    return out
