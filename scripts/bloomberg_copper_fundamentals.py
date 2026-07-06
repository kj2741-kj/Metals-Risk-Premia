"""
bloomberg_copper_fundamentals.py
=================================
Downloads fundamental factor data for LME Copper signal research.
Aligned to project date window: 2005-01-01 → today.

Prerequisites
-------------
  pip install pdblp pandas
  Bloomberg Terminal must be running on the same machine (port 8194).

Output
------
  fundamental_data/
  ├── 01_copper_physical/       warehouse stocks, cancelled warrants
  ├── 02_demand_growth/         IP, PMI, BDI, housing, GDP
  ├── 03_currencies/            DXY + all producer/consumer FX
  ├── 04_monetary_rates/        Fed Funds, SOFR, 10Y, M2 (China/US/EU)
  ├── 05_cross_commodity/       energy, LME metals, commodity indices
  ├── 06_financial_sentiment/   VIX, MOVE, CSI300, COT positioning
  ├── 07_trade_flows/           China copper imports/exports, Yangshan premium
  ├── 08_geopolitical_risk/     GPR index, EPU indices
  ├── 09_energy_transition/     clean-energy ETFs, battery metals
  └── 10_country_producers/     GDP + FX + production for each major producer

Ticker confidence levels
------------------------
  ✓  HIGH   — confirmed standard Bloomberg ticker, very unlikely to fail
  ~  MEDIUM — standard format, high probability correct, minor doubt
  ⚠  VERIFY — best-guess ticker; confirm in Bloomberg (SECF / SRCH) before use

The script prints [OK] / [FAIL] for every ticker. Failed tickers are
printed again in a summary at the end. For ⚠ VERIFY items that fail,
search Bloomberg by keyword in the SECF / ECST / CMD screens.

External sources for data NOT available on Bloomberg:
  GPR   : https://www.matteoiacoviello.com/gpr.htm         (monthly)
  EPU   : https://www.policyuncertainty.com/global_monthly.html
  ICSG  : https://www.icsg.org  (global copper supply-demand balance)
"""

import os
import sys
import time
import pandas as pd
from datetime import datetime

try:
    import pdblp
except ImportError:
    sys.exit("ERROR: Run  pip install pdblp  then restart.")

# ─── CONFIG ───────────────────────────────────────────────────────────────────
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)

START     = "20050101"                          # 1-year buffer before project start
END       = datetime.today().strftime("%Y%m%d")
OUT_ROOT  = os.path.join(_REPO_ROOT, "research", "Fundamentals Cu", "fundamental_data")
FIELD     = "PX_LAST"                           # works for prices, indices, economic series

os.makedirs(OUT_ROOT, exist_ok=True)

print("Connecting to Bloomberg Terminal (port 8194)…")
con = pdblp.BCon(debug=False, port=8194, timeout=15000)
con.start()
print(f"Connected. Fetching {START} → {END}\n")

# ─── TRACKING ─────────────────────────────────────────────────────────────────
results = []   # (domain, ticker, name, status, note)


# ─── HELPERS ──────────────────────────────────────────────────────────────────
def _download(ticker, name, domain, periodicity, field=FIELD):
    """Core download + save routine. Returns True on success."""
    folder = os.path.join(OUT_ROOT, domain)
    os.makedirs(folder, exist_ok=True)
    try:
        elms = {"periodicitySelection": periodicity}
        if periodicity == "DAILY":
            elms["nonTradingDayFillMethod"]  = "PREVIOUS_VALUE"
            elms["nonTradingDayFillOption"]  = "NON_TRADING_WEEKDAYS"
        df = con.bdh(ticker, field, START, END, elms=elms)
        if df is None or df.empty:
            raise ValueError("Empty response")
        df.columns = [name]
        df.to_csv(os.path.join(folder, f"{name}.csv"))
        print(f"  [OK]   {name:<55} {ticker}")
        results.append((domain, ticker, name, "OK", ""))
        return True
    except Exception as exc:
        print(f"  [FAIL] {name:<55} {ticker}  →  {exc}")
        results.append((domain, ticker, name, "FAIL", str(exc)))
        return False


def fetch_daily(tickers: dict, domain: str):
    print(f"\n{'─'*70}")
    print(f"  DOMAIN: {domain}  [DAILY]")
    print(f"{'─'*70}")
    for ticker, name in tickers.items():
        _download(ticker, name, domain, "DAILY")
        time.sleep(0.15)


def fetch_monthly(tickers: dict, domain: str):
    print(f"\n{'─'*70}")
    print(f"  DOMAIN: {domain}  [MONTHLY]")
    print(f"{'─'*70}")
    for ticker, name in tickers.items():
        _download(ticker, name, domain, "MONTHLY")
        time.sleep(0.15)


def fetch_quarterly(tickers: dict, domain: str):
    print(f"\n{'─'*70}")
    print(f"  DOMAIN: {domain}  [QUARTERLY]")
    print(f"{'─'*70}")
    for ticker, name in tickers.items():
        _download(ticker, name, domain, "QUARTERLY")
        time.sleep(0.15)


def fetch_weekly(tickers: dict, domain: str):
    print(f"\n{'─'*70}")
    print(f"  DOMAIN: {domain}  [WEEKLY]")
    print(f"{'─'*70}")
    for ticker, name in tickers.items():
        _download(ticker, name, domain, "WEEKLY")
        time.sleep(0.15)


# ══════════════════════════════════════════════════════════════════════════════
#  1. COPPER PHYSICAL MARKET
#     Source: Guzmán & Silva (2018) — SC_EX (total exchange stocks) is
#     one of five endogenous copper-market variables in their VAR.
#     Galán-Gutiérrez (2022) — LME stock level drives contango/backwardation.
# ══════════════════════════════════════════════════════════════════════════════

copper_physical_daily = {
    # ── Exchange Warehouse Stocks ─────────────────────────────────────────────
    "LCASYSTO Index" : "LME_Copper_Stocks_Total_kt",          # ⚠ VERIFY
    "LCAIFCNW Index" : "LME_Copper_CancelledWarrants_kt",     # ⚠ VERIFY
    "CSCOPRWS Index" : "COMEX_Copper_Stocks_kt",              # ⚠ VERIFY
    "SHFECUON Index" : "SHFE_Copper_Stocks_kt",               # ⚠ VERIFY
}

fetch_daily(copper_physical_daily, "01_copper_physical")


# ══════════════════════════════════════════════════════════════════════════════
#  2. DEMAND & GROWTH INDICATORS
#     Source: Guzmán & Silva (2018) Table 1 — BDI (bdi), S&P 500 (sp500),
#     World IP (ip_world_2) are among their 11 confirmed exogenous variables.
#     Miller-Martinez (2026) OECD/GRI: aggregate demand is the DOMINANT
#     driver of copper price shocks (demand > supply in long-run SVAR).
# ══════════════════════════════════════════════════════════════════════════════

demand_daily = {
    "BDIY Index"    : "Baltic_Dry_Index",          # ✓  Guzmán (2018) bdi
    "SPX Index"     : "SP500",                     # ✓  Guzmán (2018) sp500
    "MXEF Index"    : "MSCI_EM_Equity_Index",      # ✓  EM demand proxy
    "SHSZ300 Index" : "China_CSI300_Equity",       # ✓  China risk appetite
}

demand_monthly = {
    # ── China demand indicators ───────────────────────────────────────────────
    "CHVAIOY Index"  : "China_IndProd_YoY",              # ~  value-added industry output
    "CPMINDX Index"  : "China_PMI_Mfg_NBS",             # ~  NBS Manufacturing PMI
    "CXNPMIM Index"  : "China_PMI_Mfg_Caixin",          # ~  Caixin Manufacturing PMI
    "CNFAIGYHM Index": "China_FixedAssetInvest_YoY",    # ⚠ VERIFY — FAI growth YoY (month)

    # ── US demand indicators ──────────────────────────────────────────────────
    "NAPMPMI Index"  : "US_ISM_Manufacturing_PMI",       # ✓  Institute for Supply Mgmt
    "NHSPSTOT Index" : "US_Housing_Starts_Total",        # ~  wiring/plumbing copper demand

    # ── Global demand indicators ──────────────────────────────────────────────
    "MPMIGMAN Index" : "JPM_Global_PMI_Manufacturing",   # ⚠ VERIFY — JPMorgan Global PMI
    "CPBLWDIP Index" : "CPB_World_Industrial_Prod",      # ⚠ VERIFY — Guzmán ip_world_2
    "EUIPINYY Index" : "Eurozone_IndProd_YoY",           # ⚠ VERIFY
}

fetch_daily(demand_daily,   "02_demand_growth")
fetch_monthly(demand_monthly, "02_demand_growth")


# ══════════════════════════════════════════════════════════════════════════════
#  3. CURRENCIES
#     Source: Guzmán & Silva (2018) Table 1 — TWUSD_PRICE_ADJ (Dollar Index)
#     is statistically significant in period III VAR results (Table 4).
#     Ben-Salha (2025): post-financialisation (2004+), currency channels
#     amplify spillovers; copper strongly transmits to other metals.
#
#     Breakdown:
#       DXY           → overall USD strength (inverse copper relationship)
#       CNY           → China consumer purchasing power
#       CLP / PEN     → producer cost floor (weaker = cheaper to mine)
#       AUD           → Australia + China economic cycle proxy
# ══════════════════════════════════════════════════════════════════════════════

currencies_daily = {
    # ── Global dollar signal ──────────────────────────────────────────────────
    "DXY Curncy"    : "USD_Index_DXY",             # ✓  Guzmán TWUSD_PRICE_ADJ

    # ── Primary consumer FX ───────────────────────────────────────────────────
    "USDCNY Curncy" : "USDCNY_China",             # ✓  China = ~55% world demand
    "EURUSD Curncy" : "EURUSD",                    # ✓  Europe industrial consumer
    "USDJPY Curncy" : "USDJPY_Japan",             # ✓  Japan industrial consumer

    # ── Primary producer FX ───────────────────────────────────────────────────
    "USDCLP Curncy" : "USDCLP_Chile",             # ✓  Chile ~27% world mine supply
    "USDPEN Curncy" : "USDPEN_Peru",              # ✓  Peru  ~10% world mine supply
    "AUDUSD Curncy" : "AUDUSD_Australia",          # ✓  Australia + China growth proxy
    "USDZAR Curncy" : "USDZAR_SouthAfrica",        # ✓
    "USDBRL Curncy" : "USDBRL_Brazil",             # ✓
    "USDMXN Curncy" : "USDMXN_Mexico",            # ✓
    "USDRUB Curncy" : "USDRUB_Russia",             # ✓
    "USDIDR Curncy" : "USDIDR_Indonesia",          # ✓
    "USDKZT Curncy" : "USDKZT_Kazakhstan",        # ✓
    "USDZMW Curncy" : "USDZMW_Zambia",            # ⚠ VERIFY — sparse history expected
    "USDCDF Curncy" : "USDCDF_DRC_CongoFranc",   # ⚠ VERIFY — very sparse
}

fetch_daily(currencies_daily, "03_currencies")


# ══════════════════════════════════════════════════════════════════════════════
#  4. MONETARY POLICY & INTEREST RATES
#     Source: Guzmán & Silva (2018) Table 1 — Three M2 series (China, US, EU)
#     and real LIBOR (RNIR_LIBOR) all in their confirmed exogenous variable set.
#     Key finding (Tables 2-4): China M2 is significant across ALL three
#     sub-periods — the single most robust non-fundamental driver.
#     Frankel (2014) [cited in Guzmán]: real interest rate has negative
#     elasticity to commodity prices (rate↑ → copper↓).
# ══════════════════════════════════════════════════════════════════════════════

rates_daily = {
    "FDTR Index"    : "US_FedFunds_Target_Rate",   # ✓
    "USGG10YR Index": "US_10Y_Treasury_Yield",     # ✓
    "USGGBE10 Index": "US_10Y_Breakeven_Inflation",# ✓  real yield = 10Y - breakeven
    "SOFRRATE Index": "SOFR_Overnight",             # ⚠ VERIFY — replaced LIBOR
}

monetary_monthly = {
    # ── M2 money supply — Guzmán rm2_china / rm2_usa / rm2_eu ───────────────
    "CHMS2YY Index"  : "China_M2_YoY",            # ⚠ VERIFY — try CNMS2YOY if fails
    "M2 Index"       : "US_M2_Level",             # ⚠ VERIFY — use level, compute YoY
    "EUMSM2YY Index" : "EU_M2_YoY",               # ⚠ VERIFY

    # ── China short rate ──────────────────────────────────────────────────────
    "CHIBOR07 Index" : "China_7D_Repo_Rate",       # ⚠ VERIFY — try CNRR7D if fails

    # ── Inflation deflators (needed to compute real series per Guzmán) ────────
    "CPI Index"      : "US_CPI_Level",             # ⚠ VERIFY — US CPI index level
    "CHCPI YOY Index": "China_CPI_YoY",            # ⚠ VERIFY
}

fetch_daily(rates_daily,       "04_monetary_rates")
fetch_monthly(monetary_monthly, "04_monetary_rates")


# ══════════════════════════════════════════════════════════════════════════════
#  5. CROSS-COMMODITY PRICES
#     Source: Guzmán & Silva (2018) Table 1 — RPOIL_WTI (real WTI) and
#     GSCI are confirmed exogenous variables.
#     Wang et al. (2026): Natural gas → copper via energy/smelting cost
#     channel (especially significant for China smelting activity).
#     Ben-Salha (2025): Copper + aluminum are co-transmitters; zinc and
#     lead are net receivers from the copper system.
# ══════════════════════════════════════════════════════════════════════════════

commodities_daily = {
    # ── Energy: production cost channel ──────────────────────────────────────
    "CL1 Comdty"     : "WTI_Crude_Front",          # ✓  Guzmán RPOIL_WTI
    "CO1 Comdty"     : "Brent_Crude_Front",        # ✓
    "NG1 Comdty"     : "NatGas_HenryHub_Front",    # ✓  Wang et al. (2026)
    "NEWCSTLD Index" : "Coal_Newcastle_Spot",       # ⚠ VERIFY — smelting energy cost

    # ── Commodity indices ─────────────────────────────────────────────────────
    "SPGSCI Index"   : "GSCI_Commodity_Index",     # ✓  Guzmán GSCI
    "BCOM Index"     : "Bloomberg_Commodity_Index", # ✓

    # ── LME base metals ───────────────────────────────────────────────────────
    "LMAHDS03 Comdty": "LME_Aluminum_3M",          # ~  Ben-Salha: Cu-Al co-transmitters
    "LMZSDS03 Comdty": "LME_Zinc_3M",             # ⚠ VERIFY
    "LMNIDS03 Comdty": "LME_Nickel_3M",           # ⚠ VERIFY

    # ── Safe haven / risk-off ─────────────────────────────────────────────────
    "XAUUSD Curncy"  : "Gold_Spot_USD",            # ✓
}

fetch_daily(commodities_daily, "05_cross_commodity")


# ══════════════════════════════════════════════════════════════════════════════
#  6. FINANCIAL MARKET & RISK SENTIMENT
#     Source: Guzmán & Silva (2018) Table 1 — VIX_P50 (average VIX) is a
#     confirmed exogenous variable, significant in Period III (Table 4).
#     COM_LONG (COMEX net long positions) is an endogenous variable in their
#     system — speculative positioning affects copper price formation.
#     Galán-Gutiérrez (2022): copper is the second most financialised LME
#     metal; excess speculation index (Working-T) drove 2003-2008 boom.
# ══════════════════════════════════════════════════════════════════════════════

sentiment_daily = {
    "VIX Index"     : "VIX_CBOE_Volatility",       # ✓  Guzmán VIX_P50
    "MOVE Index"    : "MOVE_BondVolatility",        # ✓  bond market stress
}

# CFTC Commitments of Traders — weekly release (Friday)
# Copper = HG on COMEX. Net non-commercial = speculative net long.
# ⚠ VERIFY all COT tickers — Bloomberg CFTC data format varies by version.
# Alternative: download directly from CFTC at cftc.gov/dea/futures/deacmesf.htm
cot_weekly = {
    "NHLCNCUCU Index": "COMEX_Copper_Net_NonCommercial_Longs",  # ⚠ VERIFY
    "NHLCNU Index"   : "COMEX_Copper_OpenInterest",             # ⚠ VERIFY
}

fetch_daily(sentiment_daily, "06_financial_sentiment")
fetch_weekly(cot_weekly,     "06_financial_sentiment")


# ══════════════════════════════════════════════════════════════════════════════
#  7. TRADE FLOW DATA — CHINA & KEY PRODUCERS
#     China customs separate refined copper (cathode) from ores/concentrates:
#       Refined  → measures final demand absorption in China
#       Concs    → measures smelter feedstock demand (leads refined by 2-3m)
#       Scrap    → substitution signal (scrap↑ = less refined demand)
#     Yangshan Premium: daily implied cost of importing copper into China.
#       Premium↑ = Chinese buyers willing to pay more = bullish demand signal.
# ══════════════════════════════════════════════════════════════════════════════

# Yangshan premium is a daily observable — critical real-time signal
yangshan_daily = {
    "CCUPREMD Index" : "Yangshan_Copper_Import_Premium_USD_t",  # ⚠ VERIFY
}

trade_monthly = {
    # ── China imports ─────────────────────────────────────────────────────────
    "CNIMCURF Index" : "China_Copper_Imports_Refined_kt",          # ⚠ VERIFY
    "CNIMCUOC Index" : "China_Copper_Imports_OreConcentrate_kt",   # ⚠ VERIFY
    "CNIMCUSC Index" : "China_Copper_Imports_Scrap_kt",            # ⚠ VERIFY

    # ── China exports ─────────────────────────────────────────────────────────
    "CNEXCURM Index" : "China_Copper_Exports_Refined_kt",          # ⚠ VERIFY

    # ── Country-level copper export values ────────────────────────────────────
    "CHLECOEX Index" : "Chile_Copper_Exports_Value_USD",           # ⚠ VERIFY
    "PECOEXCU Index" : "Peru_Copper_Exports_Value_USD",            # ⚠ VERIFY
}

fetch_daily(yangshan_daily,  "07_trade_flows")
fetch_monthly(trade_monthly, "07_trade_flows")


# ══════════════════════════════════════════════════════════════════════════════
#  8. GEOPOLITICAL RISK & POLICY UNCERTAINTY
#     Source: Ben-Salha et al. (2025) — GPR and EPU are confirmed drivers
#     of all non-ferrous metal prices; copper is the most prominent transmitter
#     and receiver of GPR shocks in the TVP-VAR system.
#     Miller-Martinez (2026): export restrictions on TCMs rose 5× in past
#     decade; trade announcements move nickel ±8%, copper ±5%.
#
#  ⚠ NOTE: Caldara-Iacoviello GPR Index may NOT be in Bloomberg as a standard
#    ticker. If the download fails:
#      → Download monthly CSV from: https://www.matteoiacoviello.com/gpr.htm
#      → EPU from: https://www.policyuncertainty.com/global_monthly.html
# ══════════════════════════════════════════════════════════════════════════════

geopolitical_monthly = {
    "GEOPOLRSK Index" : "GPR_Geopolitical_Risk_Index",     # ⚠ VERIFY — may not be in BBG
    "BBEPUGLOB Index" : "EPU_Global_Policy_Uncertainty",   # ⚠ VERIFY
    "USEPUINDX Index" : "EPU_United_States",               # ⚠ VERIFY
    "CHNEPUIND Index" : "EPU_China",                       # ⚠ VERIFY
}

fetch_monthly(geopolitical_monthly, "08_geopolitical_risk")


# ══════════════════════════════════════════════════════════════════════════════
#  9. ENERGY TRANSITION DEMAND INDICATORS
#     Source: Miller-Martinez (2026) OECD/GRI — 6 key indicators of
#     structural demand shift for transition-critical metals.
#     Under IEA Net Zero scenario: copper demand expected to ~double by 2040.
#     EV alone = ~83 kg Cu/vehicle (vs ~23 kg ICE); solar = ~5 t Cu/MW.
#
#  ETFs used as proxies where direct series unavailable on Bloomberg:
#    ICLN — iShares Global Clean Energy ETF (monthly AUM as demand signal)
#    TAN  — Invesco Solar ETF
#    QCLN — First Trust NASDAQ Clean Edge ETF
# ══════════════════════════════════════════════════════════════════════════════

energy_transition_daily = {
    "ICLN US Equity" : "iShares_Global_CleanEnergy_ETF",  # ✓  proxy for global CE capex
    "TAN US Equity"  : "Invesco_Solar_ETF",               # ✓  solar buildout proxy
    "QCLN US Equity" : "FirstTrust_CleanEdge_ETF",        # ✓
}

energy_transition_monthly = {
    # ── Battery metals (demand multipliers for green transition) ──────────────
    "LITHIUMSP Index" : "Lithium_Carbonate_Price",         # ⚠ VERIFY
    "COBALTMP Index"  : "Cobalt_Metal_Price",              # ⚠ VERIFY

    # ── EV sales (China) ──────────────────────────────────────────────────────
    "CNNEVSSAM Index" : "China_NEV_Sales_Monthly",         # ⚠ VERIFY
}

fetch_daily(energy_transition_daily,   "09_energy_transition")
fetch_monthly(energy_transition_monthly, "09_energy_transition")


# ══════════════════════════════════════════════════════════════════════════════
#  10. COUNTRY-LEVEL DATA: MAJOR COPPER CONCENTRATE PRODUCERS
#      Per your request: GDP + currency + copper mine production for each
#      major producing nation.
#
#  Ranked by approximate 2023 mine production share:
#    Chile ~27% | Peru ~10% | DRC ~10% | China ~8% | Australia ~5%
#    Russia ~4% | Zambia ~3% | Mexico ~3% | Indonesia ~3% | Kazakhstan ~2%
#
#  GDP downloaded quarterly (only frequency available for most countries).
#  Currency and production data downloaded monthly.
#  All currencies are also in Section 3 (daily) for signal alignment.
# ══════════════════════════════════════════════════════════════════════════════

# 10A — Country GDP YoY (Quarterly)
country_gdp_quarterly = {
    # ── GDP YoY growth rate ───────────────────────────────────────────────────
    "CLGDPYOY Index" : "Chile_GDP_YoY",           # ⚠ VERIFY
    "PEGDPYOY Index" : "Peru_GDP_YoY",            # ⚠ VERIFY
    "CDGDPYOY Index" : "DRC_GDP_YoY",             # ⚠ VERIFY — sparse; DRC = Congo
    "CHGDPYOY Index" : "China_GDP_YoY",           # ~  fairly standard
    "AUGDPYOY Index" : "Australia_GDP_YoY",       # ~
    "RGGDPYOY Index" : "Russia_GDP_YoY",          # ⚠ VERIFY — try RUGDPYOY if fails
    "ZMGDPYOY Index" : "Zambia_GDP_YoY",          # ⚠ VERIFY — sparse
    "MXGDPYOY Index" : "Mexico_GDP_YoY",          # ~
    "IDGDPYOY Index" : "Indonesia_GDP_YoY",       # ~
    "KZGDPYOY Index" : "Kazakhstan_GDP_YoY",      # ⚠ VERIFY
}

fetch_quarterly(country_gdp_quarterly, "10_country_producers")

# 10B — Country Copper Mine Production (Monthly, tonnes)
# ⚠ ALL production tickers are best guesses — verify in Bloomberg ECST / CMD.
# If Bloomberg fails, source from: ICSG (icsg.org) or USGS Mineral Resources.
country_production_monthly = {
    "CHLECOPP Index" : "Chile_Copper_MineProduction_t",    # ⚠ VERIFY
    "PECOPPER Index" : "Peru_Copper_MineProduction_t",     # ⚠ VERIFY
    "AUCOPPRO Index" : "Australia_Copper_MineProduction_t",# ⚠ VERIFY
    "MXCOPPRO Index" : "Mexico_Copper_MineProduction_t",   # ⚠ VERIFY
    "IDCOPPRO Index" : "Indonesia_Copper_MineProduction_t",# ⚠ VERIFY
    "RUCOPPRO Index" : "Russia_Copper_MineProduction_t",   # ⚠ VERIFY
    "ZMCOPPRO Index" : "Zambia_Copper_MineProduction_t",   # ⚠ VERIFY
    "KZCOPPRO Index" : "Kazakhstan_Copper_MineProduction_t",# ⚠ VERIFY
}

fetch_monthly(country_production_monthly, "10_country_producers")

# 10C — Country currencies (daily for signal alignment)
# These are the same FX rates as Section 3 but saved separately here
# so the country_producers folder is self-contained for analysis.
country_fx_daily = {
    "USDCLP Curncy" : "Chile_USDCLP",      # ✓
    "USDPEN Curncy" : "Peru_USDPEN",        # ✓
    "AUDUSD Curncy" : "Australia_AUDUSD",   # ✓
    "USDCNY Curncy" : "China_USDCNY",       # ✓
    "USDRUB Curncy" : "Russia_USDRUB",      # ✓
    "USDZMW Curncy" : "Zambia_USDZMW",      # ⚠ VERIFY
    "USDMXN Curncy" : "Mexico_USDMXN",     # ✓
    "USDIDR Curncy" : "Indonesia_USDIDR",  # ✓
    "USDKZT Curncy" : "Kazakhstan_USDKZT", # ✓
    "USDCDF Curncy" : "DRC_USDCDF",        # ⚠ VERIFY — DRC Congo Franc, very sparse
}

fetch_daily(country_fx_daily, "10_country_producers")


# ══════════════════════════════════════════════════════════════════════════════
#  END — SUMMARY REPORT
# ══════════════════════════════════════════════════════════════════════════════

print(f"\n{'═'*70}")
print("  DOWNLOAD COMPLETE")
print(f"{'═'*70}")

ok_count   = sum(1 for r in results if r[3] == "OK")
fail_count = sum(1 for r in results if r[3] == "FAIL")
print(f"\n  Succeeded : {ok_count}")
print(f"  Failed    : {fail_count}")

# Save manifest CSV
manifest = pd.DataFrame(results, columns=["domain","ticker","filename","status","error"])
manifest.to_csv(os.path.join(OUT_ROOT, "manifest.csv"), index=False)
print(f"\n  Manifest  : {os.path.join(OUT_ROOT, 'manifest.csv')}")

# Print all failures grouped by domain
failed = [(r[0], r[1], r[2]) for r in results if r[3] == "FAIL"]
if failed:
    print(f"\n{'─'*70}")
    print("  FAILED TICKERS — verify in Bloomberg (SECF / ECST / CMD screen):")
    print(f"{'─'*70}")
    cur_domain = ""
    for domain, ticker, name in failed:
        if domain != cur_domain:
            print(f"\n  [{domain}]")
            cur_domain = domain
        print(f"    {ticker:<30}  ({name})")

print(f"\n{'─'*70}")
print("  External sources for tickers likely absent from Bloomberg:")
print("    GPR  : https://www.matteoiacoviello.com/gpr.htm")
print("    EPU  : https://www.policyuncertainty.com/global_monthly.html")
print("    ICSG : https://www.icsg.org  (global supply-demand balance)")
print(f"{'─'*70}\n")

con.stop()
