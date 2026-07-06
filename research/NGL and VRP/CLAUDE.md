# NGL Optionality & Energy Risk Premia — Replication Projects

Two independent Python projects replicating Mark Bogorad's papers. Both are fully functional and producing results as of the last session.

---

## Directory Layout

```
NGL and VRP/
├── CLAUDE.md                          ← this file
├── NGL_PAPER1.xlsx                    ← NGL swap prices (CAP/BAP/DAE/PCW/PGP, F1–F27)
├── NG_QS_Futures.xlsx                 ← Henry Hub NG + ICE Gasoil QS futures
├── Crack_futures_data.xlsx            ← CL/CO/XB/HO futures (5-col per contract)
├── expiry_calendars_20260526.xlsx     ← Roll calendar — ONLY has CL (276 rows); all other
│                                        EXPIRY_TICKERS fall back to EOM roll dates
├── paper1_ngl_optionality/            ← Paper 1 project
└── paper2_energy_risk_premia/         ← Paper 2 project
```

---

## Paper 1 — "Evaluating Trading Real Optionality Between NGLs and Petrochemicals"

### Sample Period
`2011-07-07` to `2025-02-24` (3,431 trading days)

### Key Files
| File | Purpose |
|------|---------|
| `config.py` | All constants: grids, densities, regime dates |
| `data_loader.py` | Loads NGL_PAPER1.xlsx; 3-col stride (Price/Volume/Open); F1 at col 1 |
| `spreads.py` | Builds intra/inter crack spreads; PCW in $/lb × 4.91 → $/gal |
| `cointegration.py` | Engle-Granger pairwise-complete (fixed — no longer all-5-complete) |
| `strategy.py` | Signal generation (lagged 1 day), EOM roll cost, Sharpe/MaxDD/RoD |
| `optimize.py` | Grid search over LOOKBACK_GRID × THRESHOLD_GRID |
| `plots.py` | Figures 2-6 + cointegration p-value bar chart |
| `main.py` | Full orchestration |

### Current Results (last run)
**Intra-spreads (optimised Sharpe):**
| Spread | Sharpe | Best n | Best eps | Paper Target |
|--------|--------|--------|----------|-------------|
| Ethane-Ethylene | 0.298 | 120 | 0.90 | ~0.30 |
| Propane-Ethylene | 0.495 | 200 | 0.60 | ~0.47 |
| Propane-Propylene | 0.460 | 250 | 0.90 | 0.47 ✓ |
| Butane-Ethylene | 0.468 | 200 | 0.60 | 0.56 (marginal gap) |
| Butane-Propylene | 0.482 | 200 | 0.90 | 0.46 ✓ |

**Inter-spreads (optimised Sharpe):**
| Pair | Sharpe | Paper Target |
|------|--------|-------------|
| EE vs PP | 0.630 | 0.64 ✓ |
| PE vs PP | 0.530 | ~0.53 |
| PE vs BE | 0.356 | 0.68 ← gap |
| PP vs BE | 0.408 | ~0.41 |
| BE vs BP | 0.529 | ~0.53 |

**Cointegration (Engle-Granger, 5% threshold):**
- PP (p=0.001) ✓ cointegrated — matches paper
- BP (p=0.002) ✓ cointegrated — matches paper
- BE (p=0.113) NOT cointegrated — paper says 0.047 (cointegrated)
- EE (p=0.259), PE (p=0.241) both non-cointegrated — matches paper

**Known open issue:** BE cointegration p=0.113 vs paper's 0.047. All 5 tickers have identical date coverage (3,433 obs each from 2011-07-07), so pairwise fix cannot help. The discrepancy is a data-vintage difference in the PCW or DAE Bloomberg series. The PE vs BE inter-spread Sharpe gap (0.36 vs 0.68) is likely downstream of this.

### Run Command
```
cd paper1_ngl_optionality
python main.py
```
Outputs to `output/` (charts + CSV performance tables).

---

## Paper 2 — "Risk Premia in Diversified Energy Portfolios"

### Sample Periods
- Data loaded from: `2010-01-04` (for 10yr value MA initialisation)
- Formal backtest: `2015-01-02` to `2025-12-01`
- Sub-sample 1 (Pre-2022): `2015-01-02` to `2021-12-31`
- Sub-sample 2 (Post-2022): `2022-01-02` to `2025-12-01`

### Key Files
| File | Purpose |
|------|---------|
| `config.py` | All constants including exact periods above |
| `data_loader.py` | 3-col loader for NGL/NG/QS; 5-col for CL/CO/XB/HO |
| `rolling.py` | Ratio backward adjustment; EOM fallback for tickers not in calendar |
| `units.py` | Convert raw prices to $/bbl (cps_gal×0.42, usd_gal×42, usd_mmbtu×5.8, usd_mt÷7.45) |
| `signals/momentum.py` | MA crossovers (1,5)(5,20)(10,60) in trading days; lagged 1 day |
| `signals/carry.py` | F4 vs F15; +1=backwardation, −1=contango; lagged 1 day |
| `signals/value.py` | 10yr (2520 TD) rolling MA on F12; ±10% threshold; lagged 1 day |
| `signals/stat_arb.py` | 8 spread pairs, 20-day MA, 10% threshold on F12 prices |
| `portfolio/allocation.py` | EW = plain 1/N (NO vol-scaling); DRP = inv-variance EWMA+LW |
| `portfolio/backtest.py` | IR = AnnRet/AnnVol; sub-sample periods defined here |
| `main.py` | Full orchestration |
| `print_results.py` | Standalone results summary script (does NOT re-run backtest) |

### Current Results (last run)
| Strategy | Full IR | Pre-2022 IR | Post-2022 IR |
|----------|---------|-------------|--------------|
| Momentum EW | +0.353 | +0.500 | +0.065 |
| Momentum DRP | +0.444 | +0.665 | +0.021 |
| Carry EW | −0.003 | +0.338 | −0.601 |
| Carry DRP | +0.153 | +0.595 | −0.611 |
| Value EW | +0.439 | +0.232 | +0.929 |
| Value DRP | +0.422 | +0.215 | +0.902 |
| StatArb EW | +0.267 | +0.371 | +0.088 |

All qualitative patterns match paper: momentum pre>post, carry strong pre-2022 then breaks down, value very strong post-2022 (COVID deviation + reversion), stat-arb consistently positive.

### Run Command
```
cd paper2_energy_risk_premia
python main.py
```
Outputs to `output/` (charts + CSV tables). For just the results table without re-running:
```
python print_results.py
```

---

## Critical Bugs Fixed (do not revert)

### Paper 1
1. **Unicode print** — `main.py` uses ASCII `to` / `eps` not `→` / `ε` (Windows cp1252 can't encode them)

### Paper 2
1. **`ROLL_COST_BPS = 0.002`** — was `20` (integer basis points used directly as a fraction → −2000%/roll day destroying all CL/NGL returns). Now correctly `0.002` (0.20% decimal).

2. **EOM fallback for roll dates** — `rolling.py:_get_roll_dates()` falls back to EOM dates when a ticker is not in the expiry calendar. Only CL has calendar data; CO/XB/HO/NG/QS all use EOM. Without this they had zero roll costs.

3. **Ratio backward adjustment** — `rolling.py` multiplies prior prices by `F2/F1` on roll dates (not additive subtraction). Additive caused CL to reach −$261,858 over 15 years of contango rolls.

4. **EW portfolio no vol-scaling** — `portfolio/allocation.py:equal_weight_portfolio` uses plain `(sig/n × ret).sum()`. An earlier version divided by annualised EWMA vol (~0.3), creating ~3.3× leverage per instrument.

5. **Stat-arb returns in log-return units** — `main.py` computes stat-arb returns as `signal × (r_leg1 − r_leg2)` using log returns, not `signal × Δspread_$/bbl`. The latter inflated stat-arb to 20× the magnitude of other strategies.

6. **Initialisation window** — `build_all_returns` / signal functions use `DATA_START="2010-01-04"` for full 10yr history before the 2015 backtest start.

---

## What Is Still Open

1. **Paper 1 BE cointegration** — p=0.113 vs paper's 0.047; likely a data vintage issue in PCW or DAE. Investigate by testing different sub-periods or checking for any data gaps/outliers in the DAE series around 2011-2013.

2. **Paper 1 PE vs BE inter Sharpe** — 0.356 vs paper's 0.68. Downstream of BE cointegration issue. Once BE data is reconciled, re-optimise this inter-spread.

3. **Paper 2 exact IR numbers** — The paper's precise numerical table values are not available for a line-by-line comparison (only the qualitative patterns were verified). If you obtain the paper's exact Table values, run `python print_results.py` to compare.

4. **IBD (Isobutane)** — Explicitly skipped for now; revisit once both main projects are finalised.

5. **Git / packaging** — Both projects are ready to be pushed to separate GitHub repos.

---

## Notes on Data

- **NGL_PAPER1.xlsx**: row 1 skipped, row 2 = contract labels, row 3 = sub-labels, row 4+ = data; 3-col stride (Price/Volume/Open Interest); F1 at column index 1 per contract.
- **Crack_futures_data.xlsx**: 5-col stride (Price/Volume/Open/High/Low).
- **NG_QS_Futures.xlsx**: 3-col stride like NGL file.
- **Expiry calendar**: Only CL rows present. All other exchange-traded tickers use EOM fallback.
- Bloomberg units: NGLs in $/gal, PCW/PGP in $/lb, XB/HO in ¢/gal, QS in $/mt, NG in $/MMBtu, CL/CO in $/bbl.
