# Metals Risk Premia — Project State

## What This Project Is
LME Copper risk premia dashboard (Streamlit, `app.py`) + standalone signal scripts.
Pilot for a 10-metal extension. Supervised research with Prof. Ilia Bouchoev framework.
Git repo: https://github.com/kj2741-kj/Metals-Risk-Premia.git  branch: main

---

## Data Files (same directory as app.py)
| File | Contents |
|------|----------|
| `Metals Cash and 3M.xlsx` | LME cash + 3M prices for 10 metals |
| `Metals Futures Curve.csv` | LME Copper + other metals, F1–F27, multi-row header Excel |
| `LME_Copper_Rolling_F1_v2.csv` | F1_raw + F1_continuous (ratio back-adjusted) |

Curve sheet names: "Copper LME", "Copper CME", "ALuminium LME", "Lead LME", "Zinc LME", …
F1_continuous column is used for ALL PnL calculations regardless of signal source.

---

## Dashboard — app.py (3200+ lines, branch: main)

### Tabs
| Tab | Content |
|-----|---------|
| Tab 1 | Market Overview |
| Tab 2 | Term Structure (date picker) |
| Tab 3 | Cash vs 3M (Carry) |
| Tab 4 | Volume & Open Interest |
| Tab 5 | Copper LME-CME Spread |
| Tab 6 | Statistics |
| Tab 7 | Momentum Signals |
| Tab 8 | Carry Signals |
| Tab 9 | Value Signals ← **last added** |

### Module-level helpers (key functions)
- `_load_copper_f1_data()` — loads LME_Copper_Rolling_F1_v2.csv
- `_carry_raw_signal(curve_prices, cash_parsed, spec)` — raw carry series
- `_carry_cum_pnl(curve_prices, cash_parsed, f1c, spec)` — cum PnL
- `_CARRY_CMP_OPTIONS` — 29 carry variants dict
- `_value_raw_signal(curve_prices, f1r, spec)` — raw value series (deviation or reversal)
- `_value_cum_pnl(curve_prices, f1r, f1c, spec)` — cum PnL
- `_VALUE_CMP_OPTIONS` — 16 value variants dict

---

## Three Signal Strategies — Implementation Summary

### Momentum (Tab 7)
- MA Crossover: best params MA(35,43) Sharpe ~0.65 (Lag-1)
- CTA Baz-Granger: response function u = z·exp(−z²/4)/0.89
- Signal from F1_raw; PnL from F1_continuous
- Lag-1 is correct entry (momentum flips gradually, no urgency)
- Output: `momentum_output_v3/` Excel tradebooks

### Carry (Tab 8)
- V1 Roll Yield: (F1-F2)/F1, (F1-F3)/F1, (Cash-3M)/Cash
- V2 Long Slope: (Fj-Fk)/Fk for 10 tenor pairs F3-F15 through F12-F24
- V3 Z-score: 252d rolling standardisation of (F1-F2)/F1
- Same-Day dominates (level signal; flip day = large directional move)
- Best: V1 (F1-F2)/F1 Same-Day, Sharpe ~0.55
- Output: `carry_output/` Excel tradebooks

### Value (Tab 9) — LAST COMPLETED
- V1 MA Reversion on Fk: deviation = (Fk − MA_N)/MA_N; ±threshold band → +1/0/−1
- V2 Baz-Granger Reversal on F1_raw: reversal = F1_raw[t−N] − F1_raw[t]; sign → +1/−1
- Lag-1 is correct (no same-day urgency; Same-Day Sharpe = −0.5 to −1.8)
- Output: `value_output/` — 15 V1 Excel files (F1–F15) + 1 V2 Excel + summary CSV

---

## Value Signal Results (from value_signals.py, ±10% threshold, Lag-1)

### V1 MA Reversion — Best Contract per Lookback
| Lookback | Best Contract | Sharpe | Ann Ret% | Max DD% | % Flat |
|----------|--------------|--------|----------|---------|--------|
| 1yr | F12/F15 | 0.061 | +1.9% | −76% | 60% |
| 3yr | F15 | 0.085 | +2.2% | −68% | 43% |
| **5yr** | **F8** | **0.277** | **+7.1%** | −67% | 38% |
| 7yr | F14 | 0.193 | +4.3% | −61% | 34% |
| 10yr | F14 | 0.165 | +3.5% | −55% | 30% |

Mark Bogorad's reference (F12, 5yr): Sharpe = 0.184 — valid for energy, sub-optimal for copper.
Empirically optimal for copper: **F8 at 5yr lookback**.

### V2 Baz-Granger — All Lookbacks
| Lookback | Sharpe | Ann Ret% | Max DD% |
|----------|--------|----------|---------|
| 1yr | −0.355 | −8.8% | −206% |
| 3yr | +0.303 | +6.9% | −59% |
| 5yr | −0.137 | −2.8% | −93% |
| 7yr | +0.157 | +3.1% | −71% |
| **10yr** | **+0.512** | **+10.2%** | **−44%** |

Best overall signal: **V2 Baz-Granger 10yr, Sharpe = 0.512**.
Non-monotonic by lookback (5yr is a trap — catches middle of trends, not the end).
Large flat zone (35–60%) for V1 at ±10% — may need wider threshold for copper.

---

## Key Cross-Signal Findings (Bouchoev lens)
1. Same-Day vs Lag-1 matters differently per strategy:
   - Carry: Same-Day wins (flip-day captures the large directional move)
   - Momentum + Value: Lag-1 wins (signals evolve slowly)
2. F12 is not optimal for copper (it is for energy): copper's optimal V1 contract is F8
3. Value is strongly regime-conditional: most of the edge comes from COVID dislocation (2020→2021)
4. V2 5yr is a trap: non-monotonic, negative Sharpe despite long lookback

---

## IMMEDIATE NEXT STEP: Portfolio Construction

### Plan
Build `portfolio.py` (standalone script) + Tab 10 in app.py.

### EW Portfolio (equal-weight, start here)
```
Port_pos[t] = (1/3) × mom_pos[t]  +  (1/3) × carry_pos[t]  +  (1/3) × value_pos[t]
```
Where:
- mom_pos = MA(35,43) Lag-1 on F1_raw
- carry_pos = V1 (F1-F2)/F1 Same-Day
- value_pos = V1 F8 5yr Lag-1 OR V2 10yr Lag-1

### Critical Pre-Build Check
Compute pairwise correlations of the three position series before combining.
If all < 0.3: EW Sharpe ≈ √3 × avg individual = 0.7–0.8 (the diversification claim).
If momentum–carry correlation is high: investigate (both trend signals).

### Sub-period table to reproduce (matches Mark's paper)
| Period | Momentum | Carry | Value | EW Portfolio |
|--------|----------|-------|-------|--------------|
| Full (2006–2025) | ~0.65 | ~0.55 | ~0.28 | TBD |
| Pre-2022 | ~0.80 | ~0.70 | ~0.08 | TBD |
| Post-2022 | ~0.25 | −0.30 | ~0.50 | TBD |

### Tab 10 sections (planned)
1. Signal correlation matrix (heatmap)
2. EW portfolio live badge (current combined position: −1 to +1)
3. 16 metric cards (EW portfolio performance)
4. Signal agreement panel (% days all 3 agree / 2 agree / split)
5. Rolling Sharpe — EW vs individual signals
6. Sub-period performance table (Pre/Post 2022)
7. Cumulative PnL — EW vs each individual signal
8. Annual PnL bars
9. Position decomposition stack chart (Mom / Carry / Value contribution each day)
10. Risk parity variant (optional: inv-vol weighted)

---

## Standalone Scripts
| Script | Purpose | Output |
|--------|---------|--------|
| `momentum_signals.py` | MA + CTA tradebooks | `momentum_output_v3/` |
| `carry_signals.py` | Carry tradebooks | `carry_output/` |
| `value_signals.py` | Value tradebooks (V1 F1–F15, V2) | `value_output/` |
| `portfolio.py` | **TO BE BUILT** | `portfolio_output/` |

---

## Commit History (recent)
- `29b7723` — Add Tab 9 (Value Signals) and value_signals.py
- `a54abf6` — Redesign: commodity terminal color scheme (copper/charcoal palette)
