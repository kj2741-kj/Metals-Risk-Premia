"""
dashboard_home.py
==================
Project hub page: overview, methodology, headline findings, and
navigation into each asset-class dashboard. Pure landing page -- no
data loading, so it renders instantly regardless of how heavy the
other pages get.
"""

import streamlit as st

from dashboard_shared import COLORS, metric_card, section_header

st.markdown('<p class="main-title">⚙️ Metals Risk Premia</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="main-subtitle">Systematic Momentum, Carry &amp; Value risk premia across LME metals, '
    'energy, and refined products &mdash; supervised research with Prof. Ilia Bouchoev</p>',
    unsafe_allow_html=True,
)
st.divider()

section_header("WHAT THIS PROJECT IS")
st.markdown(
    """
A systematic framework that decomposes commodity futures returns into three economically distinct,
near-uncorrelated risk premia — **Momentum** (trend persistence), **Carry** (curve shape / roll yield),
and **Value** (mean-reversion to a long-run anchor) — then combines them into an equal-weight portfolio.
The central result on the metals pilot (Copper, Aluminium): three orthogonal sleeves of similar
stand-alone Sharpe combine into a portfolio whose risk-adjusted return materially exceeds any single
sleeve, while roughly halving drawdown.

Stage 1 (complete) validated the framework on LME Copper and Aluminium. Stage 2 (in progress) extends
the same three strategies — kept deliberately simple for now (no CTA-paper trend, no structural
Anchors, no walk-forward OOS yet) — to Oil & Energy, Precious Metals, and refined NGL products.
"""
)

section_header("METHODOLOGY (SAME CONVENTION ACROSS EVERY ASSET CLASS)")
c1, c2, c3 = st.columns(3)
with c1:
    metric_card("PnL Basis", "F1_continuous", unit="")
    st.caption("Ratio back-adjusted continuous front-month series. Signals read raw prices (F1_raw); "
               "PnL always realises on F1_continuous.")
with c2:
    metric_card("Sharpe Convention", "Active-Day", unit="")
    st.caption("Annualised mean/std of daily returns × √252, computed over days the strategy actually "
               "holds a position — flat days don't dilute the ratio.")
with c3:
    metric_card("Transaction Costs", "On F1_raw", unit="")
    st.caption("Round-trip cost = |Δposition| × (bps/10000/2) × F1_raw, charged at every position "
               "change on the real traded price, not the adjusted series.")

st.caption(
    "Execution timing: **Same-Day** = position(t) = signal(t−1) (shift-1). **Lag-1** = position(t) = "
    "signal(t−2) (shift-2, one extra day, no look-ahead either way). Which convention wins is "
    "strategy-specific and re-checked per asset class, not assumed."
)

section_header("HEADLINE FINDINGS (STAGE 1 — METALS)")
f1, f2, f3, f4 = st.columns(4)
with f1:
    metric_card("Copper EW Portfolio", "+0.73", unit=" Sharpe")
    st.caption("Net of 5bps, full sample 2006-2025. Best single sleeve (Momentum): +0.62.")
with f2:
    metric_card("Aluminium EW Portfolio", "+0.85", unit=" Sharpe")
    st.caption("Net of 5bps, full sample 2006-2026. Best single sleeve (Carry): +0.64.")
with f3:
    metric_card("Optimal Config", "Metal-Specific", unit="")
    st.caption("Copper wants a faster trend + curve-momentum carry; Aluminium wants a slower trend + "
               "mean-reverting z-score carry. No one-size template.")
with f4:
    metric_card("Diversification", "Corr < 0.25", unit="")
    st.caption("Momentum-Carry-Value position correlations mostly modest to negative — the combination, "
               "not any one sleeve, is the product.")

st.divider()
section_header("EXPLORE THE DASHBOARDS")
st.caption("Full detail — live signals, parameter controls, equity curves, rolling Sharpe, performance "
           "metrics — lives in each asset-class page below.")

nav1, nav2, nav3, nav4 = st.columns(4)
with nav1:
    st.page_link("dashboard_metals.py", label="**Metals**  →", icon="⚙️")
    st.caption("LME Copper & Aluminium. Momentum, Carry, Value, Portfolio (10 tabs) — complete, Stage 1.")
with nav2:
    st.caption("**Energy**  →")
    st.caption("Oil & gas products. *Coming in Stage 2.*")
with nav3:
    st.caption("**Precious Metals**  →")
    st.caption("Gold, Silver, Platinum, Palladium, Copper-CME. *Coming in Stage 2.*")
with nav4:
    st.caption("**NGL / Refined Products**  →")
    st.caption("Propane, Butane, Ethane, etc. *Coming in Stage 2 — pending data verification.*")

st.divider()
st.caption(
    "Data: LME futures curves (F1–F27), NYMEX/ICE/COMEX futures curves, LME Cash & 3M prices. "
    "Research prototype for academic purposes — in-sample backtests unless stated otherwise. Not "
    "investment advice."
)
