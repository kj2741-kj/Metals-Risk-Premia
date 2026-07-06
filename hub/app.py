"""
Metals Risk Premia - Project Hub
==================================
Standalone landing dashboard: project overview, methodology, headline
findings, and conclusion, with buttons linking OUT to each asset-class
dashboard, each of which is its own SEPARATE Streamlit app deployment
(own URL, own compute/memory allocation, own git folder) rather than a
page within this one.

Why separate deployments instead of one multi-page app: each Streamlit
Community Cloud app gets its own container/resource allocation, so
splitting Metals / Energy / Precious Metals / NGL into distinct
deployments means one heavy or crashing dashboard can't starve the
others, and two people can work on two different asset-class folders
independently without touching this hub or each other's deployment.

This file has no dependency on the root app.py or on any other
dashboard folder in this repo -- fully self-contained.
"""

import streamlit as st

from dashboard_shared import inject_css, metric_card, section_header

st.set_page_config(
    page_title="Metals Risk Premia - Hub",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)
inject_css()

# ═══════════════════════════════════════════════
# Deployment URLs for each asset-class dashboard.
# Fill in as each one is deployed as its own Streamlit Cloud app.
# ═══════════════════════════════════════════════
DASHBOARD_LINKS = {
    "Metals":         {"url": "https://metals-risk-premia-kj.streamlit.app/", "ready": True,
                        "desc": "LME Copper & Aluminium. Momentum, Carry, Value, Portfolio (10 tabs) — complete, Stage 1."},
    "Energy":         {"url": None, "ready": False,
                        "desc": "WTI, Brent, RBOB, Heating Oil, Nat Gas, Gasoil + 4 extended products. Stage 2, in progress."},
    "Precious Metals": {"url": None, "ready": False,
                        "desc": "Gold, Silver, Platinum, Palladium, Copper-CME. Stage 2, in progress."},
    "NGL / Refined":  {"url": None, "ready": False,
                        "desc": "Propane, Butane, Ethane, Isobutane, Ethylene, Propylene. Stage 2 — paused pending a ticker-mapping data check."},
}

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

section_header("CONCLUSION")
st.markdown(
    """
The diversification thesis holds on every metal tested so far: combining three economically distinct
premia consistently beats any single sleeve on a risk-adjusted basis, with materially lower drawdown.
Optimal parameters are asset-specific, not a one-size template — each new product gets its own
momentum speed, carry variant, and value anchor selected on its own data, same methodology throughout.
Stage 2 extends this test to a genuinely different asset class (energy, refined products) to see
whether the same three-premia structure holds outside metals.
"""
)

st.divider()
section_header("EXPLORE THE DASHBOARDS")
st.caption("Each asset class below is its own independent dashboard — click through for live signals, "
           "parameter controls, equity curves, rolling Sharpe, and performance metrics.")

nav_cols = st.columns(4)
for col, (name, info) in zip(nav_cols, DASHBOARD_LINKS.items()):
    with col:
        st.markdown(f"**{name}**")
        if info["ready"]:
            st.link_button(f"Open {name} →", info["url"], use_container_width=True)
        else:
            st.button("Coming Soon", disabled=True, use_container_width=True, key=f"soon_{name}")
        st.caption(info["desc"])

st.divider()
st.caption(
    "Data: LME futures curves (F1–F27), NYMEX/ICE/COMEX futures curves, LME Cash & 3M prices. "
    "Research prototype for academic purposes — in-sample backtests unless stated otherwise. Not "
    "investment advice."
)
