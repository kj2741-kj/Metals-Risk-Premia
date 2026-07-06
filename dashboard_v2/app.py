"""
Metals Risk Premia - Multi-Page Dashboard (v2) - Entry Point
==============================================================
This is a SEPARATE, standalone dashboard from the root-level app.py --
it does not modify, import, or depend on the original single-file
dashboard in any way. The original app.py is untouched and keeps
working exactly as it did before.

Thin router: sets page config + global theme once, then hands off to
st.navigation(). Each asset-class dashboard lives in its own file and
is only executed by Streamlit when that page is actually selected --
unlike a single-file multi-tab app, unselected pages do zero
computation, which is the fix for "adding more asset classes will
make it slower / crash."

Run locally with:  streamlit run dashboard_v2/app.py
To deploy separately on Streamlit Cloud: point a NEW app at this file
path (dashboard_v2/app.py) in the same repo -- it gets its own URL
and does not affect the existing deployment pointed at the root app.py.

Page files (each independently executable, no shared top-level state):
  dashboard_home.py      - project hub: overview, methodology, links
  dashboard_metals.py    - LME Copper/Aluminium (copy of the original
                            app.py content, not a move -- the root
                            app.py is unaffected)
  dashboard_energy.py    - Oil & energy products               [Stage 2]
  dashboard_precious.py  - Precious metals                      [Stage 2]
  dashboard_ngl.py       - Refined NGL products                 [Stage 2, blocked]

Shared, side-effect-free constants/helpers (chart theme, CSS, generic
Sharpe/metrics engine) live in dashboard_shared.py so every page can
import them without re-executing another page as a side effect.
"""

import streamlit as st

from dashboard_shared import inject_css

st.set_page_config(
    page_title="Metals Risk Premia Dashboard",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_css()

pg = st.navigation(
    {
        "Overview": [
            st.Page("dashboard_home.py", title="Home", icon="🏠", default=True),
        ],
        "Asset Classes": [
            st.Page("dashboard_metals.py", title="Metals", icon="⚙️"),
        ],
    }
)
pg.run()
