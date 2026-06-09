import sys
from pathlib import Path

# Ensure project root is on sys.path so "from utils.xxx" works for pages
# executed via st.navigation (critical for Streamlit Cloud /mount/src/... layout).
_root = Path(__file__).resolve().parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import streamlit as st
from utils.auth import require_login

require_login()

# Explicit navigation: only these pages appear in the sidebar.
# This hides the root "app" tab/page entirely (no more broken app tab or switch errors).
t0_tool = st.Page(
    "pages/01_HSG17_T1_to_T0_Tool.py",
    title="HSG17 T0-to-T1",
    icon="🖥️",
    default=True,
)
slack_tool = st.Page(
    "pages/02_HSG17_T1_to_T0_Slack.py",
    title="HSG17 Slack Tool",
    icon="🧰",
)
dashboard = st.Page(
    "pages/10_HSG17_Dashboard.py",
    title="HSG17 Dashboard",
    icon="📊",
)

pg = st.navigation([t0_tool, slack_tool, dashboard])
pg.run()
