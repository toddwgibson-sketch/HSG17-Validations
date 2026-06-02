import streamlit as st
from utils.auth import require_login

require_login()

# Explicit navigation: only these pages appear in the sidebar.
# This hides the root "app" tab/page entirely (no more broken app tab or switch errors).
t0_tool = st.Page(
    "pages/01_HSG17_T1_to_T0_Tool.py",
    title="HSG17 T0-to-Host",
    icon="🖥️",
    default=True,
)
dashboard = st.Page(
    "pages/10_HSG17_Dashboard.py",
    title="HSG17 Dashboard",
    icon="📊",
)

pg = st.navigation([t0_tool, dashboard])
pg.run()
