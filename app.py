import streamlit as st
from utils.auth import require_login

require_login()

st.set_page_config(page_title="HSG17 Validations", layout="wide")
st.title("HSG17 Validation Tools")
st.caption("T1-to-T0 gold formatter (replacing old T0-to-Host) + retained Dashboard")

st.markdown("""
Welcome to the HSG17 validation toolkit.

**Available pages:**
- **01_HSG17_T0_to_Host** — T1-to-T0 gold formatter (replaces previous T0-to-Host impl; produces the 5 perfect tabs)
- **10_HSG17_Dashboard** — Executive view (current state + deltas; fully retained)

Use the sidebar to navigate.
""")
