import streamlit as st
from utils.auth import require_login

require_login()

st.set_page_config(page_title="HSG17 Validations", layout="wide")
st.title("HSG17 Validation Tools")
st.caption("Clean build for HSG17 site")

st.markdown("""
Welcome to the HSG17 validation toolkit.

**Available pages:**
- **01_HSG17_T0_to_Host** — Clean processor with mismatch clustering + PP enrichment
- **10_HSG17_Dashboard** — Executive current-state view by DH Block (deltas, cards, progress to zero)

Use the sidebar to navigate.
""")