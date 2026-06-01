import streamlit as st
from utils.auth import require_login

require_login()

st.set_page_config(page_title="HSG17 Validations", layout="wide")
st.title("HSG17 Validation Tools")
st.caption("Clean build for HSG17 site")

st.markdown("""
Welcome to the HSG17 validation toolkit.

**Available pages:**
- **01_HSG17_T0_to_Host** — 
- **10_HSG17_Dashboard** — 

Use the sidebar to navigate.
""")
