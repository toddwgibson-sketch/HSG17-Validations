import streamlit as st
from utils.auth import require_login

require_login()

# Hide the root app.py page entirely — immediately switch to the main tool page
# after login. The sidebar will still list the available pages.
st.switch_page("pages/01_HSG17_T0_to_Host.py")
