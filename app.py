import streamlit as st
from utils.auth import require_login

require_login()
st.switch_page("pages/01_HSG17_T0_to_Host.py")
