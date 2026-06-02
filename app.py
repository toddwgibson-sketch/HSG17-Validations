import streamlit as st
from utils.auth import require_login

require_login()

st.set_page_config(page_title="HSG17 Validations", layout="wide")
st.title("HSG17 Validation Formatter")
st.caption("")

st.markdown("""


Use the sidebar to navigate.
""")
