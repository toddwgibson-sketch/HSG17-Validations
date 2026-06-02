import streamlit as st

def require_login():
    """
    Simple username/password gate.
    Currently hardcoded to admin / admin (case-insensitive).
    """
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.title("Login Required")
        st.markdown("HSG17 Validation Tools. U:Admin P:Admin")

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login")

            if submitted:
                if username.lower() == "admin" and password.lower() == "admin":
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Invalid username or password")

        st.stop()

    # Logged in - show logout in sidebar
    with st.sidebar:
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.rerun()
