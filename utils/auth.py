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

        # Use a plain button instead of st.form — form submit loads a separate
        # JS chunk (FormSubmitContent) that intermittently fails on Streamlit Cloud.
        username = st.text_input("Username", key="login_username")
        password = st.text_input("Password", type="password", key="login_password")

        if st.button("Login", key="login_button", type="primary"):
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
