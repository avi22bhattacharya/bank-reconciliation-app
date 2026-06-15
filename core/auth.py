"""Simple login wall for every Streamlit page.

Credentials live in st.secrets["users"][username]["password_hash"] (bcrypt).
Call require_login() right after st.set_page_config(); it stops the script
until the user authenticates.

To generate a password hash locally:
    python -c "import bcrypt; print(bcrypt.hashpw(b'yourpassword', bcrypt.gensalt()).decode())"
"""

from __future__ import annotations

import streamlit as st


def require_login() -> None:
    if st.session_state.get("authenticated"):
        return

    st.title("Bank Reconciliation — Login")

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")

    if submitted:
        import bcrypt
        users = st.secrets.get("users", {})
        user = users.get(username)
        if user and bcrypt.checkpw(
            password.encode(), user["password_hash"].encode()
        ):
            st.session_state["authenticated"] = True
            st.session_state["username"] = username
            st.rerun()
        else:
            st.error("Invalid username or password.")

    st.stop()
