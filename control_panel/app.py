"""Libyana NPM private control panel - run with run_control_panel.bat.

Separate from the shared network dashboard (streamlit_dashboard.py, :8501,
open to the LAN): this app listens on 127.0.0.1:8502 only, refuses any
non-local client, and is password protected."""

import time

import streamlit as st

import auth

st.set_page_config(page_title="NPM control panel", page_icon=":material/admin_panel_settings:", layout="wide")

# Defense in depth on top of --server.address 127.0.0.1.
if st.context.ip_address not in (None, "127.0.0.1", "::1", "localhost"):
    st.error("The control panel is only available on this PC.")
    st.stop()

SESSION_MINUTES = 240

if not auth.is_configured():
    st.title("Set up the control panel")
    st.caption("First run - choose the password that will protect this panel.")
    with st.form("setup"):
        pw1 = st.text_input("New password", type="password")
        pw2 = st.text_input("Repeat password", type="password")
        if st.form_submit_button("Save password", type="primary"):
            if len(pw1) < 8:
                st.error("Use at least 8 characters.")
            elif pw1 != pw2:
                st.error("The passwords don't match.")
            else:
                auth.set_password(pw1)
                st.session_state.authed_at = time.time()
                st.rerun()
    st.stop()

authed_at = st.session_state.get("authed_at")
if not authed_at or time.time() - authed_at > SESSION_MINUTES * 60:
    st.session_state.pop("authed_at", None)
    st.title("NPM control panel")
    with st.form("login"):
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Unlock", type="primary"):
            if auth.check_password(pw):
                st.session_state.authed_at = time.time()
                st.rerun()
            else:
                time.sleep(1.5)
                st.error("Wrong password.")
    st.stop()

page = st.navigation(
    {
        "Monitor": [
            st.Page("app_pages/overview.py", title="Health overview", icon=":material/monitor_heart:", default=True),
            st.Page("app_pages/history.py", title="Task history", icon=":material/history:"),
            st.Page("app_pages/data.py", title="Data freshness", icon=":material/database:"),
            st.Page("app_pages/logs.py", title="Logs", icon=":material/article:"),
        ],
        "Control": [
            st.Page("app_pages/tasks.py", title="Scheduled tasks", icon=":material/schedule:"),
            st.Page("app_pages/processes.py", title="Scrapers & processes", icon=":material/memory:"),
            st.Page("app_pages/settings.py", title="Settings", icon=":material/settings:"),
        ],
    }
)

with st.sidebar:
    if st.button("Refresh data", icon=":material/refresh:", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    if st.button("Lock", icon=":material/lock:", width="stretch"):
        st.session_state.pop("authed_at", None)
        st.rerun()
    st.caption("Local only · 127.0.0.1:8502")

st.title(page.title)
page.run()
