"""Cached wrappers so page reruns don't re-query Task Scheduler / psutil.
Any page that changes state calls st.cache_data.clear() afterwards."""

from datetime import datetime

import streamlit as st

import monitor


@st.cache_data(ttl=20, show_spinner=False)
def tasks():
    return monitor.get_scheduled_tasks()


@st.cache_data(ttl=15, show_spinner=False)
def processes():
    return monitor.get_project_processes()


@st.cache_data(ttl=15, show_spinner=False)
def scrapers():
    return monitor.scraper_status()


@st.cache_data(ttl=30, show_spinner=False)
def periodic():
    return monitor.periodic_status()


@st.cache_data(ttl=60, show_spinner=False)
def freshness():
    return monitor.data_freshness()


@st.cache_data(ttl=30, show_spinner=False)
def system():
    return monitor.system_status()


@st.cache_data(ttl=15, show_spinner=False)
def dashboard_ok():
    return monitor.dashboard_health()[0]


@st.cache_data(ttl=60, show_spinner=False, max_entries=4)
def scheduler_runs(_mtime_key):
    return monitor.scheduler_runs()


def runs():
    log = monitor.SCHEDULER_LOG
    key = (log.stat().st_mtime, log.stat().st_size) if log.exists() else None
    return scheduler_runs(key)


def act(label, fn, *args):
    """Run a control action, report the outcome, drop cached status."""
    try:
        fn(*args)
    except Exception as e:  # noqa: BLE001
        st.toast(f"{label} failed: {e}", icon=":material/error:")
        return False
    st.cache_data.clear()
    st.toast(f"{label} - done ({datetime.now():%H:%M:%S})", icon=":material/check_circle:")
    return True
