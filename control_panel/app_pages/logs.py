from datetime import datetime

import streamlit as st

import monitor

files = monitor.list_log_files()
if not files:
    st.info("No log files found.")
    st.stop()

labels = {}
for p in files:
    try:
        rel = p.relative_to(monitor.PROJECT_ROOT)
    except ValueError:
        rel = p
    labels[f"{rel}  ·  {datetime.fromtimestamp(p.stat().st_mtime):%Y-%m-%d %H:%M}"] = p

with st.container(horizontal=True, vertical_alignment="bottom"):
    pick = st.selectbox("Log file", list(labels), key="log_pick")
    n = st.number_input("Last lines", 50, 5000, 300, step=50, key="log_lines")
    needle = st.text_input("Filter", placeholder="e.g. ERROR", key="log_filter")
    live = st.toggle("Follow", key="log_follow")


@st.fragment(run_every="10s" if live else None)
def show():
    text = monitor.tail_file(labels[pick], int(n) if not needle else int(n) * 20)
    if needle:
        text = "\n".join(l for l in text.splitlines() if needle.lower() in l.lower())
    st.code(text or "(nothing matches)", language=None, height=600)


show()
