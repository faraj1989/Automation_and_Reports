from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

import cached
import monitor

runs = pd.DataFrame(cached.runs())
smart = pd.DataFrame(monitor.smartcare_runs())
if not smart.empty:
    smart["duration_min"] = ((smart["ended"] - smart["started"]).dt.total_seconds() / 60).round(1)
    smart["target"], smart["warnings"], smart["first_error"] = "", 0, smart.pop("log")
allruns = pd.concat([runs, smart], ignore_index=True).sort_values("started", ascending=False)

with st.container(horizontal=True, vertical_alignment="bottom"):
    days = st.segmented_control("Period", [1, 7, 30, 90], default=7, format_func=lambda d: f"{d} d",
                                key="hist_days")
    jobs = st.multiselect("Jobs", sorted(allruns["job"].unique()), placeholder="All jobs", key="hist_jobs")
    only_bad = st.toggle("Problems only", key="hist_bad")

view = allruns[allruns["started"] >= datetime.now() - timedelta(days=days or 7)]
if jobs:
    view = view[view["job"].isin(jobs)]

ok = view["status"].eq("Success").sum()
with st.container(horizontal=True):
    st.metric("Runs", len(view), border=True)
    st.metric("Clean successes", ok, border=True)
    st.metric("With errors", view["status"].eq("Success with errors").sum(), border=True)
    st.metric("Failed / cut off", (~view["status"].str.startswith("Success")).sum(), border=True)
    st.metric("Success rate", f"{(view['status'].str.startswith('Success').mean() * 100 if len(view) else 0):.0f}%",
              border=True)

if only_bad:
    view = view[view["status"] != "Success"]

st.dataframe(
    view[["started", "job", "target", "status", "duration_min", "errors", "warnings", "first_error"]],
    hide_index=True,
    column_config={
        "started": st.column_config.DatetimeColumn("Started", format="YYYY-MM-DD HH:mm"),
        "job": "Job", "target": "For date", "status": "Status",
        "duration_min": st.column_config.NumberColumn("Minutes", format="%.1f"),
        "errors": "Errors", "warnings": "Warnings", "first_error": st.column_config.TextColumn("First error / log"),
    },
)
st.caption("From scheduler.log (every scheduler.py run, incl. ones started by Task Scheduler, the watchdog "
           "backstop or this panel) and logs/smartcare_reports_*.log.")

st.subheader("Watchdog activity")
events = monitor.watchdog_events()
restarts = [e for e in events if e["event"].startswith("Started ") or e["event"].startswith("Periodic job")]
st.caption(f"{len(restarts)} launches in the last {len(events)} watchdog log lines.")
st.dataframe(pd.DataFrame(events[:200]), hide_index=True,
             column_config={"time": st.column_config.DatetimeColumn("Time", format="YYYY-MM-DD HH:mm:ss"),
                            "event": "Event"})
