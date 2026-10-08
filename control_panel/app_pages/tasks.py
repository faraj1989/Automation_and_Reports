import streamlit as st

import cached
import monitor

tasks = cached.tasks()

st.dataframe(
    [{"Task": t["name"].removeprefix(monitor.TASK_PREFIX).strip(), "State": t["state"],
      "Last run": t["last_run"], "Result": t["result_text"], "Next run": t["next_run"],
      "Trigger": t["triggers"], "Runs": t["action"]} for t in tasks],
    hide_index=True,
    column_config={
        "Last run": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm"),
        "Next run": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm"),
        "Runs": st.column_config.TextColumn(width="small"),
    },
)

with st.container(border=True):
    st.markdown("**Control a task**")
    name = st.selectbox("Task", [t["name"] for t in tasks], key="task_pick")
    task = next(t for t in tasks if t["name"] == name)
    with st.container(horizontal=True):
        if st.button("Run now", icon=":material/play_arrow:", type="primary", disabled=task["state"] == "Running"):
            cached.act(f"Start {name}", monitor.task_action, name, "start")
            st.rerun()
        if st.button("Stop", icon=":material/stop:", disabled=task["state"] != "Running"):
            cached.act(f"Stop {name}", monitor.task_action, name, "stop")
            st.rerun()
        if task["state"] == "Disabled":
            if st.button("Enable", icon=":material/toggle_on:"):
                cached.act(f"Enable {name}", monitor.task_action, name, "enable")
                st.rerun()
        elif st.button("Disable", icon=":material/toggle_off:"):
            cached.act(f"Disable {name}", monitor.task_action, name, "disable")
            st.rerun()
    if name == monitor.DASHBOARD_TASK and task["state"] == "Running":
        st.caption("Stopping the dashboard task ends the shared :8501 dashboard for everyone on the LAN.")
    if task["logon"] == "Password":
        st.caption("This task runs whether or not you are logged on, under your stored Windows password.")

st.subheader("Weekly/monthly jobs")
st.caption("Success stamps from periodic_jobs.py. The scraper watchdog re-runs any job still without a success "
           "45 min after its due time, retrying every 2 h - but only while the watchdog is running.")
periodic = cached.periodic()
st.dataframe(
    [{"Job": p["job"], "Schedule": p["schedule"], "Last due": p["last_due"], "Last success": p["last_success"],
      "State": ("Running (pid %s)" % p["running_pid"]) if p["running_pid"] else ("Overdue" if p["overdue"] else "OK")}
     for p in periodic],
    hide_index=True,
    column_config={"Last due": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm"),
                   "Last success": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm")},
)

with st.container(border=True):
    st.markdown("**Run a scheduler.py job now**")
    st.caption("Starts it in the background, like the watchdog backstop does. Output: scheduler.log.")
    jobs = {"Daily pipeline (yesterday)": "", "Hourly cells update": "--hourly-cells"}
    jobs.update({p["job"]: p["flag"] for p in periodic})
    pick = st.selectbox("Job", list(jobs), key="job_pick")
    busy = any(r["kind"] == "job" and r["label"].startswith("scheduler.py") for r in cached.processes())
    if busy:
        st.warning("A scheduler.py run is already in progress - starting another one in parallel can collide "
                   "on the same CSV history files.", icon=":material/warning:")
    if st.button("Start job", icon=":material/rocket_launch:", type="primary", disabled=busy):
        cached.act(f"Start {pick}", monitor.run_scheduler_job, jobs[pick])
        st.rerun()
