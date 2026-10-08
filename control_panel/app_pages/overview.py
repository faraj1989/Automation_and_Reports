from datetime import datetime

import streamlit as st

import cached
import monitor

ICON = {"ok": ":material/check_circle:", "warn": ":material/warning:", "bad": ":material/error:"}
COLOR = {"ok": "green", "warn": "orange", "bad": "red"}


@st.fragment(run_every="30s")
def health():
    with st.spinner("Checking everything..."):
        tasks, procs, scrapers = cached.tasks(), cached.processes(), cached.scrapers()
        periodic, fresh, system = cached.periodic(), cached.freshness(), cached.system()
        dash_ok = cached.dashboard_ok()
    checks = monitor.health_checks(tasks, procs, scrapers, periodic, fresh, system, dash_ok)
    bad = [c for c in checks if c[1] == "bad"]
    warn = [c for c in checks if c[1] == "warn"]

    if bad:
        st.error(f"{len(bad)} problem(s) need attention", icon=":material/error:")
    elif warn:
        st.warning(f"All running - {len(warn)} warning(s)", icon=":material/warning:")
    else:
        st.success("Everything is healthy", icon=":material/check_circle:")

    with st.container(horizontal=True):
        st.metric("Problems", len(bad), border=True)
        st.metric("Warnings", len(warn), border=True)
        st.metric("Scrapers running", f"{sum(1 for s in scrapers if s['process'])}/{len(scrapers)}", border=True)
        st.metric("Chrome processes", system["chrome_count"], f"{monitor.human_bytes(system['chrome_mem'])}",
                  delta_color="off", border=True)
        st.metric("RAM used", f"{system['ram_used']:.0f}%", border=True)
        st.metric("PC up since", f"{system['boot']:%d %b %H:%M}", border=True)

    # Quick fixes for the two components nothing else restarts.
    kinds = {r["kind"] for r in procs}
    if "watchdog" not in kinds or not dash_ok:
        with st.container(horizontal=True):
            if "watchdog" not in kinds and st.button("Start scraper watchdog", icon=":material/play_arrow:",
                                                     type="primary"):
                cached.act("Start watchdog", monitor.task_action, monitor.WATCHDOG_TASK, "start")
                st.rerun()
            if not dash_ok and st.button("Start network dashboard", icon=":material/play_arrow:", type="primary"):
                cached.act("Start dashboard", monitor.task_action, monitor.DASHBOARD_TASK, "start")
                st.rerun()

    order = {"bad": 0, "warn": 1, "ok": 2}
    with st.container(border=True):
        for name, level, detail in sorted(checks, key=lambda c: order[c[1]]):
            st.markdown(f"{ICON[level]} **{name}** — :{COLOR[level]}[{detail}]")

    st.subheader("Latest runs")
    latest = {}
    for r in cached.runs():
        latest.setdefault(r["job"], r)
    st.dataframe(
        [{"Job": j, "Started": r["started"], "Status": r["status"], "Minutes": r["duration_min"],
          "Ago": monitor.human_age(r["started"]), "First error": r["first_error"]} for j, r in latest.items()],
        hide_index=True,
        column_config={"Started": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm")},
    )
    st.caption(f"Auto-refreshes every 30 s · checked {datetime.now():%H:%M:%S}")


health()
