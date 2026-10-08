import streamlit as st

import cached
import monitor

st.caption("The scraper watchdog relaunches any scraper that isn't running within 60 s. "
           "**Pause** keeps one stopped; **Restart** kills it and lets the watchdog (or this panel) start it fresh.")

watchdog_up = any(r["kind"] == "watchdog" for r in cached.processes())
if not watchdog_up:
    st.warning("The scraper watchdog is not running, so stopped scrapers will stay stopped.",
               icon=":material/warning:")

for s in cached.scrapers():
    name, proc = s["name"], s["process"]
    with st.container(border=True):
        if proc:
            status = (f":green[:material/check_circle: running] · pid {proc['pid']} · "
                      f"{monitor.human_bytes(proc['memory'])} in {proc['tree_size']} processes "
                      f"({proc['chrome_procs']} Chrome) · up since {proc['started']:%d %b %H:%M}")
        elif s["paused"]:
            status = ":orange[:material/pause_circle: paused]"
        else:
            status = ":red[:material/error: not running]"
        st.markdown(f"**{name}** — {status}")
        if s["login_blocked"]:
            st.error("Login failed today with the current password - the watchdog won't retry until the "
                     "password in Settings changes (or tomorrow).", icon=":material/key_off:")
        with st.container(horizontal=True):
            if proc:
                if st.button("Restart", key=f"rs_{name}", icon=":material/restart_alt:"):
                    cached.act(f"Restart {name}", lambda n=name: (monitor.stop_scraper(n, pause=False),
                                                                    monitor.start_scraper(n)))
                    st.rerun()
                if st.button("Pause", key=f"ps_{name}", icon=":material/pause:"):
                    cached.act(f"Pause {name}", monitor.stop_scraper, name, True)
                    st.rerun()
            elif st.button("Start", key=f"st_{name}", icon=":material/play_arrow:", type="primary"):
                cached.act(f"Start {name}", monitor.start_scraper, name)
                st.rerun()
            if s["login_blocked"] and st.button("Clear login block", key=f"lb_{name}", icon=":material/lock_open:"):
                cached.act(f"Clear login block for {name}", monitor.clear_login_failure, name)
                st.rerun()

st.subheader("All project processes")
procs = cached.processes()
st.dataframe(
    [{"PID": r["pid"], "What": r["label"], "Type": r["kind"], "Started": r["started"],
      "Memory": monitor.human_bytes(r["memory"]), "Processes": r["tree_size"], "Command": r["cmd"]} for r in procs],
    hide_index=True,
    column_config={"Started": st.column_config.DatetimeColumn(format="YYYY-MM-DD HH:mm")},
)

others = [r for r in procs if r["kind"] not in ("scraper", "panel")]
if others:
    with st.expander("Stop a process"):
        st.caption("Ends that one process and its children by exact PID - nothing else is touched.")
        labels = {f"{r['label']} (pid {r['pid']})": r["pid"] for r in others}
        pick = st.selectbox("Process", list(labels))
        confirm = st.checkbox("Yes, stop it")
        if st.button("Stop process", icon=":material/stop_circle:", disabled=not confirm):
            cached.act(f"Stop {pick}", monitor.kill_tree, labels[pick])
            st.rerun()
