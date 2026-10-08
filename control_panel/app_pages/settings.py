import importlib.util

import pandas as pd
import streamlit as st

import auth
import cached
import monitor

# Load backend/config_manager.py on its own - importing the `backend` package
# would pull in every processor (pandas-heavy) just to edit ftp_config.json.
_spec = importlib.util.spec_from_file_location("config_manager", monitor.PROJECT_ROOT / "backend" / "config_manager.py")
config_manager = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(config_manager)

# Which running scraper reads which .env prefix.
PREFIX_TO_SCRAPER = {"MAE_": "MAE Combined Scraper", "NETECO_": "NetEco Combined Scraper",
                     "NCE_": "NCE Combined Scraper"}

st.caption(f"Every save first copies the old file to control_panel/backups/. "
           f"Running programs read their settings at start - restart them after a change.")

env_tab, ftp_tab, files_tab, pw_tab = st.tabs(["Credentials & paths (.env)", "SFTP (daily KPI download)",
                                               "KPI & rule files", "Panel password"])

with env_tab:
    sections = monitor.read_env_sections()
    for title, items in sections:
        with st.expander(f"{title}  ({len(items)})"):
            with st.form(f"env_{title}"):
                new = {}
                for key, value in items:
                    new[key] = st.text_input(key, value, type="password" if monitor.is_secret(key) else "default",
                                             key=f"env_{key}")
                if st.form_submit_button("Save section", type="primary"):
                    changes = {k: v for k, v in new.items() if v != dict(items)[k]}
                    if not changes:
                        st.info("Nothing changed.")
                    else:
                        monitor.write_env_values(changes)
                        affected = sorted({s for p, s in PREFIX_TO_SCRAPER.items() for k in changes if k.startswith(p)})
                        st.session_state.env_saved = (sorted(changes), affected)
                        st.rerun()

    if saved := st.session_state.get("env_saved"):
        keys, affected = saved
        st.success("Saved: " + ", ".join(keys), icon=":material/check_circle:")
        if affected:
            st.write("These scrapers use what you changed - restart them to pick it up:")
            if st.button("Restart " + ", ".join(affected), icon=":material/restart_alt:", type="primary"):
                for name in affected:
                    monitor.clear_login_failure(name)
                    cached.act(f"Restart {name}", lambda n=name: (monitor.stop_scraper(n, pause=False),
                                                                    monitor.start_scraper(n)))
                st.session_state.pop("env_saved")
                st.rerun()
        else:
            st.caption("Scheduled jobs pick up the new values on their next run.")

    with st.expander("Add a new setting"):
        with st.form("env_add", clear_on_submit=True):
            k = st.text_input("Name", placeholder="SOME_SETTING").strip().upper()
            v = st.text_input("Value")
            if st.form_submit_button("Add") and k:
                if any(k == key for _, items in sections for key, _ in items):
                    st.error(f"{k} already exists - edit it above.")
                else:
                    monitor.write_env_values({k: v})
                    st.rerun()

with ftp_tab:
    cfg = config_manager.ConfigManager(str(monitor.FTP_CONFIG))
    st.caption("ftp_config.json - SFTP on port 22, used by the daily pipeline, hourly cells and weekly "
               "interference downloads. (The FTPS-on-21 account for Cell Info lives in .env as FTP_*.)")
    with st.form("ftp"):
        host = st.text_input("Host", cfg.get("host", ""))
        port = st.text_input("Port", str(cfg.get("port", "22")))
        user = st.text_input("Username", cfg.get("username", ""))
        pw = st.text_input("Password", cfg.get("password", ""), type="password")
        remote = cfg.get("remote_path", "")
        remote_new = st.text_input("Remote path", remote.replace("\xa0", " "),
                                   help="The server folder name uses non-breaking spaces; they're kept unless you "
                                        "change this field.")
        local_root = st.text_input("Local download folder", cfg.get("local_root", ""))
        if st.form_submit_button("Save SFTP settings", type="primary"):
            monitor.backup_file(monitor.FTP_CONFIG)
            for key, val in (("host", host), ("port", port), ("username", user), ("password", pw),
                             ("local_root", local_root)):
                cfg.set(key, val)
            if remote_new != remote.replace("\xa0", " "):
                cfg.set("remote_path", remote_new)
            cfg.save()
            st.success("Saved - used from the next scheduled run.", icon=":material/check_circle:")

with files_tab:
    csvs = sorted(monitor.CONFIG_DIR.glob("*.csv")) + sorted(monitor.CONFIG_DIR.glob("*.json"))
    pick = st.selectbox("File", [p.name for p in csvs], key="cfg_file")
    path = monitor.CONFIG_DIR / pick
    if path.suffix == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        st.caption(f"{len(df)} rows - add or delete rows at the bottom of the table.")
        edited = st.data_editor(df, num_rows="dynamic", key=f"edit_{pick}", height=500)
        if st.button("Save file", icon=":material/save:", type="primary", key="save_csv"):
            monitor.backup_file(path)
            edited.to_csv(path, index=False)
            st.toast(f"Saved {pick}", icon=":material/check_circle:")
    else:
        text = st.text_area("Content", path.read_text(encoding="utf-8"), height=500, key=f"edit_{pick}")
        if st.button("Save file", icon=":material/save:", type="primary", key="save_json"):
            import json
            try:
                json.loads(text)
            except ValueError as e:
                st.error(f"Not valid JSON: {e}")
            else:
                monitor.backup_file(path)
                path.write_text(text, encoding="utf-8")
                st.toast(f"Saved {pick}", icon=":material/check_circle:")
    st.caption("The dashboard caches these - use its refresh/clear-cache option to see changes there right away.")

with pw_tab:
    with st.form("pw", clear_on_submit=True):
        old = st.text_input("Current password", type="password")
        p1 = st.text_input("New password", type="password")
        p2 = st.text_input("Repeat new password", type="password")
        if st.form_submit_button("Change password", type="primary"):
            if not auth.check_password(old):
                st.error("Current password is wrong.")
            elif len(p1) < 8 or p1 != p2:
                st.error("New passwords must match and be at least 8 characters.")
            else:
                auth.set_password(p1)
                st.success("Password changed.", icon=":material/check_circle:")
