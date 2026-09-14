"""Headless watchdog that keeps the 6 continuous NOC alarm scrapers
(MAE/NetEco/NCE, current + historical) running - ported from the sibling
NOC Automation Suite's service_watchdog.py/script_registry.py pattern.

Point Windows Task Scheduler at this once (run at log on) and it checks
every CHECK_INTERVAL_SECONDS whether each managed scraper is alive,
relaunching any that crashed, were never started, or got closed along with
a VS Code/terminal session - forever, until the machine reboots or this
watchdog process itself is stopped. It refuses to run a second copy of
itself (is_only_instance) so double-triggering the scheduled task can't
spawn two supervisors racing to relaunch the same scrapers.

A scraper whose login failed today is deliberately NOT relaunched - see
project_config.is_login_failed_today. That marker records a fingerprint of
the password that failed, so fixing MAE_PASSWORD/NETECO_PASSWORD/
NCE_PASSWORD in .env unblocks it on the very next check (no need to wait
for a new calendar day, and no risk of hammering a portal's login page
with a password already known to be wrong)."""


import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import psutil

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import PROJECT_ROOT, env_str, is_login_failed_today

PYTHON_EXE = sys.executable
LOG_DIR = PROJECT_ROOT / "logs" / "scrapers"
CHECK_INTERVAL_SECONDS = 60

# name: shown in logs and as the mark_login_failed() key each scraper uses.
# file: relative path the scraper script itself is launched with.
# password_env: which .env variable's fingerprint gates the login-failed
# skip - must match the PASSWORD each scraper reads (see each script's
# `PASSWORD = env_str("..._PASSWORD")` line).
CONTINUOUS_SCRIPTS = [
    # Combined 2026-09-14 (see chat that day): mae_scraper.py +
    # mae_historical_alarms_scraper.py merged into one Chrome process
    # (two tabs, one login session) to cut memory - validated live over 7+
    # clean cycles plus a forced Chrome-crash recovery test before being
    # adopted here. If NCE/NetEco get the same treatment later, replace
    # their two entries below the same way.
    {"name": "MAE Combined Scraper", "file": "scrapers/mae_combined_scraper.py", "password_env": "MAE_PASSWORD"},
    # Combined 2026-09-14: nce_active_alarms_scraper.py + nce_historical_
    # alarms_scraper.py merged the same way - validated live over 3+ clean
    # cycles (worked on the first attempt, no bugs to fix, since the two
    # fixes MAE's pilot needed were built in from the start).
    {"name": "NCE Combined Scraper", "file": "scrapers/nce_combined_scraper.py", "password_env": "NCE_PASSWORD"},
    # Combined 2026-09-14/15: "neteco_continuous all alrams.py" (Current) +
    # neteco_historical_alarms_scraper.py merged the same way - the most
    # structurally different of the three (ID-based export flow, not
    # text-based), validated live over 3+ clean cycles, worked on the
    # first attempt.
    {"name": "NetEco Combined Scraper", "file": "scrapers/neteco_combined_scraper.py", "password_env": "NETECO_PASSWORD"},
]


def log(message):
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}", flush=True)


def is_process_running(script_file):
    """Match by filename substring in cmdline - see script_registry.py's
    original docstring (sibling project) for why both sides are normalized
    to forward slashes before comparing."""
    target = script_file.replace("\\", "/")
    for proc in psutil.process_iter(["pid", "cmdline"]):
        try:
            cmdline = " ".join(proc.info["cmdline"] or [])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        normalized = cmdline.replace("\\", "/")
        if target in normalized and "python" in cmdline.lower():
            return proc.info["pid"]
    return None


def _self_process_tree_pids() -> set:
    this_pid = os.getpid()
    pids = {this_pid}
    try:
        me = psutil.Process(this_pid)
        for ancestor in me.parents():
            pids.add(ancestor.pid)
        for descendant in me.children(recursive=True):
            pids.add(descendant.pid)
    except psutil.NoSuchProcess:
        pass
    return pids


def is_only_instance():
    exclude_pids = _self_process_tree_pids()
    for proc in psutil.process_iter(["pid", "cmdline"]):
        if proc.info["pid"] in exclude_pids:
            continue
        try:
            cmdline = " ".join(proc.info["cmdline"] or [])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if "scraper_watchdog.py" in cmdline.replace("\\", "/") and "python" in cmdline.lower():
            return False
    return True


def start(entry):
    script_path = PROJECT_ROOT / entry["file"]
    if not script_path.exists():
        log(f"ERROR: script not found, skipping: {entry['name']} ({script_path})")
        return

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"{entry['name'].replace(' ', '_')}_{datetime.now():%Y%m%d_%H%M%S}.log"
    child_env = os.environ.copy()
    child_env["PYTHONUTF8"] = "1"
    child_env["PYTHONIOENCODING"] = "utf-8"
    child_env["PYTHONUNBUFFERED"] = "1"

    with open(log_file, "w", encoding="utf-8", errors="replace") as f:
        f.write(f"=== {entry['name']} started by watchdog at {datetime.now()} ===\n")
        f.write(f"Script: {script_path}\n{'=' * 60}\n\n")
        f.flush()
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        subprocess.Popen(
            [PYTHON_EXE, str(script_path)],
            cwd=str(PROJECT_ROOT),
            stdout=f,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            env=child_env,
            creationflags=creationflags,
            start_new_session=True,
        )
    log(f"Started {entry['name']} (log: {log_file.name})")


def main():
    if not is_only_instance():
        log("Another watchdog process is already running - exiting so there's only one.")
        return

    log("Watchdog starting. Managing: " + ", ".join(e["name"] for e in CONTINUOUS_SCRIPTS))
    log(f"Check interval: {CHECK_INTERVAL_SECONDS}s")
    import time
    while True:
        for entry in CONTINUOUS_SCRIPTS:
            if is_process_running(entry["file"]):
                continue
            current_password = env_str(entry["password_env"])
            if is_login_failed_today(entry["name"], current_password):
                log(f"{entry['name']} is not running - skipping, login failed today with the "
                    f"current {entry['password_env']} (will retry as soon as that's updated in "
                    f".env, or automatically tomorrow).")
                continue
            log(f"{entry['name']} is not running - starting it.")
            start(entry)
        time.sleep(CHECK_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
