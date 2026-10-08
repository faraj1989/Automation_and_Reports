"""Status collection and control actions for the private control panel.

Everything here is plain Python (no Streamlit) so it can be tested or reused
from a script. Sources of truth, all read live:
- Windows Task Scheduler ("Libyana NPM *" tasks) - via PowerShell.
- Running processes (psutil) - scrapers, watchdog, dashboard, scheduler runs.
- periodic_jobs.py success stamps / locks.
- scheduler.log, logs/scraper_watchdog.log, logs/smartcare_reports_*.log.
- Output files' modification times (data freshness)."""

import json
import os
import re
import shutil
import subprocess
import sys
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

import psutil

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import periodic_jobs  # noqa: E402
from project_config import (  # noqa: E402
    ENV_PATH, LOGIN_FAILURE_DIR, is_login_failed_today, is_paused, load_env_file,
    parse_env_file, set_paused,
)
from scrapers import scraper_watchdog  # noqa: E402

TASK_PREFIX = "Libyana NPM"
WATCHDOG_TASK = "Libyana NPM Scraper Watchdog"
DASHBOARD_TASK = "Libyana NPM Dashboard"
DASHBOARD_HEALTH_URL = "http://127.0.0.1:8501/_stcore/health"
SCHEDULER_LOG = PROJECT_ROOT / "scheduler.log"
WATCHDOG_LOG = PROJECT_ROOT / "logs" / "scraper_watchdog.log"
LOG_DIR = PROJECT_ROOT / "logs"
OUTPUT_CSV_DIR = PROJECT_ROOT / "output" / "csv"
CONFIG_DIR = PROJECT_ROOT / "config"
FTP_CONFIG = PROJECT_ROOT / "ftp_config.json"
BACKUP_DIR = Path(__file__).resolve().parent / "backups"

CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

# Task Scheduler LastTaskResult codes worth naming.
TASK_RESULTS = {
    0: "Success",
    1: "Failed (exit 1)",
    2: "Failed (exit 2)",
    267008: "Ready",
    267009: "Running now",
    267010: "Disabled",
    267011: "Never run",
    267014: "Stopped by user",
    2147750687: "Already running",
    3221225786: "Interrupted (Ctrl+C / window closed)",
}


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------

def run_powershell(script, timeout=60):
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout, creationflags=CREATE_NO_WINDOW,
    )
    return proc.returncode, proc.stdout.strip(), proc.stderr.strip()


def _ps_quote(value):
    return "'" + str(value).replace("'", "''") + "'"


def human_age(when, now=None):
    if when is None:
        return "never"
    secs = int(((now or datetime.now()) - when).total_seconds())
    if secs < 0:
        return "in the future"
    if secs < 90:
        return f"{secs}s ago"
    if secs < 5400:
        return f"{secs // 60}m ago"
    if secs < 172800:
        return f"{secs / 3600:.1f}h ago"
    return f"{secs / 86400:.1f}d ago"


def human_bytes(n):
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if abs(n) < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit in ("B", "KB") else f"{n:.1f} {unit}"
        n /= 1024


# --------------------------------------------------------------------------
# Scheduled tasks
# --------------------------------------------------------------------------

_TASKS_PS = r"""
$ErrorActionPreference = 'SilentlyContinue'
$out = foreach ($t in Get-ScheduledTask | Where-Object { $_.TaskName -like '%s*' }) {
    $i = $t | Get-ScheduledTaskInfo
    [pscustomobject]@{
        name     = $t.TaskName
        state    = [string]$t.State
        last_run = if ($i.LastRunTime -and $i.LastRunTime.Year -gt 2000) { $i.LastRunTime.ToString('s') } else { $null }
        next_run = if ($i.NextRunTime) { $i.NextRunTime.ToString('s') } else { $null }
        result   = [int64]$i.LastTaskResult
        missed   = [int]$i.NumberOfMissedRuns
        action   = ($t.Actions | ForEach-Object { ($_.Execute + ' ' + $_.Arguments).Trim() }) -join ' ; '
        triggers = ($t.Triggers | ForEach-Object {
            $kind = $_.CimClass.CimClassName -replace '^MSFT_Task','' -replace 'Trigger$',''
            $at = if ($_.StartBoundary) { ([datetime]$_.StartBoundary).ToString('ddd yyyy-MM-dd HH:mm') } else { '' }
            $rep = if ($_.Repetition.Interval) { ' every ' + $_.Repetition.Interval } else { '' }
            $days = if ($_.DaysOfWeek) { ' days=' + $_.DaysOfWeek } else { '' }
            "$kind $at$days$rep".Trim()
        }) -join ' | '
        logon    = [string]$t.Principal.LogonType
    }
}
@($out) | ConvertTo-Json -Depth 3 -Compress
""" % TASK_PREFIX


def get_scheduled_tasks():
    code, out, err = run_powershell(_TASKS_PS)
    if code != 0 or not out:
        raise RuntimeError(err or "Get-ScheduledTask returned nothing")
    data = json.loads(out)
    if isinstance(data, dict):
        data = [data]
    for t in data:
        t["result_text"] = TASK_RESULTS.get(t["result"], f"0x{t['result'] & 0xFFFFFFFF:08X}")
        t["last_run"] = datetime.fromisoformat(t["last_run"]) if t["last_run"] else None
        t["next_run"] = datetime.fromisoformat(t["next_run"]) if t["next_run"] else None
        t["ok"] = t["state"] != "Disabled" and t["result"] in (0, 267009, 267011)
    return sorted(data, key=lambda t: t["name"])


def task_action(name, action):
    """action: start | stop | enable | disable."""
    if not name.startswith(TASK_PREFIX):
        raise ValueError(f"Not a project task: {name}")
    cmdlet = {"start": "Start-ScheduledTask", "stop": "Stop-ScheduledTask",
              "enable": "Enable-ScheduledTask", "disable": "Disable-ScheduledTask"}[action]
    code, out, err = run_powershell(f"{cmdlet} -TaskName {_ps_quote(name)} -ErrorAction Stop | Out-Null")
    if code != 0:
        raise RuntimeError(err or f"{cmdlet} failed")


# --------------------------------------------------------------------------
# Processes
# --------------------------------------------------------------------------

def _classify(cmdline):
    c = cmdline.replace("\\", "/")
    for entry in scraper_watchdog.CONTINUOUS_SCRIPTS:
        if entry["file"] in c:
            return "scraper", entry["name"]
    if "scraper_watchdog.py" in c:
        return "watchdog", "Scraper watchdog"
    if "streamlit_dashboard.py" in c:
        return "dashboard", "Network dashboard (8501)"
    if "control_panel/app.py" in c:
        return "panel", "Control panel"
    if "scheduler.py" in c:
        flag = next((a for a in cmdline.split() if a.startswith("--")), "")
        return "job", f"scheduler.py {flag or '(daily pipeline)'}".strip()
    for script, label in (("smartcare_cem_scraper.py", "SmartCare CEM export"),
                          ("run_smartcare_analysis_task.py", "SmartCare CEM analysis"),
                          ("weekly_device_penetration_scraper.py", "Device penetration export"),
                          ("main_isp_state_scraper.py", "Main ISP state scraper"),
                          ("main_gui.py", "Desktop GUI (main_gui.py)")):
        if script in c:
            return "job", label
    if "/scrapers/" in c or "/reports/" in c:
        return "job", Path(c.split()[-1]).name
    return None, None


def get_project_processes():
    """One row per logical project process. The venv's python.exe is a
    launcher that re-spawns the real interpreter with the same command line,
    so each script shows up as a parent/child pair - keep the top one and
    fold the whole tree's memory (incl. chromedriver/Chrome) into it."""
    procs = {}
    root = str(PROJECT_ROOT).replace("\\", "/").lower()
    for p in psutil.process_iter(["pid", "ppid", "name", "cmdline", "create_time"]):
        try:
            cmd = " ".join(p.info["cmdline"] or [])
        except (psutil.Error, TypeError):
            continue
        if "python" not in (p.info["name"] or "").lower():
            continue
        kind, label = _classify(cmd)
        if not kind:
            continue
        cwd_ok = root in cmd.replace("\\", "/").lower()
        if not cwd_ok:
            try:
                cwd_ok = p.cwd().replace("\\", "/").lower() == root
            except psutil.Error:
                cwd_ok = False
        if cwd_ok:
            procs[p.info["pid"]] = dict(p.info, cmd=cmd, kind=kind, label=label)

    rows = []
    for pid, info in procs.items():
        parent = procs.get(info["ppid"])
        if parent and parent["cmd"] == info["cmd"]:
            continue  # launcher child - folded into its parent
        try:
            proc = psutil.Process(pid)
            tree = [proc] + proc.children(recursive=True)
        except psutil.Error:
            continue
        rss, chrome = 0, 0
        for t in tree:
            try:
                rss += t.memory_info().rss
                if "chrome" in t.name().lower():
                    chrome += 1
            except psutil.Error:
                pass
        rows.append({
            "pid": pid, "kind": info["kind"], "label": info["label"],
            "started": datetime.fromtimestamp(info["create_time"]),
            "memory": rss, "tree_size": len(tree), "chrome_procs": chrome, "cmd": info["cmd"],
        })
    return sorted(rows, key=lambda r: (r["kind"], r["label"]))


def kill_tree(pid):
    """Terminate one process and its children by exact PID (never by name -
    several live scrapers share python.exe/chrome.exe)."""
    proc = psutil.Process(pid)
    tree = proc.children(recursive=True) + [proc]
    for p in tree:
        try:
            p.terminate()
        except psutil.Error:
            pass
    _, alive = psutil.wait_procs(tree, timeout=8)
    for p in alive:
        try:
            p.kill()
        except psutil.Error:
            pass


def scraper_status():
    procs = {r["label"]: r for r in get_project_processes() if r["kind"] == "scraper"}
    env = parse_env_file()
    rows = []
    for entry in scraper_watchdog.CONTINUOUS_SCRIPTS:
        name = entry["name"]
        rows.append({
            "name": name, "file": entry["file"], "process": procs.get(name),
            "paused": is_paused(name),
            "login_blocked": is_login_failed_today(name, env.get(entry["password_env"])),
        })
    return rows


def start_scraper(name):
    entry = next(e for e in scraper_watchdog.CONTINUOUS_SCRIPTS if e["name"] == name)
    set_paused(name, False)
    if scraper_watchdog.is_process_running(entry["file"]):
        return
    load_env_file(override=True)  # hand the child today's .env, not this process's startup copy
    scraper_watchdog.start(entry)


def stop_scraper(name, pause=True):
    """pause=True keeps it stopped (watchdog skips it); False = restart via watchdog."""
    if pause:
        set_paused(name, True)
    for r in get_project_processes():
        if r["kind"] == "scraper" and r["label"] == name:
            kill_tree(r["pid"])


def clear_login_failure(name):
    (LOGIN_FAILURE_DIR / f"{name.replace(' ', '_')}.txt").unlink(missing_ok=True)


def dashboard_health():
    try:
        with urllib.request.urlopen(DASHBOARD_HEALTH_URL, timeout=4) as r:
            return r.status == 200, r.read(50).decode(errors="replace")
    except Exception as e:  # noqa: BLE001
        return False, str(e)


def run_scheduler_job(flag):
    """Start scheduler.py <flag> detached (same way the watchdog backstop does)."""
    allowed = {spec["flag"] for spec in periodic_jobs.PERIODIC_JOBS.values()} | {"--hourly-cells", ""}
    if flag not in allowed:
        raise ValueError(flag)
    env = os.environ.copy()
    env.update(PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    log = LOG_DIR / "job_state" / f"panel_run_{(flag.strip('-') or 'daily')}.log"
    log.parent.mkdir(parents=True, exist_ok=True)
    with open(log, "w", encoding="utf-8", errors="replace") as f:
        subprocess.Popen(
            [sys.executable, str(PROJECT_ROOT / "scheduler.py")] + ([flag] if flag else []),
            cwd=str(PROJECT_ROOT), stdout=f, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
            env=env, creationflags=CREATE_NO_WINDOW, start_new_session=True,
        )


# --------------------------------------------------------------------------
# Periodic jobs (weekly/monthly backstop state)
# --------------------------------------------------------------------------

def periodic_status(now=None):
    now = now or datetime.now()
    rows = []
    for job, spec in periodic_jobs.PERIODIC_JOBS.items():
        when = (f"Sun {spec['hour']:02d}:{spec['minute']:02d}" if spec.get("weekday") == 6 else
                f"weekday {spec['weekday']} {spec['hour']:02d}:{spec['minute']:02d}" if "weekday" in spec else
                f"day {spec['monthday']} {spec['hour']:02d}:{spec['minute']:02d}")
        rows.append({
            "job": job, "flag": spec["flag"], "schedule": when,
            "last_due": periodic_jobs.last_due(job, now),
            "last_success": periodic_jobs.last_success(job),
            "overdue": periodic_jobs.is_overdue(job, now),
            "running_pid": periodic_jobs.lock_holder(job),
        })
    return rows


# --------------------------------------------------------------------------
# Run history ("what got done") from the logs
# --------------------------------------------------------------------------

_LINE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}),\d+ - (\w+) - (.*)$")
_JOBS = [
    # job, start marker, success marker, failure marker
    ("Daily pipeline", "STARTING DAILY PIPELINE", "PIPELINE COMPLETED SUCCESSFULLY", "❌ Pipeline failed"),
    ("Hourly cells", "STARTING HOURLY CELLS UPDATE", "HOURLY CELLS UPDATE COMPLETED", "❌ Hourly cells update failed"),
    ("Weekly interference", "STARTING WEEKLY INTERFERENCE UPDATE", "WEEKLY INTERFERENCE UPDATE COMPLETED",
     "❌ Weekly interference update failed"),
    ("Weekly PS traffic", "STARTING WEEKLY PS TRAFFIC PER SITE UPDATE", "WEEKLY PS TRAFFIC UPDATE COMPLETED",
     "❌ Weekly PS traffic update failed"),
    ("Monthly cell info", "STARTING MONTHLY CELL INFO UPDATE", "MONTHLY CELL INFO UPDATE COMPLETED",
     "❌ Monthly cell info update failed"),
    ("Monthly CEM report", "BUILDING MONTHLY CEM COMPREHENSIVE ANALYSIS", "MONTHLY CEM REPORT COMPLETE",
     "❌ Monthly CEM report failed"),
]


def scheduler_runs(path=SCHEDULER_LOG):
    """Pair each job's start line with its completion/failure line. Error and
    warning lines in between are attributed to the most recently started open
    run (jobs can overlap - e.g. hourly cells during the daily pipeline)."""
    runs, open_runs = [], {}
    if not path.exists():
        return runs
    with open(path, encoding="utf-8", errors="replace") as f:
        for line in f:
            if " - INFO - " in line and "STARTING" not in line and "COMPLETE" not in line \
                    and "BUILDING MONTHLY CEM" not in line:
                continue
            m = _LINE.match(line.rstrip("\n"))
            if not m:
                continue
            ts, level, msg = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S"), m.group(2), m.group(3)
            for job, start, ok, fail in _JOBS:
                if start in msg:
                    if job in open_runs:
                        open_runs[job]["status"] = "No end logged (killed/crashed?)"
                        runs.append(open_runs.pop(job))
                    target = msg.split(" - ", 1)[1].strip() if " - " in msg else ""
                    open_runs[job] = {"job": job, "target": target, "started": ts, "ended": None,
                                      "status": "Running", "errors": 0, "warnings": 0, "first_error": ""}
                    break
                if ok in msg or fail in msg:
                    run = open_runs.pop(job, None)
                    if run:
                        run["ended"] = ts
                        run["status"] = "Success" if ok in msg else "Failed"
                        if fail in msg and not run["first_error"]:
                            run["first_error"] = msg
                        runs.append(run)
                    break
            else:
                if "Missing day(s)" in msg and "Monthly CEM report" in open_runs:
                    run = open_runs.pop("Monthly CEM report")
                    run.update(ended=ts, status="Incomplete (missing days)", first_error=msg.strip())
                    runs.append(run)
                elif level in ("ERROR", "WARNING") and open_runs:
                    run = max(open_runs.values(), key=lambda r: r["started"])
                    if level == "ERROR":
                        run["errors"] += 1
                        if not run["first_error"]:
                            run["first_error"] = msg.strip()
                    else:
                        run["warnings"] += 1
    stale = datetime.now() - timedelta(hours=3)
    for run in open_runs.values():
        if run["started"] < stale:
            run["status"] = "No end logged (killed/crashed?)"
        runs.append(run)
    for r in runs:
        r["duration_min"] = round((r["ended"] - r["started"]).total_seconds() / 60, 1) if r["ended"] else None
        if r["status"] == "Success" and r["errors"]:
            r["status"] = "Success with errors"
    return sorted(runs, key=lambda r: r["started"], reverse=True)


def smartcare_runs():
    rows = []
    for path in sorted(LOG_DIR.glob("smartcare_reports_*.log"), reverse=True):
        text = path.read_text(encoding="utf-8", errors="replace")
        m = re.search(r"smartcare_reports_(\d{8}_\d{6})(?:_(\w+))?", path.name)
        started = datetime.strptime(m.group(1), "%Y%m%d_%H%M%S") if m else None
        errors = len(re.findall(r"Traceback|\[ERROR\]| - ERROR - ", text))
        finished = "Finished" in text.splitlines()[-1] if text.strip() else False
        rows.append({
            "job": f"SmartCare ({(m.group(2) if m and m.group(2) else 'all')})",
            "started": started, "ended": datetime.fromtimestamp(path.stat().st_mtime),
            "status": ("Failed" if errors and not finished else
                       "Success with errors" if errors else "Success" if finished else "Running / cut off"),
            "errors": errors, "log": path.name,
        })
    return rows


def watchdog_events(path=WATCHDOG_LOG, limit=400):
    if not path.exists():
        return []
    rows = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines()[-limit:]:
        m = re.match(r"^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] (.*)$", line)
        if m:
            rows.append({"time": datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S"), "event": m.group(2)})
    return list(reversed(rows))


# --------------------------------------------------------------------------
# Data freshness
# --------------------------------------------------------------------------

def _newest_file(base, depth=2):
    """Newest file under base, looking only into the newest 3 subfolders at
    each level (export folders are <base>/<YYYY-MM-DD>/... and can hold
    thousands of files - a full walk would be far too slow)."""
    base = Path(base)
    if not base.exists():
        return None, None
    best = (None, None)
    try:
        entries = list(os.scandir(base))
    except OSError:
        return None, None
    for e in entries:
        try:
            if e.is_file():
                mt = e.stat().st_mtime
                if best[0] is None or mt > best[0]:
                    best = (mt, e.path)
        except OSError:
            pass
    if depth > 0:
        dirs = sorted((e for e in entries if e.is_dir()), key=lambda e: e.stat().st_mtime, reverse=True)[:3]
        for d in dirs:
            mt, p = _newest_file(d.path, depth - 1)
            if mt and (best[0] is None or mt > best[0]):
                best = (mt, p)
    return best


def data_freshness():
    env = parse_env_file()
    sources = []
    for key, label in (("MAE_EXPORT_BASE_DIR", "Current alarms (MAE/NetEco)"),
                       ("MAE_HISTORICAL_EXPORT_BASE_DIR", "Historical alarms (MAE/NetEco)"),
                       ("NCE_ACTIVE_EXPORT_BASE_DIR", "NCE current alarms"),
                       ("NCE_HISTORICAL_EXPORT_BASE_DIR", "NCE historical alarms")):
        if env.get(key):
            sources.append((label, env[key], timedelta(minutes=45)))
    for key, label, max_age in (("SMARTCARE_OUTPUT_DIR", "SmartCare CEM exports", timedelta(hours=30)),
                                ("WEEKLY_DEVICE_PENETRATION_OUTPUT_DIR", "Device penetration exports", timedelta(days=8)),
                                ("CELL_INFO_OUTPUT_DIR", "Cell info (monthly)", timedelta(days=32)),
                                ("PS_TRAFFIC_OUTPUT_DIR", "PS traffic per site (weekly)", timedelta(days=8))):
        if env.get(key):
            sources.append((label, env[key], max_age))
    try:
        local_root = json.loads(FTP_CONFIG.read_text(encoding="utf-8")).get("local_root")
        if local_root:
            sources.append(("SFTP daily downloads", local_root, timedelta(hours=30)))
    except (OSError, ValueError):
        pass

    now = datetime.now()
    rows = []
    for label, path, max_age in sources:
        mt, newest = _newest_file(path)
        when = datetime.fromtimestamp(mt) if mt else None
        rows.append({"source": label, "newest": when, "age": human_age(when, now),
                     "stale": when is None or now - when > max_age, "max_age_h": max_age.total_seconds() / 3600,
                     "file": newest or path})
    for csv in sorted(OUTPUT_CSV_DIR.glob("*.csv")):
        name = csv.stem
        max_age = (timedelta(days=8) if "Interference" in name else
                   timedelta(hours=7) if name.endswith("Cell_Hourly") else timedelta(hours=30))
        when = datetime.fromtimestamp(csv.stat().st_mtime)
        rows.append({"source": f"output/csv/{csv.name}", "newest": when, "age": human_age(when, now),
                     "stale": now - when > max_age, "max_age_h": max_age.total_seconds() / 3600,
                     "file": str(csv), "size": csv.stat().st_size})
    return rows


# --------------------------------------------------------------------------
# System
# --------------------------------------------------------------------------

def system_status():
    drives = {}
    env = parse_env_file()
    for p in (PROJECT_ROOT, env.get("DATA_ROOT")):
        if p and Path(p).exists():
            anchor = Path(p).anchor
            if anchor not in drives:
                du = shutil.disk_usage(anchor)
                drives[anchor] = {"total": du.total, "used": du.used, "free": du.free}
    vm = psutil.virtual_memory()
    chrome = [p for p in psutil.process_iter(["name", "memory_info"]) if "chrome" in (p.info["name"] or "").lower()]
    return {
        "boot": datetime.fromtimestamp(psutil.boot_time()),
        "cpu": psutil.cpu_percent(interval=0.3),
        "ram_used": vm.percent, "ram_total": vm.total,
        "drives": drives,
        "chrome_count": len(chrome),
        "chrome_mem": sum(p.info["memory_info"].rss for p in chrome if p.info["memory_info"]),
            }


# --------------------------------------------------------------------------
# Health roll-up (the "watchdog of everything")
# --------------------------------------------------------------------------

def health_checks(tasks, procs, scrapers, periodic, freshness, system, dash_ok):
    """List of (component, level, detail) - level: ok | warn | bad."""
    checks = []
    kinds = {r["kind"] for r in procs}

    checks.append(("Scraper watchdog", "ok" if "watchdog" in kinds else "bad",
                   "running" if "watchdog" in kinds else "NOT running - crashed scrapers will not be restarted"))
    checks.append(("Network dashboard", "ok" if dash_ok else "bad",
                   "responding on :8501" if dash_ok else "not responding on :8501"))
    for s in scrapers:
        if s["paused"]:
            checks.append((s["name"], "warn", "paused from the control panel"))
        elif s["login_blocked"]:
            checks.append((s["name"], "bad", "login failed today - fix the password in Settings"))
        elif s["process"]:
            checks.append((s["name"], "ok", f"running, {human_bytes(s['process']['memory'])}"))
        else:
            checks.append((s["name"], "bad", "not running"))
    for t in tasks:
        if t["state"] == "Disabled":
            checks.append((f"Task: {t['name']}", "warn", "disabled"))
        elif not t["ok"]:
            checks.append((f"Task: {t['name']}", "bad", f"last result: {t['result_text']}"))
    for p in periodic:
        if p["overdue"] and not p["running_pid"] and datetime.now() > p["last_due"] + timedelta(hours=3):
            checks.append((f"Job: {p['job']}", "bad", f"no success since due {p['last_due']:%Y-%m-%d %H:%M}"))
    stale = [f["source"] for f in freshness if f["stale"]]
    if stale:
        checks.append(("Data freshness", "warn", f"{len(stale)} stale: " + ", ".join(stale[:4])
                       + (" ..." if len(stale) > 4 else "")))
    else:
        checks.append(("Data freshness", "ok", "all sources fresh"))
    for anchor, du in system["drives"].items():
        free = du["free"] / du["total"]
        checks.append((f"Disk {anchor}", "bad" if free < 0.05 else "warn" if free < 0.10 else "ok",
                       f"{human_bytes(du['free'])} free of {human_bytes(du['total'])}"))
    if system["ram_used"] > 92:
        checks.append(("Memory", "warn", f"RAM {system['ram_used']:.0f}% used"))
    return checks


# --------------------------------------------------------------------------
# Settings files
# --------------------------------------------------------------------------

SECRET_HINTS = ("PASSWORD", "TOKEN", "SECRET", "KEY")


def is_secret(key):
    return any(h in key.upper() for h in SECRET_HINTS)


def backup_file(path):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    dest = BACKUP_DIR / f"{Path(path).name}.{datetime.now():%Y%m%d_%H%M%S}.bak"
    shutil.copy2(path, dest)
    return dest


def read_env_sections(path=ENV_PATH):
    """[(section_title, [(key, value), ...])] in file order; section titles
    come from the comment block above each group of keys."""
    sections, title, items, pending_comment = [], "General", [], None
    for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.strip()
        if line.startswith("#"):
            if pending_comment is None:
                pending_comment = line.lstrip("# ").strip()
            continue
        if not line:
            continue
        if "=" not in line:
            continue
        if pending_comment is not None:
            if items:
                sections.append((title, items))
            title, items, pending_comment = pending_comment or "General", [], None
        key = line.split("=", 1)[0].strip()
        items.append(key)
    if items:
        sections.append((title, items))
    values = parse_env_file(path)
    return [(t, [(k, values.get(k, "")) for k in keys]) for t, keys in sections]


def _quote_env(value):
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def write_env_values(changes, path=ENV_PATH):
    """Update existing keys in place (comments/order preserved), append new ones."""
    backup_file(path)
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    remaining = dict(changes)
    for i, raw in enumerate(lines):
        s = raw.strip()
        if not s or s.startswith("#") or "=" not in s:
            continue
        key = s.split("=", 1)[0].strip()
        if key in remaining:
            lines[i] = f"{key}={_quote_env(remaining.pop(key))}"
    for key, value in remaining.items():
        lines.append(f"{key}={_quote_env(value)}")
    tmp = path.with_suffix(".tmp")
    tmp.write_text("\n".join(lines) + "\n", encoding="utf-8")
    os.replace(tmp, path)


def tail_file(path, lines=300):
    path = Path(path)
    size = path.stat().st_size
    with open(path, "rb") as f:
        f.seek(max(0, size - lines * 400))
        data = f.read().decode("utf-8", errors="replace")
    return "\n".join(data.splitlines()[-lines:])


def list_log_files():
    files = [SCHEDULER_LOG] if SCHEDULER_LOG.exists() else []
    files += list(LOG_DIR.rglob("*.log"))
    extra = parse_env_file().get("PROJECT_LOG_DIR")
    if extra and Path(extra).exists():
        files += list(Path(extra).glob("*.log"))
    return sorted(set(files), key=lambda p: p.stat().st_mtime, reverse=True)
