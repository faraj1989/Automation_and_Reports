"""Weekly/monthly scheduler.py jobs - when each one is due, when it last
succeeded, and a per-job run lock. Shared by scheduler.py (takes the lock,
writes the success stamp) and scrapers/scraper_watchdog.py (re-runs any job
whose latest due time passed without a success - a backstop for Task
Scheduler, which skips a weekly run outright if the PC was off, asleep or on
battery at trigger time)."""

import json
import os
from contextlib import contextmanager
from datetime import datetime, timedelta

import psutil

from project_config import PROJECT_ROOT

STATE_DIR = PROJECT_ROOT / "logs" / "job_state"

# Mirrors the "Libyana NPM Weekly Interference" / "Weekly PS Traffic" /
# "Monthly Cell Info" Task Scheduler triggers - keep the two in sync.
# weekday: Monday=0 ... Sunday=6.
PERIODIC_JOBS = {
    "interference-weekly": {"flag": "--interference-weekly", "weekday": 6, "hour": 7, "minute": 0},
    "ps-traffic-weekly": {"flag": "--ps-traffic-weekly", "weekday": 6, "hour": 7, "minute": 30},
    "cell-info-monthly": {"flag": "--cell-info-monthly", "monthday": 1, "hour": 6, "minute": 0},
    # Previous month's CEM workbook for Tripoli HQ; reports failure while days
    # are missing, so this backstop also re-runs it until the month is complete.
    "cem-monthly": {"flag": "--cem-monthly", "monthday": 2, "hour": 6, "minute": 0},
}


def last_due(job, now=None):
    """Most recent scheduled trigger time at or before `now`."""
    spec = PERIODIC_JOBS[job]
    now = now or datetime.now()
    at = now.replace(hour=spec["hour"], minute=spec["minute"], second=0, microsecond=0)
    if "weekday" in spec:
        at -= timedelta(days=(at.weekday() - spec["weekday"]) % 7)
        if at > now:
            at -= timedelta(days=7)
    else:
        at = at.replace(day=spec["monthday"])
        if at > now:
            prev_month_end = at.replace(day=1) - timedelta(days=1)
            at = at.replace(year=prev_month_end.year, month=prev_month_end.month)
    return at


def _state_file(job):
    return STATE_DIR / f"{job}.json"


def last_success(job):
    try:
        with open(_state_file(job), encoding="utf-8") as f:
            return datetime.fromisoformat(json.load(f)["last_success"])
    except (OSError, ValueError, KeyError):
        return None


def mark_success(job, when=None):
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    when = when or datetime.now()
    tmp = _state_file(job).with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"last_success": when.isoformat(timespec="seconds")}, f)
    os.replace(tmp, _state_file(job))


def is_overdue(job, now=None):
    now = now or datetime.now()
    done = last_success(job)
    return done is None or done < last_due(job, now)


def _lock_file(job):
    return STATE_DIR / f"{job}.lock"


def lock_holder(job):
    """PID of a live scheduler.py process holding `job`'s lock, else None
    (a lock left behind by a crashed/killed run counts as free)."""
    try:
        pid = int(_lock_file(job).read_text(encoding="utf-8").strip())
        cmdline = " ".join(psutil.Process(pid).cmdline())
    except (OSError, ValueError, psutil.Error):
        return None
    return pid if "scheduler.py" in cmdline else None


@contextmanager
def job_lock(job):
    """Yields True if this process got the lock, False if another live run of
    the same job already holds it - so Task Scheduler's catch-up run and the
    watchdog's backstop run can never overlap."""
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    path = _lock_file(job)
    if lock_holder(job) not in (None, os.getpid()):
        yield False
        return
    path.write_text(str(os.getpid()), encoding="utf-8")
    try:
        yield True
    finally:
        try:
            if path.read_text(encoding="utf-8").strip() == str(os.getpid()):
                path.unlink()
        except OSError:
            pass
