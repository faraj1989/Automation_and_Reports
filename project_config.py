import hashlib
import json
import os
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
ENV_PATH = PROJECT_ROOT / ".env"
LOGIN_FAILURE_DIR = PROJECT_ROOT / "logs" / "login_failures"
PAUSED_DIR = PROJECT_ROOT / "logs" / "paused"


class LoginFailedError(RuntimeError):
    """Raised when a scraper detects bad/expired credentials rather than a
    transient failure. Distinguishing this lets a watchdog stop relaunching
    that one script for the rest of today instead of hammering the login
    page with a known-bad password until a human fixes it."""


def _strip_quotes(value):
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        quoted = value[1:-1]
        result = []
        escaped = False
        for char in quoted:
            if escaped:
                result.append({"n": "\n", "r": "\r", "t": "\t"}.get(char, char))
                escaped = False
            elif char == "\\":
                escaped = True
            else:
                result.append(char)
        if escaped:
            result.append("\\")
        return "".join(result)
    return value


def parse_env_file(path=ENV_PATH):
    values = {}
    if not path.exists():
        return values

    for raw_line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        if key:
            values[key] = _strip_quotes(value)
    return values


def load_env_file(path=ENV_PATH, override=False):
    """Load project .env values without overriding real environment variables."""
    for key, value in parse_env_file(path).items():
        if override or key not in os.environ:
            os.environ[key] = value


def env_str(name, default=""):
    load_env_file()
    value = os.getenv(name)
    if not value:
        return default
    return value


def env_int(name, default):
    value = env_str(name, str(default)).strip()
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def env_path(name, default):
    return Path(env_str(name, str(default)))


def env_path_str(name, default):
    return str(env_path(name, default))


def keep_only_latest_export(base_dir, pattern, keep_path):
    """Delete every file under base_dir (searched one dated-subfolder deep,
    matching the active-alarm scrapers' own <base_dir>/<YYYY-MM-DD>/<file>
    layout, plus a flat fallback) matching pattern except keep_path.

    Active/current-alarm scrapers (MAE, NetEco All Alarms, NCE Active) only
    need their single most recent export - each cycle's file supersedes the
    last one entirely, unlike the historical scrapers which intentionally
    accumulate every export for historical analysis."""
    base_dir = Path(base_dir)
    if not base_dir.exists():
        return
    keep_resolved = Path(keep_path).resolve()
    candidates = list(base_dir.glob(f"*/{pattern}")) + list(base_dir.glob(pattern))
    for old in candidates:
        if not old.is_file() or old.resolve() == keep_resolved:
            continue
        try:
            old.unlink()
        except OSError:
            pass


def _login_failure_marker_path(script_name):
    safe_name = script_name.replace(" ", "_")
    return LOGIN_FAILURE_DIR / f"{safe_name}.txt"


def _password_fingerprint(password):
    if not password:
        return None
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def mark_login_failed(script_name, password=None):
    """Record that script_name's login failed today. Dated (not just a
    boolean) so the block clears itself automatically tomorrow with no
    separate reset step. When password is given, its fingerprint (never the
    password itself) is stored alongside the date - see
    is_login_failed_today for why: a watchdog restarting this script
    shouldn't need to wait until tomorrow just because someone already fixed
    the password in .env five minutes after it expired."""
    LOGIN_FAILURE_DIR.mkdir(parents=True, exist_ok=True)
    payload = {"date": date.today().isoformat(), "password_fingerprint": _password_fingerprint(password)}
    _login_failure_marker_path(script_name).write_text(json.dumps(payload), encoding="utf-8")


def is_login_failed_today(script_name, password=None):
    """True only if script_name's login was marked as failed today
    specifically - a stale marker from a previous day (left over from before
    a password was fixed) must not keep blocking it forever. If password is
    given and its fingerprint differs from the one recorded at failure time,
    the block also lifts immediately (whoever runs this fixed the .env
    password since the failure, so there's no reason to keep waiting for a
    new calendar day before retrying)."""
    marker = _login_failure_marker_path(script_name)
    if not marker.exists():
        return False
    try:
        raw = marker.read_text(encoding="utf-8").strip()
    except OSError:
        return False
    try:
        payload = json.loads(raw)
    except ValueError:
        payload = {"date": raw, "password_fingerprint": None}  # pre-fingerprint marker format
    if payload.get("date") != date.today().isoformat():
        return False
    recorded = payload.get("password_fingerprint")
    if password is not None and recorded is not None and _password_fingerprint(password) != recorded:
        return False
    return True


def _paused_marker_path(script_name):
    return PAUSED_DIR / f"{script_name.replace(' ', '_')}.txt"


def is_paused(script_name):
    """True while script_name is paused from the control panel - the watchdog
    then leaves it stopped instead of relaunching it every check."""
    return _paused_marker_path(script_name).exists()


def set_paused(script_name, paused):
    marker = _paused_marker_path(script_name)
    if paused:
        PAUSED_DIR.mkdir(parents=True, exist_ok=True)
        marker.write_text(date.today().isoformat(), encoding="utf-8")
    else:
        marker.unlink(missing_ok=True)
