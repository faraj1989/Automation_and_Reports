"""Password gate for the control panel. The hash lives in
control_panel/auth.json (gitignored) - PBKDF2-SHA256, never the password."""

import hashlib
import hmac
import json
import os
import secrets
from pathlib import Path

AUTH_FILE = Path(__file__).resolve().parent / "auth.json"
ITERATIONS = 240_000


def _hash(password, salt):
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), ITERATIONS).hex()


def is_configured():
    return AUTH_FILE.exists()


def set_password(password):
    salt = secrets.token_hex(16)
    tmp = AUTH_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps({"salt": salt, "hash": _hash(password, salt)}), encoding="utf-8")
    os.replace(tmp, AUTH_FILE)


def check_password(password):
    try:
        data = json.loads(AUTH_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    return hmac.compare_digest(_hash(password, data["salt"]), data["hash"])
