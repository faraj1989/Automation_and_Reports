#!/usr/bin/env python3
"""
iNOC (Whale Cloud ZSmart OFM) ticket sync - READ-ONLY.

Pulls the logged-in user's "My Task" queue from https://inoc.libyana.ly/oss/
plus per-ticket detail (site, customer/site coordinates, fault kind, last
suggestion) and same-site complaint history, and writes them under
output/inoc/ for the dashboard and complaint analysis.

iNOC login needs an SMS one-time code and has no "remember me", so this
script never logs in. A person logs in once in a dedicated Chrome window
(`--launch`) that listens on a local debug port; each sync attaches to that
window and calls the portal's own JSON API from inside the page (so it uses
the page's session cookie and CSRF token). Regular syncs also keep the
session from idling out.

Only endpoints in READ_ENDPOINTS can be called - the ticket-changing ones
(workorder/function, checkout, saveDraft, ...) are deliberately absent so a
bug here cannot process, reassign or comment on a ticket.

usage:
  python scrapers/inoc_ticket_sync.py --launch   # open Chrome, then log in by hand
  python scrapers/inoc_ticket_sync.py            # sync (exit 2 = not logged in, 3 = unreachable)
"""
import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT_DIR))

from project_config import env_int, env_path_str, env_str, load_env_file  # noqa: E402
from project_logging import setup_logger  # noqa: E402

load_env_file()
log = setup_logger("inoc_ticket_sync")

INOC_URL = env_str("INOC_URL", "https://inoc.libyana.ly/oss/")
DEBUG_PORT = env_int("INOC_DEBUG_PORT", 9333)
CHROME_EXE = env_str("INOC_CHROME_EXE", r"C:\Program Files\Google\Chrome\Application\chrome.exe")
# Kept outside the repo: the profile holds the live iNOC session cookie.
PROFILE_DIR = env_path_str("INOC_CHROME_PROFILE",
                           str(Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "LibyanaNPM" / "inoc_chrome_profile"))
OUTPUT_DIR = ROOT_DIR / "output" / "inoc"
QUEUE_CSV = OUTPUT_DIR / "inoc_my_tasks.csv"
DETAIL_CACHE = OUTPUT_DIR / "inoc_detail_cache.json"
STATUS_JSON = OUTPUT_DIR / "inoc_sync_status.json"
SNAPSHOT_DIR = OUTPUT_DIR / "snapshots"

PAGE_SIZE = 200
DETAIL_BATCH = 25  # tickets per in-page round trip

READ_ENDPOINTS = {
    "ofm/api/requesttype/priv/list/v1",
    "ofm/api/libyana/requesttype/statistics/requestList/count/v1",
    "ofm/api/libyana/request/list/v1",
    "ofm/api/servicedesk/detail/v1",
    "ofm/api/libyana/customerComplaint/queryCustomerComplaintHisListByRequestNo/v1",
}

# Bump when the cached detail shape changes, so old entries are re-fetched.
CACHE_VERSION = 2

# Ticket-page labels (Basic Information / Customer Information /
# Classification, in on-screen order) -> detail field. Fields the portal
# leaves out when empty (Finish Date, Contact Phone, Phenomenon, Trouble
# Reason, Activity Description) are listed with their best-known codes and
# stay blank until a ticket carries them.
PAGE_FIELDS = [
    ("Request No.", "REQUEST_NO"),
    ("Request Type", "REQUEST_TYPE_NAME"),
    ("Status", "TO_STATUS"),
    ("Create User", "CREATE_USER_NAME"),
    ("Create Date", "CREATE_DATE"),
    ("Finish Date", "FINISH_DATE"),
    ("Site Latitude", "SITE_LATITUDE"),
    ("Site Longitude", "SITE_LONGITUDE"),
    ("Site Name", "SITE_NAME"),
    ("Customer Latitude", "CUSTOMER_LATITUDE"),
    ("Customer Longitude", "CUSTOMER_LONGITUDE"),
    ("Requestor", "REPORTED_USER_NAME"),
    ("Email of Requestor", "CALLBACK_VAL"),
    ("Phone of Requestor", "CALLBACK_TEL"),
    ("Summary", "SUMMARY"),
    ("Description", "DESCRIPTION"),
    ("Suggestion", "SUGGESTION"),
    ("Activity Description", "ACTIVITY_DESC"),
    ("Access No", "ACCESS_NO"),
    ("Customer Name", "CUST_NAME"),
    ("Customer Type", "CUST_TYPE"),
    ("Contact Phone", "CONTACT_PHONE"),
    ("Customer Address", "CUST_ADDRESS"),
    ("Trouble Type", "FAULT_KIND_NAME"),
    ("Priority", "PRIORITY"),
    ("Resolution Warn SLA", "TIME_WARN_DATE"),
    ("Resolution SLA", "TIME_LIMIT_DATE"),
    ("Phenomenon", "FAULT_PHENOMENA_NAME"),
    ("Trouble Reason", "FAULT_REASON_NAME"),
    ("Assign to Org Level", "ASSIGN_LEVEL"),
    ("Assign to Department", "HANDLE_ORG_NAME"),
]
# Workflow columns from the queue list (not shown in the Basic Information panel).
WORKFLOW_FIELDS = [
    ("Previous Activity", "PRE_ACTIVITY_NAME"),
    ("Activity", "ACTIVITY_NAME"),
    ("Order State", "ORDER_STATE_NAME"),
    ("Handler", "HANDLER"),
    ("Owner", "OWNER_NAME"),
    ("Issue Solved", "ISSUE_SOLVED"),
    ("Next Action Level", "NEXT_ACTION_TO_LEVEL"),
    ("Area", "AREA_NAME"),
    ("Happen Date", "HAPPEN_DATE"),
    ("Modify Date", "MODIFY_DATE"),
    ("Workorder No.", "WORKORDER_NO"),
    ("Same Ticket No.", "SAME_TICKET_NO"),
]
# Internal ids/flags with no meaning outside the portal - dropped from the
# "everything else" tail of the CSV (still kept in the detail cache).
INTERNAL_FIELDS = {
    "ACTIVITY_ID", "ACTIVITY_NO", "ASSIGN_ORG", "CALLBACK_METHOD", "CATALOG_ID", "CATEGORY_ID",
    "CREATE_USER", "DEPLOY_FLAG", "FUNCTION_CODE", "FUNCTION_ID", "MAIN_FLAG", "MODIFY_USER",
    "NOT_FIRST", "OPER_STATE", "ORDER_STATE", "OWNER", "OWNER_TYPE", "PAGE_CODE",
    "PROCESSINSTANCE_NO", "PROCESS_FINISH_FLAG", "PROCESS_NO", "PROJECT_ID", "REPORTED_SOURCE",
    "REPORTED_USER", "REQUEST_TYPE", "REQUEST_TYPE_CATALOG", "SEQ", "SERVICE_TYPE", "SLA_ID",
    "SP_ID", "STATE", "WORKORDER_ASSIGN_STAFF", "WORKORDER_FROM_STATUS", "WORKORDER_STATE",
    "CUSTOMER_COMPLAINT_LIST", "RN", "CREATE_JOB", "CREATE_ORG", "CREATE_STAFF", "ASSIGN_STAFF",
    "PRE_ACTIVITY_ID", "PRE_ACTIVITY_NO", "COLOR", "TASK_HANDLE", "TIME_LIMIT_DATE_SORT",
}

# Runs inside the iNOC page. arguments: [calls, readEndpoints]; each call is
# [endpoint, body]. Returns one {ok, data|err} per call.
_JS_CALLS = r"""
var done = arguments[arguments.length - 1];
var calls = arguments[0], allowed = arguments[1];
(async function () {
  if (!(window.portal && portal.appGlobal && portal.appGlobal.get('staffId'))) return {loggedIn: false};
  var csrf = portal.appGlobal.get('_csrf'), out = [];
  for (var i = 0; i < calls.length; i++) {
    var ep = calls[i][0], body = calls[i][1];
    if (allowed.indexOf(ep) < 0) { out.push({ok: false, err: 'endpoint not whitelisted: ' + ep}); continue; }
    try {
      var r = await fetch(ep, {method: 'POST', credentials: 'same-origin',
        headers: {'Content-Type': 'application/json', 'X-CSRF-TOKEN': csrf}, body: JSON.stringify(body)});
      var t = await r.text();
      if (!r.ok) { out.push({ok: false, err: 'HTTP ' + r.status + ' ' + t.slice(0, 200)}); continue; }
      out.push({ok: true, data: JSON.parse(t)});
    } catch (e) { out.push({ok: false, err: String(e)}); }
  }
  return {loggedIn: true, staffName: portal.appGlobal.get('staffName'), results: out};
})().then(done, function (e) { done({loggedIn: null, err: String(e)}); });
"""


class NotLoggedIn(RuntimeError):
    pass


def _port_open() -> bool:
    import socket
    with socket.socket() as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", DEBUG_PORT)) == 0


def launch_chrome():
    if _port_open():
        log.info("iNOC Chrome already running on port %s - log in there if needed.", DEBUG_PORT)
        return
    Path(PROFILE_DIR).mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen([
        CHROME_EXE, f"--remote-debugging-port={DEBUG_PORT}", f"--user-data-dir={PROFILE_DIR}",
        "--no-first-run", "--ignore-certificate-errors", "--window-size=1600,1000", INOC_URL,
    ])
    log.info("Launched iNOC Chrome (PID %s). Log in with username, password and SMS code.", proc.pid)


def attach():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    if not _port_open():
        raise NotLoggedIn(f"iNOC Chrome is not running on port {DEBUG_PORT} (run with --launch)")
    opts = Options()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
    driver = webdriver.Chrome(options=opts)
    driver.set_script_timeout(300)
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if "inoc.libyana.ly" in driver.current_url:
            return driver
    raise NotLoggedIn("no iNOC tab open in the iNOC Chrome window")


def call(driver, calls):
    """Run [endpoint, body] calls in the page; returns the data payloads in order."""
    for ep, _ in calls:
        if ep not in READ_ENDPOINTS:
            raise ValueError(f"refusing non-read endpoint: {ep}")
    res = driver.execute_async_script(_JS_CALLS, calls, sorted(READ_ENDPOINTS))
    if not res or res.get("loggedIn") is False:
        raise NotLoggedIn("iNOC session is not logged in (log in again in the iNOC Chrome window)")
    if res.get("loggedIn") is None:
        raise RuntimeError(f"in-page call failed: {res.get('err')}")
    out = []
    for (ep, _), r in zip(calls, res["results"]):
        if not r["ok"]:
            if "Failed to fetch" in r["err"]:
                raise ConnectionError(f"iNOC server unreachable from this PC (network/VPN?) at {ep}")
            raise RuntimeError(f"{ep}: {r['err']}")
        data = r["data"]
        if isinstance(data, dict) and data.get("resultCode") not in (None, "0"):
            raise RuntimeError(f"{ep}: {data.get('resultCode')} {data.get('resultDesc')}")
        out.append(data)
    return out, res.get("staffName")


def fetch_queue(driver):
    types = call(driver, [["ofm/api/requesttype/priv/list/v1", {"SERVICE_NAME": "QOFM_REQUEST_TYPE"}]])[0][0]
    type_ids = ",".join(t["REQUEST_TYPE"] for t in types["resultData"])
    (counts,), staff = call(driver, [["ofm/api/libyana/requesttype/statistics/requestList/count/v1", {
        "COUNT_FLAG": "1", "CONDITION": "ALLMYTASK", "type": "STA", "SPID": "0", "requestType": type_ids}]])
    counters = counts["resultData"][0] if counts["resultData"] else {}
    rows = []
    for key, n in counters.items():
        if not key.startswith("mytask_") or int(n or 0) == 0:
            continue
        for start in range(1, int(n) + 1, PAGE_SIZE):
            (page,), _ = call(driver, [["ofm/api/libyana/request/list/v1", {
                "CONDITION": key, "QUERY_METHOD": "FULL_FUZZY", "QUERY": None, "REQUEST_NO": None,
                "PAGE_SIZE": str(PAGE_SIZE), "START_POS": str(start), "END_POS": str(start + PAGE_SIZE - 1),
                "SPID": "0"}]])
            rows.extend(page["resultData"]["DR"])
        log.info("%s: %s tickets listed", key, n)
    return rows, staff


def fetch_details(driver, rows, cache):
    """Fill cache[WORKORDER_NO] for tickets that are new or modified since last sync."""
    def stale(r):
        c = cache.get(r["WORKORDER_NO"], {})
        return c.get("V") != CACHE_VERSION or c.get("MODIFY_DATE") != r.get("MODIFY_DATE")
    todo = [r for r in rows if stale(r)]
    log.info("details: %d cached, %d to fetch", len(rows) - len(todo), len(todo))
    for i in range(0, len(todo), DETAIL_BATCH):
        batch = todo[i:i + DETAIL_BATCH]
        calls = []
        for r in batch:
            calls.append(["ofm/api/servicedesk/detail/v1", {
                "REQUEST_NO": r["REQUEST_NO"], "WORKORDER_NO": r["WORKORDER_NO"], "SPID": "0", "LANG_ID": "en"}])
            calls.append(["ofm/api/libyana/customerComplaint/queryCustomerComplaintHisListByRequestNo/v1",
                          {"requestNo": r["REQUEST_NO"]}])
        data, _ = call(driver, calls)
        for j, r in enumerate(batch):
            d = (data[2 * j].get("resultData") or {}).get("DRDATA") or {}
            hist = data[2 * j + 1].get("resultData") or []
            entry = {"V": CACHE_VERSION, "MODIFY_DATE": r.get("MODIFY_DATE"), "DETAIL": d}
            entry["SITE_HISTORY"] = [
                {k: h.get(k) for k in ("COMPLAINT_TICKET_NO", "ACCEPTANCE_DATE", "TROUBLE_TYPE", "TITLE", "SITE_NAME")}
                for h in hist]
            cache[r["WORKORDER_NO"]] = entry
        log.info("details: %d/%d", min(i + DETAIL_BATCH, len(todo)), len(todo))
        time.sleep(0.5)  # be gentle with the production portal
    return cache


def build_table(rows, cache):
    """One row per ticket: page-labelled columns, workflow columns, then every
    other non-internal field the portal returned (raw field names)."""
    records = []
    for r in rows:
        c = cache.get(r["WORKORDER_NO"], {})
        rec = {**r, **{k: v for k, v in c.get("DETAIL", {}).items() if v not in (None, "")}}
        hist = c.get("SITE_HISTORY", [])
        rec["_HIST_COUNT"] = len(hist)
        rec["_HIST_TICKETS"] = "; ".join(f"{h.get('COMPLAINT_TICKET_NO')} ({h.get('TITLE', '').strip()})"
                                         for h in hist)
        records.append(rec)
    raw = pd.DataFrame(records)

    out = pd.DataFrame(index=raw.index)
    used = set()
    for label, field in PAGE_FIELDS + WORKFLOW_FIELDS:
        out[label] = raw[field] if field in raw else None
        used.add(field)
    # LIMIT_FLAG = days left until the resolution SLA (negative = overdue).
    out["Days To SLA"] = pd.to_numeric(raw.get("LIMIT_FLAG"), errors="coerce").round(2)
    out["Age Days"] = (pd.Timestamp.now() - pd.to_datetime(raw["CREATE_DATE"], errors="coerce")).dt.days
    out["Site Complaint History Count"] = raw["_HIST_COUNT"]
    out["Site Complaint History"] = raw["_HIST_TICKETS"]
    used |= {"LIMIT_FLAG", "_HIST_COUNT", "_HIST_TICKETS"}
    for field in sorted(set(raw.columns) - used - INTERNAL_FIELDS):
        out[field] = raw[field]

    for c in ("Site Latitude", "Site Longitude", "Customer Latitude", "Customer Longitude"):
        out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


def write_status(**kw):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    kw["checked_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    STATUS_JSON.write_text(json.dumps(kw, indent=1, ensure_ascii=False), encoding="utf-8")


def sync():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    cache = json.loads(DETAIL_CACHE.read_text(encoding="utf-8")) if DETAIL_CACHE.exists() else {}
    driver = attach()  # attaching to an existing browser: never quit() it
    rows, staff = fetch_queue(driver)
    cache = fetch_details(driver, rows, cache)
    live = {r["WORKORDER_NO"] for r in rows}
    cache = {k: v for k, v in cache.items() if k in live}  # drop tickets that left the queue
    DETAIL_CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")

    df = build_table(rows, cache)
    tmp = QUEUE_CSV.with_suffix(".tmp")
    df.to_csv(tmp, index=False, encoding="utf-8-sig")
    try:
        os.replace(tmp, QUEUE_CSV)
    except PermissionError:  # usually the CSV is open in Excel
        alt = QUEUE_CSV.with_name(f"{QUEUE_CSV.stem}_{datetime.now():%Y%m%d_%H%M%S}.csv")
        os.replace(tmp, alt)
        log.warning("%s is locked (open in Excel?) - wrote %s instead", QUEUE_CSV.name, alt.name)
    df.to_csv(SNAPSHOT_DIR / f"inoc_my_tasks_{datetime.now():%Y%m%d}.csv", index=False, encoding="utf-8-sig")
    overdue = int((df["Days To SLA"] < 0).sum())
    write_status(logged_in=True, ok=True, staff=staff, tickets=len(df), overdue=overdue)
    log.info("sync done: %d tickets (%d overdue), %d columns -> %s", len(df), overdue, df.shape[1], QUEUE_CSV)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--launch", action="store_true", help="open the dedicated iNOC Chrome window for manual login")
    args = ap.parse_args()
    if args.launch:
        launch_chrome()
        return 0
    try:
        sync()
        return 0
    except NotLoggedIn as e:
        log.warning("%s", e)
        write_status(logged_in=False, ok=False, error=str(e))
        return 2
    except ConnectionError as e:
        log.warning("%s", e)
        write_status(logged_in=None, ok=False, unreachable=True, error=str(e))
        return 3
    except Exception as e:
        log.exception("iNOC sync failed")
        write_status(logged_in=None, ok=False, error=str(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
