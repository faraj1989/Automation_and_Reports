#!/usr/bin/env python3
"""
Libyana NPM - History Integrity Check

Validates the daily KPI history CSVs in output/csv/ against the rule:

    Raw source exports may contain partial/current-day records;
    KPI history must contain only complete business days.

Per feed it checks:
  - no rows dated today or later (an incomplete day stored as history)
  - no duplicate business keys (e.g. two rows for the same Date)
  - latest stored date is the expected one (yesterday, minus any known
    per-feed report lag)
  - no missing dates inside the recent window

Runs at the end of every scheduler pipeline (results logged + written to
output/history_audit/), and standalone:

    python -m backend.history_integrity
"""

import os
import sys
import logging
from datetime import datetime, timedelta

import pandas as pd

if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

logger = logging.getLogger(__name__)

# feed -> (date column, business key columns)
HISTORY_FEEDS = {
    'SiteSummary': ('day', ['day']),
    'User_Summary': ('Date', ['Date']),
    'User_CS_Roaming': ('Date', ['Date']),
    'User_CS_Subscribers': ('Date', ['Date']),
    'User_PS_Roaming': ('Date', ['Date']),
    'User_PS_Subscribers': ('Date', ['Date']),
    'User_VoLTE': ('Date', ['Date']),
    '2G_NWBH': ('Date', ['Date', 'Whole Network']),
    '2G_NW_Daily': ('Date', ['Date', 'Whole Network']),
    '3G_NWBH': ('Date', ['Date', 'Whole Network']),
    '3G_NW_Daily': ('Date', ['Date', 'Whole Network']),
    '4G_NWBH': ('Date', ['Date', 'Whole Network']),
    '4G_NW_Daily': ('Date', ['Date', 'Whole Network']),
    'Gi_Interface_Traffic': ('Date', ['Date', 'Whole Network']),
    '2G_Cell_CSBH': ('Date', ['Date', 'GBSC', 'Cell CI']),
    '3G_Cell_CSBH': ('Date', ['Date', 'RNC', 'Cell ID']),
    '4G_Cell_BH': ('Date', ['Date', 'eNodeB Name', 'LocalCell Id']),
    'Traffic_Network_2G': ('Date', ['Date']),
    'Traffic_Network_3G': ('Date', ['Date']),
    'Traffic_Network_4G': ('Date', ['Date']),
    'Traffic_2G': ('Date', ['Date', 'Site']),
    'Traffic_3G': ('Date', ['Date', 'Site']),
    'Traffic_4G': ('Date', ['Date', 'Site']),
    'Transmission_KPIs': ('Date', ['Date', 'GBSC', 'Adjacent Node ID']),
    'Packet_Loss_Site_Daily': ('Date', ['Date', 'Site']),
}

# Only look for missing dates this far back from the expected latest date -
# older gaps are already known/unrecoverable and would just be noise.
GAP_WINDOW_DAYS = 30


def _expected_lag(feed):
    try:
        from backend.report_generator import FRESHNESS_EXPECTED_LAG
        return FRESHNESS_EXPECTED_LAG.get(feed, 0)
    except Exception:
        return 0


def check_history_integrity(csv_folder=os.path.join("output", "csv"), today=None):
    """Return a DataFrame with one row per feed:
    Feed, Rows, Latest Date, Expected Latest, Future Rows, Duplicate Keys,
    Missing Dates (recent window), Status (PASS/WARN/FAIL), Issues."""
    today = pd.Timestamp(today or datetime.now().date()).normalize()
    rows = []
    for feed, (date_col, key_cols) in HISTORY_FEEDS.items():
        expected = today - timedelta(days=1 + _expected_lag(feed))
        rec = {'Feed': feed, 'Rows': 0, 'Latest Date': 'N/A',
               'Expected Latest': expected.strftime('%Y-%m-%d'),
               'Future Rows': 0, 'Duplicate Keys': 0, 'Missing Dates': 0,
               'Status': 'PASS', 'Issues': ''}
        issues, fail = [], False
        path = os.path.join(csv_folder, f"{feed}.csv")
        try:
            if not os.path.exists(path):
                raise FileNotFoundError("file missing")
            df = pd.read_csv(path, usecols=lambda c: c in key_cols)
            missing_keys = [c for c in key_cols if c not in df.columns]
            if missing_keys:
                raise ValueError(f"key column(s) missing: {missing_keys}")
            rec['Rows'] = len(df)
            dates = pd.to_datetime(df[date_col], errors='coerce', format='mixed').dt.normalize()

            unparsable = int(dates.isna().sum())
            if unparsable:
                issues.append(f"{unparsable} unparsable date(s)")
                fail = True

            future = dates >= today
            rec['Future Rows'] = int(future.sum())
            if rec['Future Rows']:
                fut = sorted(dates[future].dt.strftime('%Y-%m-%d').unique())
                issues.append(f"incomplete/future day(s) stored: {', '.join(fut)}")
                fail = True

            keyed = df.assign(**{date_col: dates.dt.strftime('%Y-%m-%d')})
            dup = keyed.duplicated(subset=key_cols, keep=False)
            rec['Duplicate Keys'] = int(keyed[dup].drop_duplicates(subset=key_cols).shape[0])
            if rec['Duplicate Keys']:
                issues.append(f"{rec['Duplicate Keys']} duplicated key(s)")
                fail = True

            valid = dates.dropna()
            if not valid.empty:
                latest = valid.max()
                rec['Latest Date'] = latest.strftime('%Y-%m-%d')
                if latest < expected:
                    issues.append(f"stale: {(expected - latest).days} day(s) behind expected")

                window_start = max(valid.min(), expected - timedelta(days=GAP_WINDOW_DAYS - 1))
                window_end = min(latest, expected)
                if window_end >= window_start:
                    have = set(valid[(valid >= window_start) & (valid <= window_end)])
                    missing = [d for d in pd.date_range(window_start, window_end) if d not in have]
                    rec['Missing Dates'] = len(missing)
                    if missing:
                        shown = ', '.join(d.strftime('%m-%d') for d in missing[:8])
                        more = f" (+{len(missing) - 8} more)" if len(missing) > 8 else ''
                        issues.append(f"missing date(s) in last {GAP_WINDOW_DAYS}d: {shown}{more}")
            else:
                issues.append("no valid dates")
                fail = True
        except Exception as e:
            issues.append(str(e))
            fail = True

        rec['Status'] = 'FAIL' if fail else ('WARN' if issues else 'PASS')
        rec['Issues'] = '; '.join(issues)
        rows.append(rec)
    return pd.DataFrame(rows)


def run_and_log(csv_folder=os.path.join("output", "csv"), audit_folder=os.path.join("output", "history_audit")):
    """Scheduler hook: run the check, log a summary, save the table. Never raises."""
    try:
        result = check_history_integrity(csv_folder)
        os.makedirs(audit_folder, exist_ok=True)
        out = os.path.join(audit_folder, f"integrity_{datetime.now():%Y-%m-%d}.csv")
        result.to_csv(out, index=False)
        counts = result['Status'].value_counts().to_dict()
        logger.info(f"History integrity: {counts.get('PASS', 0)} PASS, {counts.get('WARN', 0)} WARN, "
                    f"{counts.get('FAIL', 0)} FAIL - saved {out}")
        for _, r in result[result['Status'] != 'PASS'].iterrows():
            log = logger.error if r['Status'] == 'FAIL' else logger.warning
            log(f"   {r['Status']} {r['Feed']}: {r['Issues']}")
        return result
    except Exception as e:
        logger.error(f"History integrity check failed to run: {e}")
        return None


if __name__ == "__main__":
    pd.set_option('display.width', 250)
    pd.set_option('display.max_colwidth', 90)
    res = check_history_integrity()
    print(res.to_string(index=False))
    sys.exit(1 if (res['Status'] == 'FAIL').any() else 0)
