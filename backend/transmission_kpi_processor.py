#!/usr/bin/env python3
"""
Libyana NPM - Transmission KPI Processor (IUB/ABIS backhaul packet loss)

Reads the daily site_BSC6900_kpI_PACKETLOSS export - a rolling 7-day window
of hourly ping loss/delay per backhaul link (~260k rows) - and turns it into
four archives via backend/packet_loss_engine.py:

    Transmission_KPIs.csv        one row per link per day (key Date+GBSC+ID)
    Packet_Loss_Site_Daily.csv   one row per site per day (hour counts) - permanent
    Packet_Loss_Hub_Events.csv   one row per FN/HUB event per day - permanent
    Packet_Loss_Site_Hourly.csv  hourly site detail, rolling retention (drill-down)

Only complete past days are returned. Because every export repeats the last
7 days, a missed download self-heals on the next run, and the whole history
can be rebuilt from the day folders with backfill_transmission_history().
"""

import os
import re
import sys
import glob
import logging
from datetime import datetime, timedelta

# Add parent directory to path for imports when running standalone
if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.csv_loader import read_csv_skip_metadata
from backend.site_processor import find_file
from backend import packet_loss_engine as engine

logger = logging.getLogger(__name__)

TRANSMISSION_KPI_FILES = {
    'Transmission_KPIs': {
        'patterns': ['*BSC6900*PACKETLOSS*.csv', '*PACKETLOSS*.csv'],
        'sheet_name': 'Transmission_KPIs',
        'key_columns': ['Date', 'GBSC', 'Adjacent Node ID'],
    }
}
PATTERNS = TRANSMISSION_KPI_FILES['Transmission_KPIs']['patterns']

PACKET_LOSS_SHEETS = ['Transmission_KPIs', 'Packet_Loss_Site_Daily',
                      'Packet_Loss_Hub_Events', 'Packet_Loss_Site_Hourly']


def _process_file(file_path, rules=None, topology=None, only_dates=None, log=print):
    raw = read_csv_skip_metadata(file_path)
    if raw is None or raw.empty:
        log(f"   ⚠️ {os.path.basename(file_path)} is empty or could not be read")
        return {}
    raw = raw.dropna(how='all')
    results = engine.process_raw(raw, rules=rules, topology=topology, only_dates=only_dates)
    if results:
        dates = sorted(results['Packet_Loss_Site_Daily']['Date'].unique())
        log(f"   ✅ {len(raw):,} hourly rows -> {len(dates)} complete day(s) "
            f"{dates[0]} .. {dates[-1]}, {results['Packet_Loss_Site_Daily']['Site'].nunique()} sites, "
            f"{len(results['Packet_Loss_Hub_Events'])} hub-event row(s)")
    else:
        log("   ⚠️ No complete past days in this export")
    return results


def process_transmission_kpis(day_folder, log_callback=None):
    """Process the day folder's PACKETLOSS export. Returns {sheet: DataFrame}
    for CSVHistoryManager.update_transmission_kpis (see PACKET_LOSS_SHEETS)."""
    def log(msg):
        (log_callback or print)(msg)

    log("=" * 60)
    log("📡 PROCESSING TRANSMISSION / PACKET LOSS (IUB/ABIS)")
    log("=" * 60)
    file_path = find_file(day_folder, PATTERNS)
    if not file_path:
        log("   ⚠️ No PACKETLOSS file found")
        return {}
    log(f"📄 {os.path.basename(file_path)}")
    try:
        return _process_file(file_path, log=log)
    except Exception as e:
        log(f"   ❌ Error processing packet loss export: {e}")
        logger.exception("Packet loss processing failed")
        return {}
    finally:
        log("=" * 60)


def _export_end_date(path):
    """The export's end timestamp from its file name
    (..._20260726192821-20260929011020(2GBSC).csv -> 2026-09-29), or None."""
    m = re.search(r'-(\d{8})\d{6}', os.path.basename(path))
    try:
        return datetime.strptime(m.group(1), '%Y%m%d').date() if m else None
    except ValueError:
        return None


def backfill_transmission_history(local_root, history_mgr, log_callback=None):
    """Rebuild the packet-loss archives from every PACKETLOSS export still in
    the day folders. Newest export first; each date is taken from the newest
    export that fully covers it (latest counter revision wins), and an export
    whose 7 days are all covered already is skipped without being read."""
    def log(msg):
        (log_callback or print)(msg)

    files = []
    for folder in glob.glob(os.path.join(local_root, '*', 'unzipped')):
        f = find_file(folder, PATTERNS)
        if f:
            files.append(f)
    files.sort(key=lambda f: (_export_end_date(f) or datetime.min.date(), f), reverse=True)
    log(f"🔁 Packet loss backfill: {len(files)} export(s) found")

    rules, topology = engine.load_rules(), engine.load_topology()
    covered, batches = set(), []
    for f in files:
        end = _export_end_date(f)
        if end:
            expected = {(end - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(1, 8)}
            if expected <= covered:
                continue
        log(f"📄 {os.path.basename(f)}")
        try:
            res = _process_file(f, rules=rules, topology=topology, log=log)
            if not res:
                continue
            new_dates = set(res['Packet_Loss_Site_Daily']['Date']) - covered
            if not new_dates:
                continue
            res = {k: _filter_dates(k, v, new_dates) for k, v in res.items()}
            covered |= new_dates
            batches.append(res)
        except Exception as e:
            log(f"   ❌ {e}")

    if not batches:
        log("   Nothing to backfill")
        return 0
    import pandas as pd
    merged = {k: pd.concat([b[k] for b in batches if k in b], ignore_index=True) for k in PACKET_LOSS_SHEETS}
    history_mgr.update_transmission_kpis(merged)
    log(f"✅ Backfill done: {len(covered)} day(s) {min(covered)} .. {max(covered)}")
    return len(covered)


def _filter_dates(sheet, df, dates):
    if df is None or df.empty:
        return df
    if sheet == 'Packet_Loss_Site_Hourly':
        return df[df['Time'].str[:10].isin(dates)]
    return df[df['Date'].isin(dates)]


# ---------------------------- CLI ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    import argparse
    ap = argparse.ArgumentParser(description="Packet loss (IUB/ABIS) processing")
    ap.add_argument('path', help="day folder (…/unzipped) or, with --backfill, the local FTP root")
    ap.add_argument('--backfill', action='store_true', help="rebuild history from all day folders and save")
    args = ap.parse_args()

    if args.backfill:
        from backend.csv_history_manager import CSVHistoryManager
        backfill_transmission_history(args.path, CSVHistoryManager())
    else:
        for name, df in process_transmission_kpis(args.path).items():
            print(f"\n{name}: {len(df)} rows")
            print(df.head(5).to_string())
