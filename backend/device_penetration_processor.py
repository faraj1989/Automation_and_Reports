#!/usr/bin/env python3
"""
Libyana NPM - Weekly Device Penetration Reader

Reads (read-only) Weekly_Device_Penetration_Historical.xlsx, produced by a
sibling NOC Automation Suite's weekly Device Penetration scraper (same
SmartCare portal/login as the CEM pipeline in smartcare_cem_processor.py,
different dashboard). See backend/noc_alarm_processor.py's docstring for
the general "sibling suite, read-only" pattern this follows.

Single sheet, one row per device model per weekly snapshot ('Time'):
Time, Device Model, Device Brand, Device Type, Device OS, Device
Technology, Number of Users(count), Penetration Rate(%). Network-wide,
no per-site breakdown.
"""
import os
import logging
from datetime import datetime
from typing import Dict, Optional

import pandas as pd

logger = logging.getLogger(__name__)

DEVICE_PENETRATION_PATH = os.environ.get(
    "DEVICE_PENETRATION_HISTORY_PATH",
    r"C:\Users\user\Desktop\Libyana_Data\Output\Weekly_Device_Penetration_Exports\Weekly_Device_Penetration_Historical.xlsx",
)

# Same weekly cadence as SmartCare CEM (schedule_config.json: Sunday 00:00)
# - see smartcare_cem_processor.CEM_STALE_DAYS for the same reasoning.
DEVICE_PENETRATION_STALE_DAYS = 10

# 4G/5G-capable vs. legacy-only, derived from the 'Device Technology'
# column (e.g. "2G/3G/LTE", "2G/3G/LTE/NR", "2G", "3G/LTE") - used for a
# quick network-wide "how much of the device base could actually use a
# newly activated LTE band" read, complementing the per-site 4G-band
# capacity advice in ReportGenerator.build_site_capacity_advice().
BROADBAND_TERMS = ("LTE", "NR")


def is_available() -> bool:
    return os.path.exists(DEVICE_PENETRATION_PATH)


def _file_age_days() -> Optional[float]:
    if not os.path.exists(DEVICE_PENETRATION_PATH):
        return None
    return (datetime.now().timestamp() - os.path.getmtime(DEVICE_PENETRATION_PATH)) / 86400.0


def load_device_penetration_overview() -> Dict:
    """Full history plus the latest snapshot's breakdowns by Device Type,
    Device Brand, and broadband (LTE/NR) capability - the sibling suite
    reruns this weekly, so a stale file here means a missed run upstream."""
    if not os.path.exists(DEVICE_PENETRATION_PATH):
        return {'loaded': False}
    try:
        df = pd.read_excel(DEVICE_PENETRATION_PATH, engine='openpyxl')
    except Exception as e:
        logger.warning(f"Could not read {DEVICE_PENETRATION_PATH}: {e}")
        return {'loaded': False}

    if df.empty or 'Time' not in df.columns:
        return {'loaded': False}

    latest_time = df['Time'].max()
    latest = df[df['Time'] == latest_time].copy()

    if 'Device Technology' in latest.columns:
        latest['Broadband Capable'] = latest['Device Technology'].apply(
            lambda v: any(t in str(v) for t in BROADBAND_TERMS) if pd.notna(v) else False
        )

    age_days = _file_age_days()
    return {
        'loaded': True,
        'age_days': age_days,
        'is_stale': age_days is not None and age_days > DEVICE_PENETRATION_STALE_DAYS,
        'last_updated': datetime.fromtimestamp(os.path.getmtime(DEVICE_PENETRATION_PATH)),
        'latest_time': latest_time,
        'latest_snapshot': latest,
        'history': df,
    }


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    overview = load_device_penetration_overview()
    print("Loaded:", overview.get('loaded'), "| age (days):", overview.get('age_days'))
    if overview.get('loaded'):
        snap = overview['latest_snapshot']
        print("Latest snapshot:", overview['latest_time'], "-", len(snap), "device models")
        if 'Broadband Capable' in snap.columns:
            broadband_users = snap.loc[snap['Broadband Capable'], 'Number of Users(count)'].sum()
            total_users = snap['Number of Users(count)'].sum()
            print(f"Broadband(LTE/NR)-capable users: {broadband_users:,} / {total_users:,} "
                  f"({100 * broadband_users / total_users:.1f}%)")
        print(snap.sort_values('Number of Users(count)', ascending=False).head(10).to_string())
