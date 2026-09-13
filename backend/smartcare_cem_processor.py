#!/usr/bin/env python3
"""
Libyana NPM - SmartCare CEM (Customer Experience Management) Reader

Reads (read-only) Comprehensive_Analysis_Historical.xlsx, produced by a
sibling NOC Automation Suite's SmartCare CEM scraper + analysis pipeline
(runs weekly, Sunday 00:00) - see backend/noc_alarm_processor.py's
docstring for the general "sibling suite, read-only" pattern this follows.

Unlike the alarm feed, this data is DPI/probe-based application traffic and
TCP-level connection quality - network-wide only, no per-site breakdown,
so it's a standalone dashboard section rather than wired into per-site
investigation the way the alarm feed is.

Three sheets:
  - 'Top100 per day'             - top 100 apps by traffic volume, one
                                    day's snapshot per group (date,
                                    application, bytes/GB)
  - 'Top10 Application per week' - same, weekly rollup, top 10 only
  - 'Metrics'                    - one row per day: TCP connection success
                                    rate, retransmission rate, packet loss
                                    rate(s), total traffic - a subscriber-
                                    experience quality trend independent of
                                    the network-counter KPIs elsewhere in
                                    this dashboard.
"""
import os
import logging
from datetime import datetime
from typing import Dict, Optional

import pandas as pd

logger = logging.getLogger(__name__)

CEM_WORKBOOK_PATH = os.environ.get(
    "SMARTCARE_CEM_HISTORY_PATH",
    r"C:\Users\user\Desktop\Libyana_Data\Output\Processed_Analysis\Comprehensive_Analysis_Historical.xlsx",
)

# SmartCare CEM runs weekly (Sunday 00:00) per the sibling suite's own
# schedule_config.json - a workbook older than 10 days means at least one
# weekly run was missed, not just "due later this week."
CEM_STALE_DAYS = 10


def is_available() -> bool:
    return os.path.exists(CEM_WORKBOOK_PATH)


def _file_age_days() -> Optional[float]:
    if not os.path.exists(CEM_WORKBOOK_PATH):
        return None
    return (datetime.now().timestamp() - os.path.getmtime(CEM_WORKBOOK_PATH)) / 86400.0


def load_cem_overview() -> Dict:
    """All three sheets plus a staleness flag - the sibling suite reruns
    this weekly, so a stale workbook here means a missed run upstream, not
    a bug in this reader."""
    if not os.path.exists(CEM_WORKBOOK_PATH):
        return {'loaded': False}
    try:
        xl = pd.ExcelFile(CEM_WORKBOOK_PATH)
        top100 = xl.parse('Top100 per day') if 'Top100 per day' in xl.sheet_names else pd.DataFrame()
        top10_week = xl.parse('Top10 Application per week') if 'Top10 Application per week' in xl.sheet_names else pd.DataFrame()
        metrics = xl.parse('Metrics') if 'Metrics' in xl.sheet_names else pd.DataFrame()
    except Exception as e:
        logger.warning(f"Could not read {CEM_WORKBOOK_PATH}: {e}")
        return {'loaded': False}

    if not metrics.empty and 'date' in metrics.columns:
        metrics = metrics.sort_values('date').reset_index(drop=True)

    age_days = _file_age_days()
    return {
        'loaded': True,
        'age_days': age_days,
        'is_stale': age_days is not None and age_days > CEM_STALE_DAYS,
        'last_updated': datetime.fromtimestamp(os.path.getmtime(CEM_WORKBOOK_PATH)),
        'top100_per_day': top100,
        'top10_per_week': top10_week,
        'metrics': metrics,
    }


def top_apps_for_date(top100_df: pd.DataFrame, target_date, n: int = 10) -> pd.DataFrame:
    """Top N applications by traffic for one date from the Top100 sheet."""
    if top100_df is None or top100_df.empty or 'date' not in top100_df.columns:
        return pd.DataFrame()
    day = top100_df[top100_df['date'].astype(str) == str(target_date)]
    if day.empty:
        return pd.DataFrame()
    sort_col = 'total_traffic_gb' if 'total_traffic_gb' in day.columns else day.columns[-1]
    return day.sort_values(sort_col, ascending=False).head(n)


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    overview = load_cem_overview()
    print("Loaded:", overview.get('loaded'), "| age (days):", overview.get('age_days'))
    if overview.get('loaded'):
        print("Metrics rows:", len(overview['metrics']))
        print(overview['metrics'].tail(3).to_string())
        latest_date = overview['top100_per_day']['date'].max() if not overview['top100_per_day'].empty else None
        print("\nTop apps for", latest_date)
        print(top_apps_for_date(overview['top100_per_day'], latest_date, 10).to_string())
