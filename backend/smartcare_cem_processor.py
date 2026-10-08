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


# ----------------------------------------------------------------------
# Monthly Comprehensive Analysis (Tripoli HQ) - the same 3 sheets as the
# history workbook, restricted to one calendar month. HQ asks for it at the
# start of each month; built on demand from the HQ Reports tab and
# automatically on the 2nd (scheduler.py --cem-monthly).
# ----------------------------------------------------------------------

MONTHLY_REPORT_DIR = os.environ.get(
    "CEM_MONTHLY_REPORT_DIR",
    os.path.join(os.path.dirname(os.path.dirname(CEM_WORKBOOK_PATH)), "Monthly_CEM_Reports"),
)
# HQ's Metrics format - the original 7 columns only (the *_weighted and
# source_rows columns stay in the internal history workbook).
HQ_METRICS_COLUMNS = [
    'date', 'tcp_connection_success_rate', 'downlink_tcp_retransmission_rate',
    'average_tcp_packet_loss_rate', 'downlink_tcp_packet_loss_rate',
    'tcp_connection_success_rate_included_rst', 'total_traffic_gb',
]
# Informational only: days with <85% of the month's typical row count. Cut-off
# days of 100,000-row exports are already dropped by the pipeline before they
# reach the history, and some real days are just smaller (e.g. 2026-08-08/09
# had ~3,450 rows in every export that covered them), so this never makes a
# month "incomplete" - it's shown so a human can eyeball it.
LOW_ROWS_DAY_SHARE = 0.85


def _month_bounds(month: str):
    start = pd.Timestamp(f"{month}-01")
    return start, start + pd.offsets.MonthEnd(0)


def monthly_report_filename(month: str) -> str:
    """'2026-09' -> 'Comprehensive_Analysis_Historical - ONLY -SEP-2026.xlsx'
    (the naming HQ already receives)."""
    start, _ = _month_bounds(month)
    return f"Comprehensive_Analysis_Historical - ONLY -{start.strftime('%b').upper()}-{start.year}.xlsx"


def _load_history_sheets():
    xl = pd.ExcelFile(CEM_WORKBOOK_PATH)
    top100 = xl.parse('Top100 per day')
    metrics = xl.parse('Metrics')
    for df in (top100, metrics):
        df['date'] = pd.to_datetime(df['date'], errors='coerce').dt.strftime('%Y-%m-%d')
    return top100, metrics


def available_months() -> list:
    """Months (YYYY-MM, newest first) with at least one day in the history."""
    if not is_available():
        return []
    _, metrics = _load_history_sheets()
    return sorted(metrics['date'].dropna().str[:7].unique().tolist(), reverse=True)


def month_completeness(month: str, metrics: Optional[pd.DataFrame] = None) -> Dict:
    """Which days of the month are present, missing, or look partial."""
    if metrics is None:
        _, metrics = _load_history_sheets()
    start, end = _month_bounds(month)
    all_days = pd.date_range(start, end).strftime('%Y-%m-%d').tolist()
    m = metrics[metrics['date'].str[:7] == month]
    present = sorted(set(m['date']))
    low_rows = []
    if 'source_rows' in m.columns and m['source_rows'].notna().any():
        typical = m['source_rows'].median()
        low_rows = sorted(m.loc[m['source_rows'] < LOW_ROWS_DAY_SHARE * typical, 'date'].tolist())
    missing = [d for d in all_days if d not in set(present)]
    return {
        'month': month,
        'days_in_month': len(all_days),
        'present': present,
        'missing': missing,
        'low_rows': low_rows,
        'complete': not missing,
    }


def build_monthly_cem_workbook(month: str):
    """(xlsx bytes, filename, completeness dict) for one month: Top100 per
    day, Top10 Application per week (re-rolled from THIS month's days only,
    with days_in_week so a week cut by the month boundary is visible), and
    Metrics (HQ's 7 columns)."""
    import io
    from reports.download_analysis_pipeline import build_top10_weekly_applications
    from backend.report_generator import autofit_excel_columns

    top100, metrics = _load_history_sheets()
    top = top100[top100['date'].str[:7] == month].copy()
    met = metrics[metrics['date'].str[:7] == month].copy()
    info = month_completeness(month, metrics)

    # Sort orders match the reference file HQ received for 2026-09: Top100
    # newest day first; weekly and Metrics oldest first.
    top = top.sort_values(['date', 'total_traffic_bytes'], ascending=[False, False])
    weekly = build_top10_weekly_applications(top).sort_values(
        ['week', 'total_traffic_bytes'], ascending=[True, False])
    met = met[[c for c in HQ_METRICS_COLUMNS if c in met.columns]].sort_values('date')

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for name, df in [('Top100 per day', top), ('Top10 Application per week', weekly), ('Metrics', met)]:
            df.to_excel(writer, sheet_name=name, index=False)
            writer.sheets[name].auto_filter.ref = writer.sheets[name].dimensions
        autofit_excel_columns(writer)
    return buf.getvalue(), monthly_report_filename(month), info


def save_monthly_cem_report(month: str, out_dir: str = MONTHLY_REPORT_DIR):
    """Write the month's workbook to out_dir; returns (path, completeness)."""
    data, filename, info = build_monthly_cem_workbook(month)
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, filename)
    with open(path, 'wb') as f:
        f.write(data)
    return path, info


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
