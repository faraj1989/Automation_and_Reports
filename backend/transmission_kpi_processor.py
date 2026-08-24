#!/usr/bin/env python3
"""
Libyana NPM - Transmission KPI Processor (IUB/ABIS backhaul)
Processes the daily site_BSC6900_kpI_PACKETLOSS report: hourly ping
delay/packet-loss per backhaul link (IUB = 3G NodeB<->RNC, ABIS = 2G
BTS<->BSC). The raw file is ~260k rows/day (24 hours x ~10,800 links), so
this module aggregates it down to one row per link per day before it's
handed to CSVHistoryManager for archiving - archiving the raw hourly grain
forever isn't practical.
"""

import os
import sys
import logging
import pandas as pd

# Add parent directory to path for imports when running standalone
if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.csv_loader import read_csv_skip_metadata
from backend.site_processor import find_file

logger = logging.getLogger(__name__)

TRANSMISSION_KPI_FILES = {
    'Transmission_KPIs': {
        'patterns': ['*BSC6900*PACKETLOSS*.csv', '*PACKETLOSS*.csv'],
        'sheet_name': 'Transmission_KPIs',
        'key_columns': ['Date', 'Adjacent Node ID'],
    }
}

# Raw hourly metric columns -> friendly daily-aggregate column names.
# Avg columns are meaned across the day's hourly readings, Max columns take
# the day's worst (max) hourly reading - mirrors the source report's own
# Average/Maximum split (T7816/T7812 = daily average KPIs, T7817/T7813 =
# daily maximum/worst-case KPIs).
_METRIC_AGG = {
    'T7816:Average Ping Packet Loss Rate of Adjacent Node(%)': ('Avg Packet Loss(%)', 'mean'),
    'T7817:Maximum Ping Packet Loss Rate of Adjacent Node(%)': ('Max Packet Loss(%)', 'max'),
    'T7812:Average Ping Delay of Adjacent Node(ms)': ('Avg Delay(ms)', 'mean'),
    'T7813:Maximum Ping Delay of Adjacent Node(ms)': ('Max Delay(ms)', 'max'),
}

# Identity columns carried through unaggregated (constant per link per day).
_IDENTITY_COLS = ['GBSC', 'Site Name', 'Adjacent Node Name', 'Adjacent Node Type', 'BTSID']

_OUTPUT_COLUMN_ORDER = [
    'Date', 'GBSC', 'Adjacent Node ID', 'Site Name', 'Adjacent Node Name',
    'Adjacent Node Type', 'BTSID', 'Backward Bandwidth', 'Integrity',
    'Avg Packet Loss(%)', 'Max Packet Loss(%)', 'Avg Delay(ms)', 'Max Delay(ms)',
]


def _aggregate_to_daily(df: pd.DataFrame) -> pd.DataFrame:
    """Collapse the raw hourly-per-link rows to one row per (Date, Adjacent
    Node ID) per day, matching the report's own Average/Maximum KPI split."""
    df = df.copy()

    df['Date'] = pd.to_datetime(df['Time'], errors='coerce').dt.strftime('%Y-%m-%d')
    df = df.dropna(subset=['Date', 'Adjacent Node ID'])

    # "100%" -> 100.0; take the day's worst (min) reading per link.
    df['Integrity'] = pd.to_numeric(
        df['Integrity'].astype(str).str.rstrip('%'), errors='coerce')

    df['Backward Bandwidth'] = pd.to_numeric(df['Backward Bandwidth'], errors='coerce')

    for raw_col in _METRIC_AGG:
        if raw_col in df.columns:
            df[raw_col] = pd.to_numeric(df[raw_col], errors='coerce')

    agg_spec = {col: 'first' for col in _IDENTITY_COLS if col in df.columns}
    agg_spec['Backward Bandwidth'] = 'first'
    agg_spec['Integrity'] = 'min'
    for raw_col, (_, how) in _METRIC_AGG.items():
        if raw_col in df.columns:
            agg_spec[raw_col] = how

    daily = df.groupby(['Date', 'Adjacent Node ID'], dropna=False, as_index=False).agg(agg_spec)

    rename_map = {raw_col: friendly for raw_col, (friendly, _) in _METRIC_AGG.items()}
    daily = daily.rename(columns=rename_map)

    ordered_cols = [c for c in _OUTPUT_COLUMN_ORDER if c in daily.columns]
    return daily[ordered_cols]


def process_transmission_kpis(day_folder, log_callback=None):
    """
    Find and process the daily Transmission KPI (IUB/ABIS packet loss +
    latency) file in the day folder, aggregated to one row per link per day.
    Returns a dictionary: sheet_name -> DataFrame (matches the shape of
    process_network_kpis/process_cell_kpis for CSVHistoryManager).
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📡 PROCESSING TRANSMISSION KPIs (IUB/ABIS)")
    log("=" * 60)

    results = {}

    for kpi_name, config in TRANSMISSION_KPI_FILES.items():
        patterns = config['patterns']
        sheet_name = config['sheet_name']

        file_path = find_file(day_folder, patterns)

        if not file_path:
            log(f"   ⚠️ No file found for {kpi_name}")
            results[sheet_name] = None
            continue

        log(f"📄 Processing {kpi_name}: {os.path.basename(file_path)}")
        try:
            df = read_csv_skip_metadata(file_path)
            if df is None or df.empty:
                log(f"   ⚠️ File is empty or could not be read")
                results[sheet_name] = None
                continue

            rows_before = len(df)
            df = df.dropna(how='all')
            empty_rows_removed = rows_before - len(df)
            if empty_rows_removed > 0:
                log(f"   🗑️  Removed {empty_rows_removed} completely empty rows")

            daily_df = _aggregate_to_daily(df)
            log(f"   ✅ Aggregated {rows_before} hourly rows -> {len(daily_df)} daily link rows")

            results[sheet_name] = daily_df
            if not daily_df.empty:
                dates = daily_df['Date'].unique()
                log(f"   📅 Dates: {min(dates)} to {max(dates)}")
        except Exception as e:
            log(f"   ❌ Error reading {kpi_name}: {e}")
            results[sheet_name] = None

    log("=" * 60)
    return results


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python transmission_kpi_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]
    results = process_transmission_kpis(test_folder)

    print("\n" + "=" * 60)
    print("TRANSMISSION KPI PROCESSING RESULTS")
    print("=" * 60)

    for sheet_name, df in results.items():
        print(f"\n{sheet_name}:")
        if df is not None and not df.empty:
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
            print(f"  Columns: {df.columns.tolist()}")
            print(f"  Sample:\n{df.head(5)}")
        else:
            print("  No data")
