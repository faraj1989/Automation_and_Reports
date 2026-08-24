#!/usr/bin/env python3
"""
Libyana NPM - Hourly All-Cells Processor
Processes the "all last hours all cells level" report: hourly, cell-level
KPIs for 2G/3G/4G in one rolling-window pull (currently ~120 hours), for
live/near-term cell monitoring - separate from the once-daily busy-hour
cell sheets (2G_Cell_CSBH/3G_Cell_CSBH/4G_Cell_BH) used for Scorecards.

The 3 per-technology CSVs inside the report's zip have inconsistent/odd
filenames (glob/`-like` matching on them proved unreliable while inspecting
this exact report - something non-standard in how MAE names them), so
files are identified by a distinctive column in their header instead of by
filename pattern.
"""

import os
import sys
import glob
import logging
import pandas as pd

# Add parent directory to path for imports when running standalone
if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.csv_loader import read_csv_skip_metadata

logger = logging.getLogger(__name__)

HOURLY_CELL_REPORT_GLOB = '*all last hours all cells level*.csv'

# A column that uniquely identifies which technology a file belongs to,
# since the filenames themselves aren't reliably pattern-matchable.
TECH_IDENTITY_COLUMN = {
    'GBSC': '2G_Cell_Hourly',
    'RNC': '3G_Cell_Hourly',
    'eNodeB Name': '4G_Cell_Hourly',
}


def _classify_sheet(columns) -> str:
    for col, sheet_name in TECH_IDENTITY_COLUMN.items():
        if col in columns:
            return sheet_name
    return None


def process_hourly_cell_kpis(day_folder, log_callback=None):
    """
    Find and process the 3 per-technology files in the hourly all-cells
    report. Returns a dictionary: sheet_name -> DataFrame, at full hourly
    grain (no aggregation - unlike transmission_kpi_processor.py, we want
    to keep every hour).
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📶 PROCESSING HOURLY ALL-CELLS REPORT (2G/3G/4G)")
    log("=" * 60)

    results = {}

    candidates = glob.glob(os.path.join(day_folder, HOURLY_CELL_REPORT_GLOB))
    if not candidates:
        log("   ⚠️ No hourly all-cells files found")
        return results

    for file_path in candidates:
        log(f"📄 Reading: {os.path.basename(file_path)}")
        try:
            df = read_csv_skip_metadata(file_path)
            if df is None or df.empty:
                log(f"   ⚠️ File is empty or could not be read")
                continue

            sheet_name = _classify_sheet(df.columns)
            if sheet_name is None:
                log(f"   ⚠️ Could not classify technology (no recognized identity column), skipping")
                continue

            rows_before = len(df)
            df = df.dropna(how='all')
            empty_rows_removed = rows_before - len(df)
            if empty_rows_removed > 0:
                log(f"   🗑️  Removed {empty_rows_removed} completely empty rows")

            results[sheet_name] = df
            log(f"   ✅ Classified as {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            if 'Time' in df.columns:
                times = df['Time'].dropna().unique()
                if len(times) > 0:
                    log(f"   📅 Time range: {min(times)} to {max(times)}")
        except Exception as e:
            log(f"   ❌ Error reading {os.path.basename(file_path)}: {e}")

    log("=" * 60)
    return results


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python hourly_cell_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]
    results = process_hourly_cell_kpis(test_folder)

    print("\n" + "=" * 60)
    print("HOURLY ALL-CELLS PROCESSING RESULTS")
    print("=" * 60)

    for sheet_name, df in results.items():
        print(f"\n{sheet_name}:")
        if df is not None and not df.empty:
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
            print(f"  Sample:\n{df.head(3)}")
        else:
            print("  No data")
