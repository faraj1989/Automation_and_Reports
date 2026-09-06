#!/usr/bin/env python3
"""
Libyana NPM - External Interference Processor
Processes the dedicated "2G_3G_4G interference and PRB utilization for
Automation" report (uploaded by Huawei iMaster daily at ~04:00, ready to
fetch by ~06:00): cell-level external interference for 2G (one row per
cell per day already) and 3G/4G (one row per cell per hour), plus bonus
4G PRB/throughput columns.

Same fix as hourly_cell_processor.py: the 3 per-technology CSVs inside
the report's zip have non-breaking-space suffixes in their filenames
("(2G\xa0Interference).csv" etc.), so files are identified by a
distinctive metric column in their header instead of by filename.
"""

import os
import glob
import logging

from backend.csv_loader import read_csv_skip_metadata

logger = logging.getLogger(__name__)

INTERFERENCE_REPORT_GLOB = '*interference*PRB*utilization*Automation*.csv'

# A column that uniquely identifies which technology's file this is.
TECH_IDENTITY_COLUMN = {
    'Interference Band Proportion (4~5)(%)': '2G_Interference',
    'VS.MeanRTWP': '3G_Interference_Hourly',
    'L.UL.Interference.Avg(dBm)': '4G_Interference_Hourly',
}


def _classify_sheet(columns) -> str:
    for col, sheet_name in TECH_IDENTITY_COLUMN.items():
        if col in columns:
            return sheet_name
    return None


def process_interference_kpis(day_folder, log_callback=None):
    """
    Find and process the 3 per-technology files in the dedicated
    interference+PRB report. Returns a dictionary: sheet_name -> DataFrame,
    at full native grain (2G: one row/cell/day; 3G/4G: one row/cell/hour -
    no aggregation here, that happens in build_external_interference_report).
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📡 PROCESSING EXTERNAL INTERFERENCE REPORT (2G/3G/4G)")
    log("=" * 60)

    results = {}

    candidates = glob.glob(os.path.join(day_folder, INTERFERENCE_REPORT_GLOB))
    if not candidates:
        log("   ⚠️ No interference+PRB report files found")
        return results

    for file_path in candidates:
        log(f"📄 Reading: {os.path.basename(file_path)}")
        try:
            df = read_csv_skip_metadata(file_path)
            if df is None or df.empty:
                log("   ⚠️ File is empty or could not be read")
                continue

            sheet_name = _classify_sheet(df.columns)
            if sheet_name is None:
                log("   ⚠️ Could not classify technology (no recognized identity column), skipping")
                continue

            rows_before = len(df)
            df = df.dropna(how='all')
            empty_rows_removed = rows_before - len(df)
            if empty_rows_removed > 0:
                log(f"   🗑️  Removed {empty_rows_removed} completely empty rows")

            results[sheet_name] = df
            log(f"   ✅ Classified as {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
        except Exception as e:
            log(f"   ❌ Error reading {os.path.basename(file_path)}: {e}")

    log("=" * 60)
    return results


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python interference_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]
    results = process_interference_kpis(test_folder)

    print("\n" + "=" * 60)
    print("EXTERNAL INTERFERENCE PROCESSING RESULTS")
    print("=" * 60)

    for sheet_name, df in results.items():
        print(f"\n{sheet_name}:")
        if df is not None and not df.empty:
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
            print(f"  Sample:\n{df.head(3)}")
        else:
            print("  No data")
