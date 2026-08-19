#!/usr/bin/env python3
"""
Libyana NPM - Cell KPI Processor
Processes 2G, 3G, 4G cell-level busy hour KPIs (CSBH/BH).
"""

import os
import sys
import glob
import logging
from datetime import datetime
import pandas as pd

# Add parent directory to path for imports when running standalone
if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from backend.csv_loader import read_csv_skip_metadata
from backend.site_processor import find_file

logger = logging.getLogger(__name__)

# Cell KPI file patterns
CELL_KPI_FILES = {
    '2G_Cell_CSBH': {
        'patterns': ['* (2G cell-CSBH).csv', '*(2G cell-CSBH).csv', '*2G cell-CSBH*.csv', '*2G cell-CSBH*.csv'],
        'sheet_name': '2G_Cell_CSBH',
        'key_columns': ['Date', 'Cell Name']
    },
    '3G_Cell_CSBH': {
        'patterns': ['* (3G-cells -CSBH).csv', '*(3G-cells -CSBH).csv', '*3G-cells -CSBH*.csv', '*3G-cells -CSBH*.csv'],
        'sheet_name': '3G_Cell_CSBH',
        'key_columns': ['Date', 'Cell Name']
    },
    '4G_Cell_BH': {
        'patterns': ['* (4G cell-BH).csv', '*(4G cell-BH).csv', '*4G cell-BH*.csv', '*4G cell-BH*.csv'],
        'sheet_name': '4G_Cell_BH',
        'key_columns': ['Date', 'Cell Name']
    }
}


def process_cell_kpis(day_folder, log_callback=None):
    """
    Process all 3 cell KPI files in the day folder.
    Returns a dictionary: sheet_name -> DataFrame
    Handles: Date normalization, empty row removal, duplicate detection
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📊 PROCESSING CELL KPIs (CSBH/BH)")
    log("=" * 60)

    results = {}

    for kpi_name, config in CELL_KPI_FILES.items():
        patterns = config['patterns']
        sheet_name = config['sheet_name']

        # Find the file
        file_path = find_file(day_folder, patterns)

        if file_path:
            log(f"📄 Processing {kpi_name}: {os.path.basename(file_path)}")
            try:
                df = read_csv_skip_metadata(file_path)
                if df is not None and not df.empty:
                    rows_before = len(df)
                    
                    # 1. Remove completely empty rows
                    df = df.dropna(how='all')
                    empty_rows_removed = rows_before - len(df)
                    if empty_rows_removed > 0:
                        log(f"   🗑️  Removed {empty_rows_removed} completely empty rows")
                    
                    # 2. Normalize Date column to YYYY-MM-DD format
                    if 'Date' in df.columns:
                        try:
                            df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
                            log(f"   ✅ Normalized dates to YYYY-MM-DD format")
                        except Exception as e:
                            log(f"   ⚠️ Could not normalize dates: {e}")

                    # 3. Remove duplicates within the file
                    key_cols = config.get('key_columns', ['Date'])
                    available_keys = [col for col in key_cols if col in df.columns]
                    if available_keys:
                        dup_count = df.duplicated(subset=available_keys).sum()
                        if dup_count > 0:
                            log(f"   ⚠️ Found {dup_count} duplicate rows by {available_keys}, removing")
                            df = df.drop_duplicates(subset=available_keys)

                    results[sheet_name] = df
                    log(f"   ✅ Loaded {len(df)} rows, {len(df.columns)} columns")
                    if 'Date' in df.columns:
                        dates = df['Date'].unique()
                        if len(dates) > 0:
                            log(f"   📅 Dates: {min(dates)} to {max(dates)}")
                else:
                    log(f"   ⚠️ File is empty or could not be read")
                    results[sheet_name] = None
            except Exception as e:
                log(f"   ❌ Error reading {kpi_name}: {e}")
                results[sheet_name] = None
        else:
            log(f"   ⚠️ No file found for {kpi_name}")
            results[sheet_name] = None

    log("=" * 60)
    return results


def get_cell_kpi_summary(df, kpi_name):
    """
    Extract a summary of key KPIs from a cell KPI DataFrame.
    """
    if df is None or df.empty:
        return {}

    summary = {}
    try:
        if kpi_name == '2G_Cell_CSBH':
            # Key 2G cell KPIs
            if 'RR307:TCH Availability(%)' in df.columns:
                summary['Avg TCH Availability(%)'] = df['RR307:TCH Availability(%)'].mean()
            if 'Call Setup Success Rate(%)' in df.columns:
                summary['Avg Call Setup Success Rate(%)'] = df['Call Setup Success Rate(%)'].mean()
            if 'K3014:Traffic Volume on TCH(Erl)' in df.columns:
                summary['Total TCH Traffic(Erl)'] = df['K3014:Traffic Volume on TCH(Erl)'].sum()
            if 'PS Traffic (RLC)(MB)' in df.columns:
                summary['Total PS Traffic(MB)'] = df['PS Traffic (RLC)(MB)'].sum()

        elif kpi_name == '3G_Cell_CSBH':
            if 'CS Call Drop Rate(%)' in df.columns:
                summary['Avg CS Call Drop Rate(%)'] = df['CS Call Drop Rate(%)'].mean()
            if 'Soft Handover Success Rate(%)' in df.columns:
                summary['Avg Soft Handover Success Rate(%)'] = df['Soft Handover Success Rate(%)'].mean()
            if 'Availability' in df.columns:
                summary['Avg Availability(%)'] = df['Availability'].mean()
            if 'HSDPA Throughput per user (Local Cell)(Kbps)' in df.columns:
                # Convert to numeric, handling NIL values
                throughput = pd.to_numeric(df['HSDPA Throughput per user (Local Cell)(Kbps)'], errors='coerce')
                summary['Avg HSDPA Throughput(Kbps)'] = throughput.mean()

        elif kpi_name == '4G_Cell_BH':
            if 'RRC Setup Success Rate(%)' in df.columns:
                summary['Avg RRC Setup Success Rate(%)'] = df['RRC Setup Success Rate(%)'].mean()
            if 'Service Drop Rate (All)' in df.columns:
                summary['Avg Service Drop Rate(%)'] = df['Service Drop Rate (All)'].mean()
            if 'Downlink Traffic Volume(GB)' in df.columns:
                summary['Total DL Traffic(GB)'] = df['Downlink Traffic Volume(GB)'].sum()
            if 'Uplink Traffic Volume (GB)' in df.columns:
                summary['Total UL Traffic(GB)'] = df['Uplink Traffic Volume (GB)'].sum()
            if 'VoLTE Traffic Volume (Erl)' in df.columns:
                summary['Total VoLTE Traffic(Erl)'] = df['VoLTE Traffic Volume (Erl)'].sum()
            if 'User Downlink Average Throughput (Mbps)' in df.columns:
                throughput = pd.to_numeric(df['User Downlink Average Throughput (Mbps)'], errors='coerce')
                summary['Avg DL Throughput(Mbps)'] = throughput.mean()

    except Exception as e:
        logger.error(f"Error getting summary for {kpi_name}: {e}")

    return summary


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python cell_kpi_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]
    results = process_cell_kpis(test_folder)

    print("\n" + "=" * 60)
    print("CELL KPI PROCESSING RESULTS")
    print("=" * 60)

    for sheet_name, df in results.items():
        print(f"\n{sheet_name}:")
        if df is not None and not df.empty:
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
            if 'Date' in df.columns:
                dates = df['Date'].unique()
                print(f"  Date Range: {min(dates)} to {max(dates)}")
            print(f"  Columns: {df.columns.tolist()[:5]}...")
            print(f"  Sample:\n{df.head(2)}")
        else:
            print("  No data")