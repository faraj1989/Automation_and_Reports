#!/usr/bin/env python3
"""
Libyana NPM - Network KPI Processor (NWBH/NW_Daily)
Processes 2G, 3G, 4G whole network KPIs (NWBH and NW_Daily).
Handles duplicate detection to prevent re-inserting existing data.
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

# Network KPI file patterns
NETWORK_KPI_FILES = {
    '2G_NWBH': {
        'patterns': ['* (2G-NWBH).csv', '*(2G-NWBH).csv', '*2G-NWBH*.csv'],
        'sheet_name': '2G_NWBH',
        'key_columns': ['Date', 'Whole Network']  # For duplicate detection
    },
    '2G_NW_Daily': {
        'patterns': ['* (2G-NW_Daily).csv', '*(2G-NW_Daily).csv', '*2G-NW_Daily*.csv'],
        'sheet_name': '2G_NW_Daily',
        'key_columns': ['Date', 'Whole Network']
    },
    '3G_NWBH': {
        'patterns': ['* (3G-NWBH).csv', '*(3G-NWBH).csv', '*3G-NWBH*.csv'],
        'sheet_name': '3G_NWBH',
        'key_columns': ['Date', 'Whole Network']
    },
    '3G_NW_Daily': {
        'patterns': ['* (3G-NW_Daily).csv', '*(3G-NW_Daily).csv', '*3G-NW_Daily*.csv'],
        'sheet_name': '3G_NW_Daily',
        'key_columns': ['Date', 'Whole Network']
    },
    '4G_NWBH': {
        'patterns': ['* (4G-NWBH).csv', '*(4G-NWBH).csv', '*4G-NWBH*.csv'],
        'sheet_name': '4G_NWBH',
        'key_columns': ['Date', 'Whole Network']
    },
    '4G_NW_Daily': {
        'patterns': ['* (4G-NW_Daily).csv', '*(4G-NW_Daily).csv', '*4G-NW_Daily*.csv'],
        'sheet_name': '4G_NW_Daily',
        'key_columns': ['Date', 'Whole Network']
    }
}


def process_network_kpis(day_folder, log_callback=None):
    """
    Process all 6 network KPI files in the day folder.
    Returns a dictionary: sheet_name -> DataFrame
    Handles: Date normalization, empty row removal, duplicate detection, logging
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📊 PROCESSING NETWORK KPIs (NWBH/NW_Daily)")
    log("=" * 60)

    results = {}

    for kpi_name, config in NETWORK_KPI_FILES.items():
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
                    
                    # 1. Remove completely empty rows (all NaN values)
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

                    # 3. Remove duplicate rows within the file itself
                    if 'Date' in df.columns:
                        dup_count = df.duplicated(
                            subset=['Date', 'Whole Network']).sum() if 'Whole Network' in df.columns else 0
                        if dup_count > 0:
                            log(f"   ⚠️ Found {dup_count} duplicate rows in file, removing them")
                            df = df.drop_duplicates(
                                subset=['Date', 'Whole Network']) if 'Whole Network' in df.columns else df

                    results[sheet_name] = df
                    log(f"   ✅ Loaded {len(df)} rows, {len(df.columns)} columns")
                    if 'Date' in df.columns:
                        dates = df['Date'].unique()
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


def process_network_kpis_with_duplicate_check(day_folder, existing_data=None, log_callback=None):
    """
    Process network KPIs with duplicate checking against existing Excel data.
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    # First, process the files
    results = process_network_kpis(day_folder, log_callback)

    if not results:
        return results

    # If existing_data is provided, check for duplicates
    if existing_data:
        log("\n🔄 Checking for duplicates against existing data...")
        for sheet_name, df in results.items():
            if df is None or df.empty:
                continue

            existing_df = existing_data.get(sheet_name)
            if existing_df is not None and not existing_df.empty:
                # Get the key columns for this sheet
                config = next((c for c in NETWORK_KPI_FILES.values() if c['sheet_name'] == sheet_name), None)
                if config and 'key_columns' in config:
                    key_cols = config['key_columns']
                    # Only check if key columns exist in both dataframes
                    existing_cols = set(existing_df.columns)
                    df_cols = set(df.columns)
                    available_keys = [col for col in key_cols if col in existing_cols and col in df_cols]

                    if available_keys:
                        # Create a set of existing keys for quick lookup
                        existing_keys = set()
                        for _, row in existing_df.iterrows():
                            key = tuple(str(row.get(col, '')) for col in available_keys)
                            existing_keys.add(key)

                        # Filter out rows that already exist
                        new_rows = []
                        for _, row in df.iterrows():
                            key = tuple(str(row.get(col, '')) for col in available_keys)
                            if key not in existing_keys:
                                new_rows.append(row)

                        if new_rows:
                            df_new = pd.DataFrame(new_rows)
                            log(f"   {sheet_name}: {len(new_rows)} new rows (skipped {len(df) - len(new_rows)} duplicates)")
                            results[sheet_name] = df_new
                        else:
                            log(f"   {sheet_name}: All {len(df)} rows already exist, skipping")
                            results[sheet_name] = None
                    else:
                        log(f"   {sheet_name}: No key columns found, keeping all {len(df)} rows")
                else:
                    log(f"   {sheet_name}: No key columns configured, keeping all {len(df)} rows")
            else:
                log(f"   {sheet_name}: No existing data, keeping all {len(df)} rows")

    return results


def get_network_kpi_summary(df, kpi_name):
    """
    Extract a summary of key KPIs from a network KPI DataFrame.
    Useful for dashboard display.
    """
    if df is None or df.empty:
        return {}

    summary = {}
    try:
        # Get the most recent row (assuming sorted by Date)
        latest = df.iloc[-1] if 'Date' in df.columns else df.iloc[0]

        if kpi_name == '2G_NWBH' or kpi_name == '2G_NW_Daily':
            summary['Call Setup Success Rate(%)'] = latest.get('Call Setup Success Rate(%)', None)
            summary['TCH Drop Rate(%)'] = latest.get('TCH Drop Rate(%)', None)
            summary['PS Traffic (RLC)(MB)'] = latest.get('PS Traffic (RLC)(MB)', None)
            summary['RR307:TCH Availability(%)'] = latest.get('RR307:TCH Availability(%)', None)
            summary['K3014:Traffic Volume on TCH(Erl)'] = latest.get('K3014:Traffic Volume on TCH(Erl)', None)

        elif kpi_name == '3G_NWBH' or kpi_name == '3G_NW_Daily':
            summary['PS RAB Setup Success Ratio_EFD'] = latest.get('PS RAB Setup Success Ratio_EFD', None)
            summary['PS Drop Ratio(EFD)'] = latest.get('PS Drop Ratio(EFD)', None)
            summary['PS traffic (UL+DL)(GB)'] = latest.get('PS traffic (UL+DL)(GB)', None)

        elif kpi_name == '4G_NWBH' or kpi_name == '4G_NW_Daily':
            summary['RRC Setup Success Rate(%)'] = latest.get('RRC Setup Success Rate(%)', None)
            summary['E-RAB Setup Success Rate'] = latest.get('E-RAB Setup Success Rate', None)
            summary['Service Drop Rate (All)'] = latest.get('Service Drop Rate (All)', None)
            summary['Downlink Traffic Volume(GB)'] = latest.get('Downlink Traffic Volume(GB)', None)
            summary['Uplink Traffic Volume (GB)'] = latest.get('Uplink Traffic Volume (GB)', None)
            summary['VoLTE Traffic Volume (Erl)'] = latest.get('VoLTE Traffic Volume (Erl)', None)

    except Exception as e:
        logger.error(f"Error getting summary for {kpi_name}: {e}")

    return summary


def get_all_network_kpi_columns():
    """
    Return a dictionary of all columns for each network KPI sheet.
    Used for creating Excel sheets.
    """
    return {
        '2G_NWBH': [
            'Date', 'Whole Network', 'Integrity',
            'Downlink TBF Establishment Success Rate(%)',
            'SDCCH Congestion Rate(%)', 'TCH Congestion Rate(%)',
            'TCH Drop Rate(%)', 'SDCCH Drop Rate(%)',
            'RR307:TCH Availability(%)', 'Call Setup Success Rate(%)',
            'Handover Success Rate(%)', 'PS Traffic (RLC)(MB)',
            'RCA313:Assignment Success Rate(%)',
            'Immediate Assignment Success Rate(%)',
            'K3014:Traffic Volume on TCH(Erl)',
            'K3004:Traffic Volume on SDCCH(Erl)',
            'TL9114:Average Throughput of Downlink GPRS RLC(kbit/s)',
            'TL9333:Average Throughput of Downlink EGPRS RLC(kbit/s)',
            'TL9232:Average Throughput of Uplink EGPRS RLC(kbit/s)',
            'TL9014:Average Throughput of Uplink GPRS RLC(kbit/s)',
            'TRX Availability Rate(%)', 'K3015:Available TCHs',
            'K3016:Configured TCHs',
            'Interference Band Proportion (4~5)(%)',
            'Downlink HQI Proportion (0~5)(%)',
            'Uplink HQI Proportion (0~5)(%)',
            'RF Resource Utilization Rate(%)',
            'CM33:Call Drops on Traffic Channel',
            'K3013A:Successful TCH Seizures (Traffic Channel)',
            'CH323:Number of Successful Incoming Internal Inter-Cell Handovers',
            'CH343:Successful Incoming External Inter-Cell Handovers',
            'CH313:Number of Successful Outgoing Internal Inter-Cell Handovers',
            'CH333:Successful Outgoing External Inter-Cell Handovers'
        ],
        '2G_NW_Daily': [
            'Date', 'Whole Network', 'Integrity',
            'Downlink TBF Establishment Success Rate(%)',
            'SDCCH Congestion Rate(%)', 'TCH Congestion Rate(%)',
            'TCH Drop Rate(%)', 'SDCCH Drop Rate(%)',
            'RR307:TCH Availability(%)', 'Call Setup Success Rate(%)',
            'Handover Success Rate(%)', 'PS Traffic (RLC)(MB)',
            'RCA313:Assignment Success Rate(%)',
            'Immediate Assignment Success Rate(%)',
            'K3014:Traffic Volume on TCH(Erl)',
            'K3004:Traffic Volume on SDCCH(Erl)',
            'TL9114:Average Throughput of Downlink GPRS RLC(kbit/s)',
            'TL9333:Average Throughput of Downlink EGPRS RLC(kbit/s)',
            'TL9232:Average Throughput of Uplink EGPRS RLC(kbit/s)',
            'TL9014:Average Throughput of Uplink GPRS RLC(kbit/s)',
            'TRX Availability Rate(%)', 'K3015:Available TCHs',
            'K3016:Configured TCHs',
            'Interference Band Proportion (4~5)(%)',
            'Downlink HQI Proportion (0~5)(%)',
            'Uplink HQI Proportion (0~5)(%)',
            'CM33:Call Drops on Traffic Channel',
            'K3013A:Successful TCH Seizures (Traffic Channel)',
            'CH323:Number of Successful Incoming Internal Inter-Cell Handovers',
            'CH343:Successful Incoming External Inter-Cell Handovers',
            'CH313:Number of Successful Outgoing Internal Inter-Cell Handovers',
            'CH333:Successful Outgoing External Inter-Cell Handovers'
        ],
        '3G_NWBH': [
            'Date', 'Whole Network', 'Integrity',
            'PS RAB Setup Success Ratio_EFD',
            'PS Drop Ratio(EFD)',
            'Inter-RAT Handover Success Rate PS(%)',
            'Availability', 'PS traffic (UL+DL)(GB)',
            'HSDPA Throughput per user (Local Cell)(Kbps)'
        ],
        '3G_NW_Daily': [
            'Date', 'Whole Network', 'Integrity',
            'RRC Setup Success Ratio_EFD',
            'AMR RAB Setup Success Rate(%)',
            'Call Setup Success Rate(%)_EFD',
            'CS Call Drop Rate(%)',
            'Soft Handover Success Rate(%)',
            'Inter-RAT Handover Success Rate CS(%)',
            'Availability',
            'VS.HHO.InterFreqOut.Success Ratio',
            'IntraFreq HHO SR',
            'HSDPA Throughput per user (Local Cell)(Kbps)',
            'PS RAB Setup Success Ratio_EFD',
            'PS Drop Ratio(EFD)',
            'Inter-RAT Handover Success Rate PS(%)',
            'PS traffic (UL+DL)(GB)'
        ],
        '4G_NWBH': [
            'Date', 'Whole Network', 'Integrity',
            'RRC Setup Success Rate(%)',
            'E-RAB Setup Success Rate',
            'Intra-HO execute Success Rate',
            'Service Drop Rate (All)',
            'Downlink Packet Loss',
            'Uplink Packet Loss Rate',
            'User Downlink Average Throughput (Mbps)',
            'User Uplink Average Throughput (Mbps)',
            'DL PRB Utilizing Rate(%)',
            'UL PRB Utilizing Rate(%)',
            'CSFB Preparation Success Rate',
            'L.UL.Interference.Avg(dBm)',
            'Radio Network Availability Rate(%)',
            'L.Cell.Unavail.Dur.Sys(s)',
            'L.Cell.Unavail.Dur.Manual(s)',
            'L.Traffic.User.Avg',
            'L.Traffic.User.Max',
            'Downlink Traffic Volume(GB)',
            'Uplink Traffic Volume (GB)',
            'Inter-Freq Handover Success Rate (%)(%)',
            'S1-Signal Connection Establishment Success Rate (%)(%)',
            'VoLTE Traffic Volume (Erl)',
            'VoLTE Setup Success Rate-ZM(%)'
        ],
        '4G_NW_Daily': [
            'Date', 'Whole Network', 'Integrity',
            'UL Traffic Volume(GB)',
            'DL Traffic Volume(GB)',
            'Radio Network Availability Rate(%)',
            'L.Cell.Unavail.Dur.Sys(s)',
            'L.Cell.Unavail.Dur.Manual(s)',
            'L.Traffic.User.Avg',
            'L.Traffic.User.Max',
            'DL PRB Utilizing Rate(%)',
            'UL PRB Utilizing Rate(%)',
            'VoLTE Traffic Volume (Erl)',
            'VoLTE Setup Success Rate-ZM(%)'
        ]
    }


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python network_kpi_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]
    results = process_network_kpis(test_folder)

    print("\n" + "=" * 60)
    print("NETWORK KPI PROCESSING RESULTS")
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