#!/usr/bin/env python3
"""
Libyana NPM - Site Detail Processor
Generates per-site detailed view from 2G, 3G, 4G data.
"""

import os
import re
import glob
import logging
from datetime import datetime
import pandas as pd
from backend.csv_loader import read_csv_skip_metadata
from backend.site_processor import find_file, extract_physical_name

logger = logging.getLogger(__name__)

# Huawei LTE channel bandwidth encoding (Downlink bandwidth = PRB count) ->
# actual channel bandwidth in MHz, per 3GPP TS 36.101 Table 5.6-1.
LTE_BANDWIDTH_MAP = {
    'CELL_BW_N6': '1.4MHz',
    'CELL_BW_N15': '3MHz',
    'CELL_BW_N25': '5MHz',
    'CELL_BW_N50': '10MHz',
    'CELL_BW_N75': '15MHz',
    'CELL_BW_N100': '20MHz',
}

# Scenario mapping based on cell count
SCENARIO_MAP = {
    24: 'MM',
    12: '4T6S',
    6: '2T3S(2T2R)',
    2: '2T3S(2T2R)',
    3: '2T3S(2T2R)',
    18: '4T3S(4T4R)',
    25: '2T3S(2T2R)',
}

# Special case overrides
SPECIAL_SCENARIO = {
    'S_Lampsite_1': '2T3S(2T2R)',
    'S-CATWALK-1': '4T3S(4T4R)',
}

SITE_DETAIL_HEADER = [
    'Site Name',
    '2G GSM900 Band',
    '2G DCS1800 Band',
    '3G U2100 Band',
    '3G U900 Band',
    '4G L1800 F1 Band',
    '4G L1800 F2 Band',
    '4G L2100 Band',
    '4G L900 Band',
    '4G L700 Band',
    'Scenario',
    'RAT',
    'Sectors Number',
    'Current RAT'
]


def get_4g_scenario(df_4g, site_name):
    """
    Determine the scenario for a site based on cell count and special cases.
    """
    if site_name in SPECIAL_SCENARIO:
        return SPECIAL_SCENARIO[site_name]

    if df_4g is None or df_4g.empty:
        return 'No LTE'

    # Count cells for this site
    site_data = df_4g[df_4g['eNodeB Name'] == site_name]
    cell_count = site_data['Cell Name'].nunique()

    if cell_count == 0:
        return 'No LTE'

    # Exact match
    if cell_count in SCENARIO_MAP:
        return SCENARIO_MAP[cell_count]

    # For 24 cells (MM) - common case
    if cell_count >= 20:
        return 'MM'
    elif cell_count >= 10:
        return '4T6S'
    elif cell_count >= 5:
        return '2T3S(2T2R)'
    elif cell_count >= 2:
        return '2T3S(2T2R)'
    else:
        return '2T3S(2T2R)'


def get_rat(has_2g, has_3g, has_4g):
    """Determine RAT based on technology presence."""
    if has_2g and has_3g and has_4g:
        return 'GUL'
    elif has_2g and has_3g:
        return 'GU'
    elif has_2g and has_4g:
        return 'GL'
    elif has_3g and has_4g:
        return 'UL'
    elif has_2g:
        return 'G'
    elif has_3g:
        return 'U'
    elif has_4g:
        return 'L'
    else:
        return ''


def _lte_band_label(band_name, cell_rows):
    """'L1800' -> 'L1800 (20MHz)' using the decoded Downlink bandwidth of the
    matching cells; if a band's cells carry more than one bandwidth at the
    same site (unusual but possible), all distinct values are shown."""
    if cell_rows.empty:
        return ''
    if 'Downlink bandwidth' not in cell_rows.columns:
        return band_name
    raw_values = cell_rows['Downlink bandwidth'].dropna().astype(str).str.strip().unique()
    decoded = sorted({LTE_BANDWIDTH_MAP[v] for v in raw_values if v in LTE_BANDWIDTH_MAP})
    return f"{band_name} ({'/'.join(decoded)})" if decoded else band_name


def build_current_rat(row):
    """Build the Current RAT string from band presence."""
    band_order = [
        '2G GSM900 Band',
        '2G DCS1800 Band',
        '3G U2100 Band',
        '3G U900 Band',
        '4G L1800 F1 Band',
        '4G L1800 F2 Band',
        '4G L2100 Band',
        '4G L900 Band',
        '4G L700 Band'
    ]

    bands = []
    for col in band_order:
        val = row.get(col, '')
        if val and str(val).strip():
            bands.append(str(val).strip())
        else:
            bands.append('')

    # Join with + and add trailing + for consistency
    return '+'.join(bands) + '+'


def generate_site_detail(day_folder, log_callback=None):
    """
    Generate per-site detail table from the day folder.
    Uses eNodeB Name from 4G as the PRIMARY site name (preserves FN/FTTS suffix).
    For sites without 4G, merges 2G and 3G into a single entry using 2G name.
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📊 GENERATING SITE DETAIL TABLE")
    log("=" * 60)

    # Find the 3 CSV files
    patterns_2g = ['* (2G).csv', '*(2G).csv', '*2G*.csv']
    patterns_3g = ['* (3G).csv', '*(3G).csv', '*3G*.csv']
    patterns_4g = ['* (4G).csv', '*(4G).csv', '*4G*.csv']

    file_2g = find_file(day_folder, patterns_2g)
    file_3g = find_file(day_folder, patterns_3g)
    file_4g = find_file(day_folder, patterns_4g)

    if not file_2g and not file_3g and not file_4g:
        log("❌ No 2G, 3G, or 4G files found")
        return None

    # Read the CSVs
    df_2g = read_csv_skip_metadata(file_2g) if file_2g else None
    df_3g = read_csv_skip_metadata(file_3g) if file_3g else None
    df_4g = read_csv_skip_metadata(file_4g) if file_4g else None

    log(f"2G: {len(df_2g) if df_2g is not None else 0} rows")
    log(f"3G: {len(df_3g) if df_3g is not None else 0} rows")
    log(f"4G: {len(df_4g) if df_4g is not None else 0} rows")

    # Build a mapping: physical_name -> 4G site name (with suffix)
    physical_to_4g = {}

    if df_4g is not None and not df_4g.empty:
        for site in df_4g['eNodeB Name'].unique():
            physical = extract_physical_name(site, '4g')
            physical_to_4g[physical] = site
            log(f"  4G: {site} -> physical: {physical}")

    # Get all 4G physical names for quick lookup
    physical_names = set(physical_to_4g.keys())

    # Collect all sites that need to be processed
    # We start with all 4G sites (primary)
    site_entries = {physical_to_4g[phys]: {'source': '4g'} for phys in physical_to_4g}

    # Add 2G sites that don't have a matching 4G
    if df_2g is not None and not df_2g.empty:
        for site in df_2g['Site Name'].unique():
            physical = extract_physical_name(site, '2g')
            if physical not in physical_names:
                # No 4G site matches, check if already in site_entries (from 3G)
                if site in site_entries:
                    # Update existing entry to include 2G
                    site_entries[site]['source'] = '2g+3g' if site_entries[site]['source'] == '3g' else '2g'
                else:
                    # Check if there's a 3G entry for this physical name
                    found_3g = False
                    for existing_site in list(site_entries.keys()):
                        existing_physical = extract_physical_name(existing_site, '3g')
                        if existing_physical == physical:
                            # Found matching 3G site, update its entry
                            site_entries[existing_site]['source'] = '2g+3g'
                            found_3g = True
                            log(f"  2G: {site} -> merges with 3G: {existing_site}")
                            break

                    if not found_3g:
                        # No matching 3G or 4G, add as 2G-only
                        site_entries[site] = {'source': '2g'}
                        log(f"  2G-only: {site} -> physical: {physical}")
            else:
                # Find the matching 4G site name
                matching_4g = physical_to_4g.get(physical)
                if matching_4g:
                    log(f"  2G: {site} -> matches 4G: {matching_4g}")

    # Add 3G sites that don't have a matching 4G
    if df_3g is not None and not df_3g.empty:
        for site in df_3g['NodeB Name'].unique():
            physical = extract_physical_name(site, '3g')
            if physical not in physical_names:
                # No 4G site matches, check if already in site_entries (from 2G)
                # Check if there's a 2G entry for this physical name
                found_2g = False
                for existing_site in list(site_entries.keys()):
                    existing_physical = extract_physical_name(existing_site, '2g')
                    if existing_physical == physical:
                        # Found matching 2G site, update its entry
                        site_entries[existing_site]['source'] = '2g+3g'
                        found_2g = True
                        log(f"  3G: {site} -> merges with 2G: {existing_site}")
                        break

                if not found_2g:
                    # No matching 2G, add as 3G-only
                    site_entries[site] = {'source': '3g'}
                    log(f"  3G-only: {site} -> physical: {physical}")
            else:
                matching_4g = physical_to_4g.get(physical)
                if matching_4g:
                    log(f"  3G: {site} -> matches 4G: {matching_4g}")

    log(f"Found {len(site_entries)} unique site entries to process")

    # Build the detail table
    rows = []

    for site_name, info in site_entries.items():
        row = {'Site Name': site_name}

        # Determine physical name for matching
        if info['source'] == '4g':
            physical_name = extract_physical_name(site_name, '4g')
        elif info['source'] == '2g':
            physical_name = extract_physical_name(site_name, '2g')
        elif info['source'] == '3g':
            physical_name = extract_physical_name(site_name, '3g')
        else:  # 2g+3g
            # Use the site_name (which is from 2G)
            physical_name = extract_physical_name(site_name, '2g')

        # Track technologies present
        has_2g = False
        has_3g = False
        has_4g = (info['source'] == '4g')

        # Track sectors
        sectors = set()

        # ---- 2G Processing ----
        if df_2g is not None and not df_2g.empty:
            # Find matching 2G site
            matching_2g = None
            for s in df_2g['Site Name'].unique():
                if extract_physical_name(s, '2g') == physical_name:
                    matching_2g = s
                    break

            if matching_2g:
                site_data = df_2g[df_2g['Site Name'] == matching_2g]
                freqs = site_data['DL frequency'].unique()
                has_2g = True

                row['2G GSM900 Band'] = 'GSM900' if any('900' in str(f).upper() for f in freqs) else ''
                row['2G DCS1800 Band'] = 'DCS1800' if any('1800' in str(f).upper() for f in freqs) else ''

                sectors.update(site_data['Cell Name'].unique())
            else:
                row['2G GSM900 Band'] = ''
                row['2G DCS1800 Band'] = ''
        else:
            row['2G GSM900 Band'] = ''
            row['2G DCS1800 Band'] = ''

        # ---- 3G Processing ----
        if df_3g is not None and not df_3g.empty:
            # Find matching 3G site
            matching_3g = None
            for s in df_3g['NodeB Name'].unique():
                if extract_physical_name(s, '3g') == physical_name:
                    matching_3g = s
                    break

            if matching_3g:
                site_data = df_3g[df_3g['NodeB Name'] == matching_3g]
                bands = site_data['Band Indicator'].unique()
                has_3g = True

                row['3G U2100 Band'] = 'U2100' if any('BAND1' in str(b).upper() for b in bands) else ''
                row['3G U900 Band'] = 'U900' if any('BAND8' in str(b).upper() for b in bands) else ''

                sectors.update(site_data['Cell Name'].unique())
            else:
                row['3G U2100 Band'] = ''
                row['3G U900 Band'] = ''
        else:
            row['3G U2100 Band'] = ''
            row['3G U900 Band'] = ''

        # ---- 4G Processing ----
        if has_4g and df_4g is not None and not df_4g.empty:
            # Use the site_name directly (it's already the 4G name)
            site_data = df_4g[df_4g['eNodeB Name'] == site_name]

            if not site_data.empty:
                site_data = site_data.copy()
                site_data['_earfcn'] = pd.to_numeric(site_data['Downlink EARFCN'], errors='coerce')
                site_data['_band'] = pd.to_numeric(site_data['Frequency band'], errors='coerce')

                # L1800 F1 vs F2 are the same "Band" value, distinguished only by EARFCN
                f1_rows = site_data[site_data['_earfcn'].isin([1401, 1875])]
                f2_rows = site_data[site_data['_earfcn'].isin([1250, 1257])]
                l2100_rows = site_data[site_data['_band'] == 1]
                l900_rows = site_data[site_data['_band'] == 8]
                l700_rows = site_data[site_data['_band'] == 28]

                row['4G L1800 F1 Band'] = _lte_band_label('L1800', f1_rows)
                row['4G L1800 F2 Band'] = _lte_band_label('L1800', f2_rows)
                row['4G L2100 Band'] = _lte_band_label('L2100', l2100_rows)
                row['4G L900 Band'] = _lte_band_label('L900', l900_rows)
                row['4G L700 Band'] = _lte_band_label('L700', l700_rows)

                sectors.update(site_data['Cell Name'].unique())
            else:
                row['4G L1800 F1 Band'] = ''
                row['4G L1800 F2 Band'] = ''
                row['4G L2100 Band'] = ''
                row['4G L900 Band'] = ''
                row['4G L700 Band'] = ''
        else:
            row['4G L1800 F1 Band'] = ''
            row['4G L1800 F2 Band'] = ''
            row['4G L2100 Band'] = ''
            row['4G L900 Band'] = ''
            row['4G L700 Band'] = ''

        # ---- Determine RAT ----
        row['RAT'] = get_rat(has_2g, has_3g, has_4g)

        # ---- Determine Scenario ----
        if has_4g and df_4g is not None:
            matching_4g = site_name
            if matching_4g:
                row['Scenario'] = get_4g_scenario(df_4g, matching_4g)
            else:
                row['Scenario'] = 'No LTE'
        else:
            row['Scenario'] = 'No LTE'

        # ---- Sectors Number ----
        row['Sectors Number'] = len(sectors)

        # ---- Current RAT ----
        row['Current RAT'] = build_current_rat(row)

        rows.append(row)

    # Convert to DataFrame
    df_result = pd.DataFrame(rows, columns=SITE_DETAIL_HEADER)

    log(f"✅ Generated detail table with {len(df_result)} sites")
    log("=" * 60)

    return df_result
def get_latest_available_day(local_root, log_callback=None):
    """
    Find the latest available day with actual data (not folder name).
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    if not os.path.exists(local_root):
        log(f"❌ Local root does not exist: {local_root}")
        return None

    # Get all folders sorted by date (newest first)
    folders = []
    for item in os.listdir(local_root):
        item_path = os.path.join(local_root, item)
        if os.path.isdir(item_path):
            try:
                datetime.strptime(item, '%Y-%m-%d')
                unzipped_path = os.path.join(item_path, 'unzipped')
                if os.path.exists(unzipped_path):
                    csv_count = len(glob.glob(os.path.join(unzipped_path, '*.csv')))
                    if csv_count > 0:
                        folders.append((item, unzipped_path))
            except ValueError:
                pass

    if not folders:
        log("❌ No folders with CSV data found")
        return None

    folders.sort(reverse=True)  # Newest first

    # Return the latest folder path
    latest_folder = folders[0][1]
    log(f"📁 Latest available day: {folders[0][0]} -> {latest_folder}")

    return latest_folder