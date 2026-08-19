#!/usr/bin/env python3
"""
Libyana NPM - Site Summary Processor
Processes 2G, 3G, and 4G CSVs to count sites and band combinations.
"""

import os
import re
import glob
import logging
from datetime import datetime
import pandas as pd
from backend.csv_loader import read_csv_skip_metadata

logger = logging.getLogger(__name__)

# 4G L1800 EARFCN mapping
L1800_F1_EARFCNS = [1401, 1875]
L1800_F2_EARFCNS = [1250, 1257]

# Updated header: Total Physical Sites comes right after day
SITE_SUMMARY_HEADER = [
    'Day-',
    'Total Physical Sites (2G+3G+4G)',
    '2G physical sites',
    '2G GSM900 Band',
    '2G DCS1800 Band',
    '2G 900 only',
    '2G 1800 only',
    '3G physical sites',
    '3G U2100 Band',
    '3G U900 Band',
    '3G U2100 only',
    '3G U900 only',
    '4G physical sites',
    '4G L1800 F1 Band',
    '4G L1800 F2 Band',
    '4G L2100 Band',
    '4G L900 Band',
    '4G L700 Band',
    '4G L1800 only',
    '4G L2100 only',
    '4G L900 only',
    '4G L700 only',
    '2G only sites',
    '3G only sites',
    '4G only sites',
    '2G+3G sites',
    '2G+4G sites',
    '3G+4G sites',
    '2G+3G+4G sites'
]


def extract_physical_name(name, tech):
    """
    Extract the physical site name from technology-specific naming.

    2G: 'BGZ001' -> 'BGZ001' (no change)
    3G: 'UBGZ001' -> 'BGZ001' (remove leading 'U')
    4G: 'BGZ001(FN)', 'BYDA023(FTTS)', 'BYDA019' -> 'BGZ001', 'BYDA023', 'BYDA019'
    """
    if tech == '2g':
        return name.strip()
    elif tech == '3g':
        # Remove leading 'U' if present
        if name.startswith('U'):
            return name[1:]
        return name
    elif tech == '4g':
        # Remove suffixes like (FN), (FTTS), (FTT), etc.
        # Also remove trailing spaces
        clean = re.sub(r'\s*\([^)]*\)$', '', name)
        return clean.strip()
    return name


def find_file(folder, patterns):
    """Search for a file matching any of the given patterns."""
    for pattern in patterns:
        full_pattern = os.path.join(folder, pattern)
        matches = glob.glob(full_pattern)
        if matches:
            return matches[0]
    return None


def process_2g(df):
    """Return dict with 2G counts."""
    if df is None or df.empty:
        return {
            '2G physical sites': 0,
            '2G GSM900 Band': 0,
            '2G DCS1800 Band': 0,
            '2G 900 only': 0,
            '2G 1800 only': 0
        }

    site_col = 'Site Name'
    freq_col = 'DL frequency'

    unique_sites = df[site_col].unique()
    total = len(unique_sites)

    count_900 = 0
    count_1800 = 0
    count_900_only = 0
    count_1800_only = 0

    for site in unique_sites:
        freqs = df[df[site_col] == site][freq_col].unique()
        has_900 = any('900' in str(f).upper() for f in freqs)
        has_1800 = any('1800' in str(f).upper() for f in freqs)

        if has_900:
            count_900 += 1
        if has_1800:
            count_1800 += 1

        if has_900 and not has_1800:
            count_900_only += 1
        elif has_1800 and not has_900:
            count_1800_only += 1

    return {
        '2G physical sites': total,
        '2G GSM900 Band': count_900,
        '2G DCS1800 Band': count_1800,
        '2G 900 only': count_900_only,
        '2G 1800 only': count_1800_only
    }


def process_3g(df):
    """Return dict with 3G counts."""
    if df is None or df.empty:
        return {
            '3G physical sites': 0,
            '3G U2100 Band': 0,
            '3G U900 Band': 0,
            '3G U2100 only': 0,
            '3G U900 only': 0
        }

    nodeb_col = 'NodeB Name'
    band_col = 'Band Indicator'

    unique_nodebs = df[nodeb_col].unique()
    total = len(unique_nodebs)

    count_u2100 = 0
    count_u900 = 0
    count_u2100_only = 0
    count_u900_only = 0

    for nodeb in unique_nodebs:
        bands = df[df[nodeb_col] == nodeb][band_col].unique()
        has_u2100 = any('BAND1' in str(b).upper() for b in bands)
        has_u900 = any('BAND8' in str(b).upper() for b in bands)

        if has_u2100:
            count_u2100 += 1
        if has_u900:
            count_u900 += 1

        if has_u2100 and not has_u900:
            count_u2100_only += 1
        elif has_u900 and not has_u2100:
            count_u900_only += 1

    return {
        '3G physical sites': total,
        '3G U2100 Band': count_u2100,
        '3G U900 Band': count_u900,
        '3G U2100 only': count_u2100_only,
        '3G U900 only': count_u900_only
    }


def process_4g(df):
    """Return dict with 4G counts using EARFCN for F1/F2 distinction."""
    if df is None or df.empty:
        return {
            '4G physical sites': 0,
            '4G L1800 F1 Band': 0,
            '4G L1800 F2 Band': 0,
            '4G L2100 Band': 0,
            '4G L900 Band': 0,
            '4G L700 Band': 0,
            '4G L1800 only': 0,
            '4G L2100 only': 0,
            '4G L900 only': 0,
            '4G L700 only': 0
        }

    enb_col = 'eNodeB Name'

    # Find the band column (flexible name matching)
    band_col = None
    for col in df.columns:
        col_clean = col.strip()
        if col_clean == 'Frequency band' or col_clean == 'Band':
            band_col = col
            break

    # Find EARFCN column
    earfcn_col = None
    for col in df.columns:
        if 'EARFCN' in col:
            earfcn_col = col
            break

    if enb_col not in df.columns or band_col is None:
        logger.warning(f"4G columns not found. Available: {df.columns.tolist()}")
        return {
            '4G physical sites': 0,
            '4G L1800 F1 Band': 0,
            '4G L1800 F2 Band': 0,
            '4G L2100 Band': 0,
            '4G L900 Band': 0,
            '4G L700 Band': 0,
            '4G L1800 only': 0,
            '4G L2100 only': 0,
            '4G L900 only': 0,
            '4G L700 only': 0
        }

    unique_enbs = df[enb_col].unique()
    total = len(unique_enbs)

    count_l1800_f1 = 0
    count_l1800_f2 = 0
    count_l2100 = 0
    count_l900 = 0
    count_l700 = 0
    count_l1800_only = 0
    count_l2100_only = 0
    count_l900_only = 0
    count_l700_only = 0

    for enb in unique_enbs:
        enb_data = df[df[enb_col] == enb]

        # Get bands - handle both string and numeric
        bands = []
        for b in enb_data[band_col]:
            try:
                if b is not None:
                    bands.append(int(float(b)))
            except (ValueError, TypeError):
                pass

        band_set = set(bands)

        has_l1800 = 3 in band_set
        has_l2100 = 1 in band_set
        has_l900 = 8 in band_set
        has_l700 = 28 in band_set

        # EARFCN detection for L1800 F1/F2
        has_l1800_f1 = False
        has_l1800_f2 = False
        if has_l1800 and earfcn_col is not None and earfcn_col in enb_data.columns:
            for e in enb_data[earfcn_col]:
                if e is not None:
                    try:
                        earfcn = int(float(e))
                        if earfcn in L1800_F1_EARFCNS:
                            has_l1800_f1 = True
                        elif earfcn in L1800_F2_EARFCNS:
                            has_l1800_f2 = True
                    except (ValueError, TypeError):
                        pass

        if has_l1800 and has_l1800_f1:
            count_l1800_f1 += 1
        if has_l1800 and has_l1800_f2:
            count_l1800_f2 += 1

        if has_l2100:
            count_l2100 += 1
        if has_l900:
            count_l900 += 1
        if has_l700:
            count_l700 += 1

        # Count single-band only sites
        if len(band_set) == 1:
            if has_l1800:
                count_l1800_only += 1
            elif has_l2100:
                count_l2100_only += 1
            elif has_l900:
                count_l900_only += 1
            elif has_l700:
                count_l700_only += 1

    return {
        '4G physical sites': total,
        '4G L1800 F1 Band': count_l1800_f1,
        '4G L1800 F2 Band': count_l1800_f2,
        '4G L2100 Band': count_l2100,
        '4G L900 Band': count_l900,
        '4G L700 Band': count_l700,
        '4G L1800 only': count_l1800_only,
        '4G L2100 only': count_l2100_only,
        '4G L900 only': count_l900_only,
        '4G L700 only': count_l700_only
    }

def get_physical_sites(df_2g, df_3g, df_4g):
    """
    Extract physical site names from ALL technologies and calculate overlaps.
    This ensures ABDN001 and ABDN001(FTTS) are correctly identified as one site.
    """
    physical_sites = {}

    # Extract 2G physical sites
    if df_2g is not None and not df_2g.empty:
        for name in df_2g['Site Name'].unique():
            physical = extract_physical_name(name, '2g')
            physical_sites[physical] = physical_sites.get(physical, set())
            physical_sites[physical].add('2g')
            logger.debug(f"2G: {name} -> {physical}")

    # Extract 3G physical sites
    if df_3g is not None and not df_3g.empty:
        for name in df_3g['NodeB Name'].unique():
            physical = extract_physical_name(name, '3g')
            physical_sites[physical] = physical_sites.get(physical, set())
            physical_sites[physical].add('3g')
            logger.debug(f"3G: {name} -> {physical}")

    # Extract 4G physical sites
    if df_4g is not None and not df_4g.empty:
        for name in df_4g['eNodeB Name'].unique():
            physical = extract_physical_name(name, '4g')
            physical_sites[physical] = physical_sites.get(physical, set())
            physical_sites[physical].add('4g')
            logger.debug(f"4G: {name} -> {physical}")

    # Calculate counts
    total_physical = len(physical_sites)

    count_2g_only = 0
    count_3g_only = 0
    count_4g_only = 0
    count_2g_3g = 0
    count_2g_4g = 0
    count_3g_4g = 0
    count_all_three = 0

    for site, techs in physical_sites.items():
        if techs == {'2g'}:
            count_2g_only += 1
        elif techs == {'3g'}:
            count_3g_only += 1
        elif techs == {'4g'}:
            count_4g_only += 1
        elif techs == {'2g', '3g'}:
            count_2g_3g += 1
        elif techs == {'2g', '4g'}:
            count_2g_4g += 1
        elif techs == {'3g', '4g'}:
            count_3g_4g += 1
        elif techs == {'2g', '3g', '4g'}:
            count_all_three += 1

    return {
        'Total Physical Sites (2G+3G+4G)': total_physical,
        '2G only sites': count_2g_only,
        '3G only sites': count_3g_only,
        '4G only sites': count_4g_only,
        '2G+3G sites': count_2g_3g,
        '2G+4G sites': count_2g_4g,
        '3G+4G sites': count_3g_4g,
        '2G+3G+4G sites': count_all_three
    }


def process_site_day(day_folder, log_callback=None):
    """
    Process all three CSVs and return a summary dictionary.
    The 'Total Physical Sites' column is placed right after 'day'.
    """

    def log(msg):
        if log_callback:
            log_callback(msg)

    patterns_2g = ['* (2G).csv', '*(2G).csv', '*2G*.csv']
    patterns_3g = ['* (3G).csv', '*(3G).csv', '*3G*.csv']
    patterns_4g = ['* (4G).csv', '*(4G).csv', '*4G*.csv']

    file_2g = find_file(day_folder, patterns_2g)
    file_3g = find_file(day_folder, patterns_3g)
    file_4g = find_file(day_folder, patterns_4g)

    site_row = {}
    df_2g = None
    df_3g = None
    df_4g = None

    if file_2g:
        df_2g = read_csv_skip_metadata(file_2g)
        if df_2g is not None and df_2g.empty:
            log(f"⚠️ 2G file is EMPTY: {os.path.basename(file_2g)}")
            df_2g = None
        site_row.update(process_2g(df_2g))
    else:
        site_row.update(process_2g(None))

    if file_3g:
        df_3g = read_csv_skip_metadata(file_3g)
        if df_3g is not None and df_3g.empty:
            log(f"⚠️ 3G file is EMPTY: {os.path.basename(file_3g)}")
            df_3g = None
        site_row.update(process_3g(df_3g))
    else:
        site_row.update(process_3g(None))

    if file_4g:
        df_4g = read_csv_skip_metadata(file_4g)
        if df_4g is not None and df_4g.empty:
            log(f"⚠️ 4G file is EMPTY: {os.path.basename(file_4g)}")
            df_4g = None
        site_row.update(process_4g(df_4g))
    else:
        site_row.update(process_4g(None))

    # Get physical site overlaps (this now correctly handles ABDN001 vs ABDN001(FTTS))
    physical_counts = get_physical_sites(df_2g, df_3g, df_4g)
    site_row.update(physical_counts)

    # Ensure all columns exist (fill missing with 0)
    for col in SITE_SUMMARY_HEADER:
        if col not in site_row:
            site_row[col] = 0

    return site_row


def process_all_days(local_root, log_callback=None):
    """Process all days from all unzipped folders."""

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📊 PROCESSING ALL DAYS")
    log("=" * 60)

    if not os.path.exists(local_root):
        log(f"❌ Local root does not exist: {local_root}")
        return {}

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
                        folders.append(item)
            except ValueError:
                pass

    if not folders:
        log("❌ No folders with CSV data found")
        return {}

    folders.sort(reverse=True)
    log(f"📁 Found {len(folders)} folders: {folders}")

    all_results = {}
    processed_count = 0
    empty_count = 0

    for folder in folders:
        unzipped_path = os.path.join(local_root, folder, 'unzipped')
        log(f"\n📂 Processing folder: {folder}")

        patterns_2g = ['* (2G).csv', '*(2G).csv', '*2G*.csv']
        patterns_3g = ['* (3G).csv', '*(3G).csv', '*3G*.csv']
        patterns_4g = ['* (4G).csv', '*(4G).csv', '*4G*.csv']

        file_2g = find_file(unzipped_path, patterns_2g)
        file_3g = find_file(unzipped_path, patterns_3g)
        file_4g = find_file(unzipped_path, patterns_4g)

        df_2g = None
        df_3g = None
        df_4g = None

        if file_2g:
            df_2g = read_csv_skip_metadata(file_2g)
            if df_2g is not None and df_2g.empty:
                log(f"   ⚠️ 2G file is EMPTY (0 rows)")
                df_2g = None
            elif df_2g is not None:
                log(f"   ✅ 2G loaded: {len(df_2g)} rows")
        else:
            log(f"   ⚠️ No 2G file found")

        if file_3g:
            df_3g = read_csv_skip_metadata(file_3g)
            if df_3g is not None and df_3g.empty:
                log(f"   ⚠️ 3G file is EMPTY (0 rows)")
                df_3g = None
            elif df_3g is not None:
                log(f"   ✅ 3G loaded: {len(df_3g)} rows")
        else:
            log(f"   ⚠️ No 3G file found")

        if file_4g:
            df_4g = read_csv_skip_metadata(file_4g)
            if df_4g is not None and df_4g.empty:
                log(f"   ⚠️ 4G file is EMPTY (0 rows)")
                df_4g = None
            elif df_4g is not None:
                log(f"   ✅ 4G loaded: {len(df_4g)} rows")
        else:
            log(f"   ⚠️ No 4G file found")

        def safe_extract_dates(df, date_column='Date'):
            if df is None or df.empty or date_column not in df.columns:
                return set()
            dates = set()
            for val in df[date_column]:
                try:
                    if pd.isna(val):
                        continue
                    date_str = str(val).strip()
                    if not date_str:
                        continue
                    dt = pd.to_datetime(date_str, errors='coerce')
                    if pd.isna(dt):
                        continue
                    dates.add(dt.strftime('%Y-%m-%d'))
                except (ValueError, TypeError, AttributeError):
                    continue
            return dates

        dates_2g = safe_extract_dates(df_2g, 'Date')
        dates_3g = safe_extract_dates(df_3g, 'Date')
        dates_4g = safe_extract_dates(df_4g, 'Date')

        all_dates = dates_2g.union(dates_3g).union(dates_4g)

        if not all_dates:
            log(f"   ❌ No dates found in any file")
            continue

        all_dates = sorted(list(all_dates), reverse=True)
        log(f"   📅 Found {len(all_dates)} unique dates")

        def filter_by_date(df, date_str, date_column='Date'):
            if df is None or df.empty or date_column not in df.columns:
                return None
            try:
                mask = pd.to_datetime(df[date_column], errors='coerce').dt.strftime('%Y-%m-%d') == date_str
                filtered = df[mask]
                return filtered if not filtered.empty else None
            except (ValueError, TypeError, AttributeError):
                return None

        for date_str in all_dates:
            df_2g_date = filter_by_date(df_2g, date_str, 'Date')
            df_3g_date = filter_by_date(df_3g, date_str, 'Date')
            df_4g_date = filter_by_date(df_4g, date_str, 'Date')

            if df_2g_date is None and df_3g_date is None and df_4g_date is None:
                empty_count += 1
                continue

            result = {}
            result.update(process_2g(df_2g_date))
            result.update(process_3g(df_3g_date))
            result.update(process_4g(df_4g_date))

            physical_counts = get_physical_sites(df_2g_date, df_3g_date, df_4g_date)
            result.update(physical_counts)
            result['day'] = date_str

            for col in SITE_SUMMARY_HEADER:
                if col not in result:
                    result[col] = 0

            all_results[date_str] = result
            processed_count += 1
            log(f"   ✅ Processed {date_str}: Total={result.get('Total Physical Sites (2G+3G+4G)', 0)}, 2G={result.get('2G physical sites', 0)}, 3G={result.get('3G physical sites', 0)}, 4G={result.get('4G physical sites', 0)}")

    log(f"\n" + "=" * 60)
    log(f"✅ PROCESS ALL DAYS COMPLETE")
    log(f"   Processed: {processed_count} days")
    if empty_count > 0:
        log(f"   Skipped (empty): {empty_count} days")
    log(f"=" * 60)

    return all_results


def get_latest_day_folder(local_root):
    """Find the latest day folder with unzipped data."""
    if not os.path.exists(local_root):
        return None

    folders = []
    for item in os.listdir(local_root):
        item_path = os.path.join(local_root, item)
        if os.path.isdir(item_path):
            try:
                datetime.strptime(item, '%Y-%m-%d')
                unzipped = os.path.join(item_path, 'unzipped')
                if os.path.exists(unzipped) and glob.glob(os.path.join(unzipped, '*.csv')):
                    folders.append(item)
            except ValueError:
                pass

    if not folders:
        return None

    folders.sort(reverse=True)
    latest = folders[0]
    return os.path.join(local_root, latest, 'unzipped')