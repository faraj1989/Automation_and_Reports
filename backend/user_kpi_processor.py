#!usrbinenv python3.txt
# !/usr/bin/env python3
"""
Libyana NPM - User KPI Processor
Processes 5 user CSV files with MAX aggregation per day.
PS Roaming: 2G, 3G, 4G kept as separate columns (NOT summed).
Other KPI types (Network, Cell, Traffic) are NOT affected.
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

# User KPI file patterns
USER_KPI_FILES = {
    'CS_Roaming': {
        'patterns': [
            '* (CS Roaming users).csv',
            '*(CS Roaming users).csv',
            '*CS Roaming*.csv',
            '*CS Roaming*.csv'
        ],
        'sheet_name': 'User_CS_Roaming',
        'filter': {'column': 'index', 'value': '218919121253', 'contains': True},
        'aggregate': {
            'groupby': 'Time',  # Group by hour first
            'sum_columns': ['Number of Registered Subscribers(entries)'],
            'then_max': True  # Then take MAX per day
        }
    },
    'CS_Subscribers': {
        'patterns': [
            '* (MSC Server KPI-CS Subscribers+total).csv',
            '*(MSC Server KPI-CS Subscribers+total).csv',
            '*MSC Server*.csv',
            '*MSC Server*.csv'
        ],
        'sheet_name': 'User_CS_Subscribers'
    },
    'PS_Roaming': {
        'patterns': [
            '* (PS Roaming users).csv',
            '*(PS Roaming users).csv',
            '*PS Roaming*.csv',
            '*PS Roaming*.csv'
        ],
        'sheet_name': 'User_PS_Roaming',
        'filter': [
            {'column': 'Mobile country code', 'value': 606, 'operator': 'eq'},
            {'column': 'Mobile network code', 'value': 1, 'operator': 'eq'}
        ],
        'aggregate': {
            'groupby': 'Time',  # Group by hour first
            'sum_columns': [
                'Gb mode maximum users with act PDP context per PLMN(number)',
                'Iu mode maximum users with act PDP context per PLMN(number)',
                'S1 Mode Maximum PDN Connection Number per PLMN(number)'
            ],
            'then_max': True  # Then take MAX per day
        }
    },
    'PS_Subscribers': {
        'patterns': [
            '* (PS users (2G-3G-4G)).csv',
            '*(PS users (2G-3G-4G)).csv',
            '*PS users*.csv',
            '*PS users*.csv'
        ],
        'sheet_name': 'User_PS_Subscribers'
    },
    'VoLTE_Users': {
        'patterns': [
            '* (VoLTE users).csv',
            '*(VoLTE users).csv',
            '*VoLTE*.csv',
            '*VoLTE users*.csv'
        ],
        'sheet_name': 'User_VoLTE'
    }
}


def process_user_kpis(day_folder, log_callback=None):
    """
    Process all 5 user KPI files in the day folder.
    Returns a dictionary: sheet_name -> DataFrame

    KEY CHANGE: Aggregates by Date taking MAX for all numeric columns.
    This ensures we capture the peak value for each day.
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📊 PROCESSING USER KPIs (MAX per day)")
    log("=" * 60)

    results = {}

    for kpi_name, config in USER_KPI_FILES.items():
        patterns = config['patterns']
        sheet_name = config['sheet_name']

        file_path = find_file(day_folder, patterns)

        if file_path:
            log(f"📄 Processing {kpi_name}: {os.path.basename(file_path)}")
            try:
                # Use the universal reader
                df = read_csv_skip_metadata(file_path)

                if df is not None and not df.empty:
                    # Convert Time to Date if needed
                    if 'Time' in df.columns:
                        df['DateTime'] = pd.to_datetime(df['Time'])
                        df['Date'] = df['DateTime'].dt.strftime('%Y-%m-%d')
                        df['Hour'] = df['DateTime'].dt.hour
                    elif 'Date' in df.columns:
                        df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

                    # Apply filter if configured
                    if 'filter' in config:
                        filter_config = config['filter']

                        # Handle single filter
                        if isinstance(filter_config, dict):
                            col = filter_config['column']
                            val = filter_config['value']
                            contains = filter_config.get('contains', False)
                            operator = filter_config.get('operator', 'eq')

                            if col in df.columns:
                                if contains:
                                    df = df[df[col].astype(str).str.contains(str(val), na=False)]
                                elif operator == 'eq':
                                    df = df[df[col] == val]
                                log(f"   🔍 Filtered {col}={val}: {len(df)} rows")

                        # Handle multiple filters (AND condition)
                        elif isinstance(filter_config, list):
                            for f in filter_config:
                                col = f['column']
                                val = f['value']
                                operator = f.get('operator', 'eq')

                                if col in df.columns:
                                    if operator == 'eq':
                                        df = df[df[col] == val]
                                    elif operator == 'ne':
                                        df = df[df[col] != val]
                                    elif operator == 'gt':
                                        df = df[df[col] > val]
                                    elif operator == 'lt':
                                        df = df[df[col] < val]
                                    log(f"   🔍 Filtered {col}={val}: {len(df)} rows")

                    # --- SPECIAL AGGREGATION: Group by Time (hour) and SUM, then MAX per day ---
                    if 'aggregate' in config and 'Time' in df.columns:
                        agg_config = config['aggregate']
                        groupby_col = agg_config.get('groupby', 'Time')
                        sum_cols = agg_config.get('sum_columns', [])
                        then_max = agg_config.get('then_max', True)

                        # Filter to only columns that exist in the dataframe
                        existing_sum_cols = [col for col in sum_cols if col in df.columns]

                        if existing_sum_cols and groupby_col in df.columns:
                            # Step 1: Group by Time (hour) and SUM the specified columns
                            # Need to group by Date, Hour to keep them separate
                            if 'Date' in df.columns and 'Hour' in df.columns:
                                # Group by Date and Hour
                                df_grouped = df.groupby(['Date', 'Hour'])[existing_sum_cols].sum().reset_index()
                                log(f"   📊 Grouped by Date+Hour, summed {len(existing_sum_cols)} columns: {len(df_grouped)} rows")

                                # Step 2: Group by Date and take MAX for each column
                                if then_max:
                                    df_aggregated = df_grouped.groupby('Date')[existing_sum_cols].max().reset_index()
                                    # Keep Date column
                                    df_aggregated['Date'] = df_aggregated['Date']
                                    log(f"   📊 Aggregated by Date (MAX): {len(df_aggregated)} rows")
                                else:
                                    df_aggregated = df_grouped
                            else:
                                # Fallback: just group by Date
                                df_aggregated = df.groupby('Date')[existing_sum_cols].max().reset_index()
                                log(f"   📊 Aggregated by Date (MAX): {len(df_aggregated)} rows")

                            # Replace df with aggregated version
                            df = df_aggregated
                        else:
                            log(f"   ⚠️ No sum columns found for aggregation")

                    # --- STANDARD AGGREGATION: Group by Date taking MAX for all numeric columns ---
                    elif 'Date' in df.columns:
                        # Identify numeric columns
                        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()

                        if numeric_cols:
                            # Group by Date and take MAX for all numeric columns
                            df_aggregated = df.groupby('Date')[numeric_cols].max().reset_index()

                            # Also include any non-numeric columns that should be carried through
                            non_numeric_cols = [col for col in df.columns if
                                                col not in numeric_cols and col not in ['Date', 'DateTime', 'Hour']]
                            if non_numeric_cols:
                                # For non-numeric columns (like 'Whole Network', 'Integrity'),
                                # take the first value per date (assuming it's constant)
                                for col in non_numeric_cols:
                                    if col in df.columns:
                                        first_values = df.groupby('Date')[col].first().reset_index()
                                        df_aggregated = df_aggregated.merge(first_values, on='Date', how='left')

                            log(f"   📊 Aggregated {len(df)} rows → {len(df_aggregated)} rows (MAX per day)")
                            df = df_aggregated
                        else:
                            log(f"   ⚠️ No numeric columns found to aggregate")

                    # Remove duplicates by Date (if any remain)
                    if 'Date' in df.columns:
                        dup_count = df.duplicated(subset=['Date']).sum()
                        if dup_count > 0:
                            log(f"   ⚠️ Found {dup_count} duplicate rows, removing")
                            df = df.drop_duplicates(subset=['Date'])

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
                import traceback
                log(traceback.format_exc())
                results[sheet_name] = None
        else:
            log(f"   ⚠️ No file found for {kpi_name}")
            results[sheet_name] = None

    log("=" * 60)
    return results


def aggregate_user_data(results_dict):
    """
    Aggregate user data from the 5 sheets into a unified user summary table.

    KEY CHANGES:
    - All numeric values are MAX per day
    - PS Roaming: 2G, 3G, 4G kept as SEPARATE columns (NOT summed)
    - No "Total Roaming" column
    """
    if not results_dict:
        return None

    df_cs_roaming = results_dict.get('User_CS_Roaming')
    df_cs_subs = results_dict.get('User_CS_Subscribers')
    df_ps_roaming = results_dict.get('User_PS_Roaming')
    df_ps_subs = results_dict.get('User_PS_Subscribers')
    df_volte = results_dict.get('User_VoLTE')

    # Start with CS Subscribers as base
    base_df = None
    if df_cs_subs is not None and not df_cs_subs.empty:
        cols = ['Date']
        rename_map = {}

        # Find the right columns (flexible column names)
        for col in df_cs_subs.columns:
            if '2G Subscribers in VLR' in col or 'Number of 2G Subscribers in VLR' in col:
                cols.append(col)
                rename_map[col] = '2G CS user'
            elif '3G Subscribers in VLR' in col or 'Number of 3G Subscribers in VLR' in col:
                cols.append(col)
                rename_map[col] = '3G CS user'
            elif 'Total Number of Subscribers in VLR' in col or 'Total VLR Subscribers' in col:
                cols.append(col)
                rename_map[col] = 'Total VLR Subscribers'

        if len(cols) > 1:
            base_df = df_cs_subs[cols].copy()
            base_df = base_df.rename(columns=rename_map)
            # Ensure MAX per day (already aggregated, but just in case)
            numeric_cols = base_df.select_dtypes(include=['number']).columns.tolist()
            if numeric_cols:
                base_df = base_df.groupby('Date')[numeric_cols].max().reset_index()

    if base_df is None or base_df.empty:
        base_df = pd.DataFrame(columns=['Date'])

    # --- 1. CS Roaming (MAX per day) ---
    if df_cs_roaming is not None and not df_cs_roaming.empty:
        # The data should already be aggregated by Date with MAX
        # Look for the 'Number of Registered Subscribers(entries)' column
        # which has been summed per hour and then maxed per day
        roaming_col = None
        for col in df_cs_roaming.columns:
            if 'Number of Registered Subscribers(entries)' in col or 'Registered Subscribers' in col:
                roaming_col = col
                break

        if roaming_col and 'Date' in df_cs_roaming.columns:
            roaming_df = df_cs_roaming[['Date', roaming_col]].copy()
            roaming_df = roaming_df.rename(columns={roaming_col: 'Roaming CS (Almadar)'})
            base_df = pd.merge(base_df, roaming_df, on='Date', how='outer')

    # --- 2. PS Roaming (2G, 3G, 4G SEPARATE - NO SUMMING) ---
    if df_ps_roaming is not None and not df_ps_roaming.empty:
        # The data should already be aggregated with sums per hour and max per day
        rename_map = {}

        for col in df_ps_roaming.columns:
            if 'Gb mode maximum users with act PDP context per PLMN(number)' in col:
                rename_map[col] = 'Roaming 2G PS (Gb)'
            elif 'Iu mode maximum users with act PDP context per PLMN(number)' in col:
                rename_map[col] = 'Roaming 3G PS (Iu)'
            elif 'S1 Mode Maximum PDN Connection Number per PLMN(number)' in col:
                rename_map[col] = 'Roaming 4G PS (S1)'

        if rename_map and 'Date' in df_ps_roaming.columns:
            # Select only columns that exist and are in rename_map
            cols_to_keep = ['Date'] + list(rename_map.keys())
            existing_cols = [col for col in cols_to_keep if col in df_ps_roaming.columns]

            if len(existing_cols) > 1:
                ps_roam_df = df_ps_roaming[existing_cols].copy()
                ps_roam_df = ps_roam_df.rename(columns=rename_map)
                base_df = pd.merge(base_df, ps_roam_df, on='Date', how='outer')

    # --- 3. PS Subscribers (MAX per day) ---
    if df_ps_subs is not None and not df_ps_subs.empty:
        cols = ['Date']
        rename_map = {}

        for col in df_ps_subs.columns:
            if 'Gb mode maximum attached users' in col:
                cols.append(col)
                rename_map[col] = '2G PS user'
            elif 'Iu mode maximum attached users' in col:
                cols.append(col)
                rename_map[col] = '3G PS user'
            elif 'Maximum attached users' in col and 'Gb' not in col and 'Iu' not in col:
                cols.append(col)
                rename_map[col] = '4G PS user'

        if len(cols) > 1 and 'Date' in df_ps_subs.columns:
            ps_subs_df = df_ps_subs[cols].copy()
            numeric_cols = ps_subs_df.select_dtypes(include=['number']).columns.tolist()
            if numeric_cols:
                ps_subs_df = ps_subs_df.groupby('Date')[numeric_cols].max().reset_index()
            ps_subs_df = ps_subs_df.rename(columns=rename_map)
            base_df = pd.merge(base_df, ps_subs_df, on='Date', how='outer')

    # --- 4. VoLTE Users (MAX per day) ---
    if df_volte is not None and not df_volte.empty:
        volte_col = None
        for col in df_volte.columns:
            if 'VoLTE IMS subscribers' in col or 'VoLTE' in col:
                volte_col = col
                break

        if volte_col and 'Date' in df_volte.columns:
            volte_df = df_volte[['Date', volte_col]].copy()
            volte_df = volte_df.rename(columns={volte_col: 'VoLTE user'})
            volte_df = volte_df.groupby('Date').max().reset_index()
            base_df = pd.merge(base_df, volte_df, on='Date', how='outer')

    # --- 5. Compute derived columns (NO SUMMING FOR ROAMING) ---
    # Total Subscribers (CS + PS)
    if all(col in base_df.columns for col in ['2G CS user', '3G CS user', '4G PS user']):
        base_df['Total Subscribers'] = base_df['2G CS user'] + base_df['3G CS user'] + base_df['4G PS user']

    # Fill NaN with 0
    base_df = base_df.fillna(0)

    # Sort by Date
    base_df = base_df.sort_values('Date').reset_index(drop=True)

    return base_df


def get_user_summary_columns():
    """
    Return the expected columns for the User Summary sheet.
    This helps with creating the Excel sheet structure.
    """
    return [
        'Date',
        '2G CS user',
        '3G CS user',
        'Total VLR Subscribers',
        'Roaming CS (Almadar)',
        'Roaming 2G PS (Gb)',
        'Roaming 3G PS (Iu)',
        'Roaming 4G PS (S1)',
        '2G PS user',
        '3G PS user',
        '4G PS user',
        'VoLTE user',
        'Total Subscribers'
    ]


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python user_kpi_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]

    print("\n" + "=" * 60)
    print("TEST: USER KPI PROCESSING (MAX per day)")
    print("=" * 60)

    raw_results = process_user_kpis(test_folder)

    print("\n" + "=" * 60)
    print("RAW USER KPI RESULTS")
    print("=" * 60)

    for sheet_name, df in raw_results.items():
        print(f"\n{sheet_name}:")
        if df is not None and not df.empty:
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
            print(f"  Columns: {df.columns.tolist()}")
            if 'Date' in df.columns:
                dates = df['Date'].unique()
                print(f"  Date Range: {min(dates)} to {max(dates)}")
            print(f"  Sample:\n{df.head(3)}")
        else:
            print("  No data")

    print("\n" + "=" * 60)
    print("USER SUMMARY TABLE (MAX per day, Roaming separated)")
    print("=" * 60)

    summary_df = aggregate_user_data(raw_results)
    if summary_df is not None and not summary_df.empty:
        print(f"Rows: {len(summary_df)}, Columns: {len(summary_df.columns)}")
        print(f"Columns: {summary_df.columns.tolist()}")
        print("\nSample:")
        print(summary_df.head(10))

        # Verify roaming columns are separate
        roaming_cols = [c for c in summary_df.columns if 'Roaming' in c]
        print(f"\n📊 Roaming Columns (SEPARATE - NOT SUMMED):")
        for col in roaming_cols:
            print(f"  ✅ {col}")
    else:
        print("No summary data")