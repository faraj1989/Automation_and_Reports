import argparse  # <-- ADD THIS IMPORT
import pandas as pd
import re
import os
import shutil
from io import StringIO

from datetime import datetime
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from project_config import load_env_file

load_env_file()

# =============================================================
# CENTRALIZED DIRECTORIES (from GUI)
# =============================================================
DATA_ROOT = Path(os.environ.get("DATA_ROOT", r"C:\Users\user\Desktop\NOC_Data"))

# PS Traffic paths
PS_TRAFFIC_SOURCE_DIR = Path(os.environ.get("PS_TRAFFIC_SOURCE_DIR", DATA_ROOT / "Subscribers" / "Raw Data"))
PS_TRAFFIC_OUTPUT_DIR = Path(os.environ.get("PS_TRAFFIC_OUTPUT_DIR", DATA_ROOT / "PS_Traffic_Output"))
PS_TRAFFIC_HISTORY_DIR = PS_TRAFFIC_OUTPUT_DIR / "Historical_Archive"

# Ensure directories exist
PS_TRAFFIC_SOURCE_DIR.mkdir(parents=True, exist_ok=True)
PS_TRAFFIC_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
PS_TRAFFIC_HISTORY_DIR.mkdir(parents=True, exist_ok=True)


# =============================================================


def clean_csv_content(file_content):
    """
    Clean CSV content by removing metadata headers and footer.
    Headers are removed by finding the line starting with 'Date,'
    Footer is removed by taking all but the last line.
    """
    lines = file_content.strip().split('\n')

    # Find the line where actual data starts (line containing 'Date,')
    start_idx = 0
    for i, line in enumerate(lines):
        if line.startswith('Date,'):
            start_idx = i
            break

    # Extract data rows (remove footer by taking all but last line)
    data_lines = lines[start_idx:-1]  # -1 removes the footer line

    # Remove any empty lines
    data_lines = [line for line in data_lines if line.strip()]

    return '\n'.join(data_lines)


def find_ps_traffic_files(unzipped_folder):
    """
    Find the PS Daily Traffic 2G, 3G, and 4G CSV files in the unzipped folder.
    Searches for folder starting with 'PS Daily Traffic_2G_3G_4G'
    then looks for CSV files ending with (PS Traffic 2G), (PS Traffic 3G), (PS Traffic 4G)
    """
    if not os.path.exists(unzipped_folder):
        print(f"⚠️ Unzipped folder does not exist: {unzipped_folder}")
        return None, None, None

    print(f"🔍 Searching for PS Traffic files in: {unzipped_folder}")

    # First, find the PS Traffic folder (starting with 'PS Daily Traffic_2G_3G_4G')
    ps_traffic_folder = None
    for root, dirs, files in os.walk(unzipped_folder):
        for dir_name in dirs:
            if dir_name.startswith('PS Daily Traffic_2G_3G_4G'):
                ps_traffic_folder = os.path.join(root, dir_name)
                print(f"  ✅ Found PS Traffic folder: {ps_traffic_folder}")
                break
        if ps_traffic_folder:
            break

    # If no folder found, try to find any folder containing PS Daily Traffic
    if not ps_traffic_folder:
        for root, dirs, files in os.walk(unzipped_folder):
            for dir_name in dirs:
                if 'PS Daily Traffic' in dir_name:
                    ps_traffic_folder = os.path.join(root, dir_name)
                    print(f"  ✅ Found PS Traffic folder (alternative): {ps_traffic_folder}")
                    break
            if ps_traffic_folder:
                break

    if not ps_traffic_folder:
        print("❌ PS Daily Traffic folder not found in unzipped directory")
        # List what is available
        print("\n📂 Available folders in unzipped:")
        if os.path.exists(unzipped_folder):
            for item in os.listdir(unzipped_folder):
                item_path = os.path.join(unzipped_folder, item)
                if os.path.isdir(item_path):
                    print(f"  - {item}")
                elif os.path.isfile(item_path) and item.endswith('.csv'):
                    print(f"  - {item} (file)")
        return None, None, None

    print(f"📂 PS Traffic folder: {ps_traffic_folder}")

    # Find the three files with exact naming pattern
    file_2g = None
    file_3g = None
    file_4g = None

    # List all files in the folder
    all_files = [f for f in os.listdir(ps_traffic_folder) if f.endswith('.csv')]
    print(f"\n  Found {len(all_files)} CSV files in folder:")
    for file in all_files:
        print(f"    - {file}")

        # Check for exact patterns
        if '(PS Traffic 2G)' in file or '(PS Traffic 2G).csv' in file:
            file_2g = os.path.join(ps_traffic_folder, file)
            print(f"      ✅ Identified as 2G file")
        elif '(PS Traffic 3G)' in file or '(PS Traffic 3G).csv' in file:
            file_3g = os.path.join(ps_traffic_folder, file)
            print(f"      ✅ Identified as 3G file")
        elif '(PS Traffic 4G)' in file or '(PS Traffic 4G).csv' in file:
            file_4g = os.path.join(ps_traffic_folder, file)
            print(f"      ✅ Identified as 4G file")

    # If still not found, try more flexible matching
    if not file_2g or not file_3g or not file_4g:
        print("\n  🔍 Trying alternative matching (looking for technology indicators)...")
        for file in all_files:
            # Look for 2G indicators
            if '2G' in file and 'Traffic' in file:
                if not file_2g:
                    file_2g = os.path.join(ps_traffic_folder, file)
                    print(f"    ✅ Found 2G (alternative): {file}")
            # Look for 3G indicators
            elif '3G' in file and 'Traffic' in file:
                if not file_3g:
                    file_3g = os.path.join(ps_traffic_folder, file)
                    print(f"    ✅ Found 3G (alternative): {file}")
            # Look for 4G indicators
            elif '4G' in file and 'Traffic' in file:
                if not file_4g:
                    file_4g = os.path.join(ps_traffic_folder, file)
                    print(f"    ✅ Found 4G (alternative): {file}")

    # Final check - if we have files but couldn't match by name,
    # try to identify by reading the CSV header
    if (not file_2g or not file_3g or not file_4g) and all_files:
        print("\n  🔍 Trying to identify remaining files by their content...")
        for file in all_files:
            file_path = os.path.join(ps_traffic_folder, file)

            # Skip already identified files
            if file_path in [file_2g, file_3g, file_4g]:
                continue

            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    # Read first 30 lines to find header
                    content = f.read()
                    lines = content.split('\n')

                    # Find header line
                    header_line = None
                    for line in lines:
                        if line.startswith('Date,'):
                            header_line = line
                            break

                    if header_line:
                        # Identify by column names
                        if 'eGBTS' in header_line and 'PS Traffic(GB)' in header_line:
                            file_2g = file_path
                            print(f"    ✅ Identified as 2G by columns: {file}")
                        elif 'NodeB' in header_line and 'PS traffic (GB)' in header_line:
                            file_3g = file_path
                            print(f"    ✅ Identified as 3G by columns: {file}")
                        elif 'eNodeB Name' in header_line and 'Downlink Traffic Volume(GB)' in header_line:
                            file_4g = file_path
                            print(f"    ✅ Identified as 4G by columns: {file}")
            except Exception as e:
                print(f"    ⚠️ Could not read {file}: {e}")

    print(f"\n📄 Final file list:")
    print(f"  2G: {os.path.basename(file_2g) if file_2g else '❌ NOT FOUND'}")
    print(f"  3G: {os.path.basename(file_3g) if file_3g else '❌ NOT FOUND'}")
    print(f"  4G: {os.path.basename(file_4g) if file_4g else '❌ NOT FOUND'}")

    return file_2g, file_3g, file_4g


def load_traffic_file(file_path, file_type):
    """
    Load and clean a single traffic CSV file.
    """
    if not file_path or not os.path.exists(file_path):
        print(f"⚠️ {file_type} file not found: {file_path}")
        return None

    print(f"  Reading {file_type} file: {os.path.basename(file_path)}")

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        clean_content = clean_csv_content(content)
        df = pd.read_csv(StringIO(clean_content))

        # Remove footer if present (last row)
        df = df[:-1] if len(df) > 0 else df

        print(f"  - Loaded {len(df)} records")
        print(f"  - Columns: {list(df.columns)}")

        return df
    except Exception as e:
        print(f"  ❌ Error loading {file_type} file: {e}")
        return None


def process_2g_data(df):
    """Process 2G traffic data."""
    if df is None or df.empty:
        return pd.DataFrame()

    try:
        df_2g = df[['Date', 'eGBTS', 'PS Traffic(GB)']].copy()
        df_2g.columns = ['Date', 'Site_Name', 'Traffic_GB']
        df_2g['Technology'] = '2G'
        return df_2g
    except KeyError as e:
        print(f"  ❌ Error processing 2G data: {e}")
        print(f"  Available columns: {list(df.columns)}")
        return pd.DataFrame()


def process_3g_data(df):
    """Process 3G traffic data."""
    if df is None or df.empty:
        return pd.DataFrame()

    try:
        # Use PS traffic (GB) column
        df_3g = df[['Date', 'NodeB', 'PS traffic (GB)']].copy()
        df_3g.columns = ['Date', 'Site_Name', 'Traffic_GB']
        df_3g['Technology'] = '3G'
        return df_3g
    except KeyError as e:
        print(f"  ❌ Error processing 3G data: {e}")
        print(f"  Available columns: {list(df.columns)}")
        return pd.DataFrame()


def process_4g_data(df):
    """Process 4G traffic data (DL + UL)."""
    if df is None or df.empty:
        return pd.DataFrame()

    try:
        df_4g = df[['Date', 'eNodeB Name', 'Downlink Traffic Volume(GB)', 'UL Traffic  Volume(GB)']].copy()
        df_4g.columns = ['Date', 'Site_Name', 'DL_GB', 'UL_GB']
        df_4g['Traffic_GB'] = df_4g['DL_GB'] + df_4g['UL_GB']
        df_4g['Technology'] = '4G'
        return df_4g[['Date', 'Site_Name', 'Traffic_GB', 'Technology']]
    except KeyError as e:
        print(f"  ❌ Error processing 4G data: {e}")
        print(f"  Available columns: {list(df.columns)}")
        return pd.DataFrame()


def extract_region(site_name):
    """Extract region from site name (prefix before numbers/underscore/parentheses)."""
    if pd.isna(site_name):
        return 'Unknown'
    site_str = str(site_name)
    # Remove parentheses and anything inside them (like FN)
    site_str = re.sub(r'\([^)]*\)', '', site_str)
    # Extract letters at the beginning (region name)
    match = re.match(r'^([A-Za-z]+)', site_str)
    if match:
        return match.group(1)
    return 'Other'


def analyze_site_coverage(data_frames):
    """Analyze which sites have which technologies."""
    print("\n" + "=" * 60)
    print("📊 SITE COVERAGE ANALYSIS")
    print("=" * 60)

    # Get unique sites per technology
    tech_sites = {}
    for tech, df in data_frames.items():
        if not df.empty:
            tech_sites[tech] = set(df['Site_Name'].unique())
            print(f"\n{tech} sites: {len(tech_sites[tech])}")

    if len(tech_sites) < 2:
        print("⚠️ Not enough technologies for coverage analysis")
        return None

    # Find intersections
    all_techs = list(tech_sites.keys())

    # Sites with all technologies
    if len(all_techs) >= 3 and all(tech in tech_sites for tech in ['2G', '3G', '4G']):
        all_three = tech_sites['2G'] & tech_sites['3G'] & tech_sites['4G']
        print(f"\n✅ Sites with ALL THREE technologies (2G+3G+4G): {len(all_three)}")
        if len(all_three) > 0 and len(all_three) <= 20:
            print(f"   {sorted(list(all_three))}")

    # Sites with 2G+3G (no 4G)
    if '2G' in tech_sites and '3G' in tech_sites:
        two_three = tech_sites['2G'] & tech_sites['3G'] - tech_sites.get('4G', set())
        print(f"\n📱 Sites with 2G+3G (no 4G): {len(two_three)}")
        if len(two_three) > 0 and len(two_three) <= 20:
            print(f"   {sorted(list(two_three))}")

    # Sites with 2G+4G (no 3G)
    if '2G' in tech_sites and '4G' in tech_sites:
        two_four = tech_sites['2G'] & tech_sites['4G'] - tech_sites.get('3G', set())
        print(f"\n📱 Sites with 2G+4G (no 3G): {len(two_four)}")
        if len(two_four) > 0 and len(two_four) <= 20:
            print(f"   {sorted(list(two_four))}")

    # Sites with 3G+4G (no 2G)
    if '3G' in tech_sites and '4G' in tech_sites:
        three_four = tech_sites['3G'] & tech_sites['4G'] - tech_sites.get('2G', set())
        print(f"\n📱 Sites with 3G+4G (no 2G): {len(three_four)}")
        if len(three_four) > 0 and len(three_four) <= 20:
            print(f"   {sorted(list(three_four))}")

    # Sites with only 2G
    if '2G' in tech_sites:
        only_2g = tech_sites['2G'] - tech_sites.get('3G', set()) - tech_sites.get('4G', set())
        print(f"\n📱 Sites with ONLY 2G: {len(only_2g)}")
        if len(only_2g) > 0 and len(only_2g) <= 20:
            print(f"   {sorted(list(only_2g))}")

    # Sites with only 3G
    if '3G' in tech_sites:
        only_3g = tech_sites['3G'] - tech_sites.get('2G', set()) - tech_sites.get('4G', set())
        print(f"\n📱 Sites with ONLY 3G: {len(only_3g)}")
        if len(only_3g) > 0 and len(only_3g) <= 20:
            print(f"   {sorted(list(only_3g))}")

    # Sites with only 4G
    if '4G' in tech_sites:
        only_4g = tech_sites['4G'] - tech_sites.get('2G', set()) - tech_sites.get('3G', set())
        print(f"\n📱 Sites with ONLY 4G: {len(only_4g)}")
        if len(only_4g) > 0 and len(only_4g) <= 20:
            print(f"   {sorted(list(only_4g))}")

    # Total unique sites across all technologies
    all_sites = set()
    for sites in tech_sites.values():
        all_sites.update(sites)
    print(f"\n📍 TOTAL UNIQUE SITES (all technologies): {len(all_sites)}")

    return tech_sites


def combine_traffic_data(data_frames):
    """
    Combine 2G, 3G, and 4G data into a single pivoted dataframe.
    """
    # Filter out empty dataframes
    valid_frames = {tech: df for tech, df in data_frames.items() if not df.empty}

    if not valid_frames:
        print("❌ No valid data to combine")
        return pd.DataFrame()

    # Combine all technologies
    all_data = pd.concat(valid_frames.values(), ignore_index=True)

    print(f"\nTotal records before processing: {len(all_data)}")
    print(f"Records by technology:\n{all_data['Technology'].value_counts()}")

    # Convert date - try YYYY-MM-DD first, then infer
    all_data['Date'] = pd.to_datetime(all_data['Date'], errors='coerce')

    # If all dates are NaT, try another format
    if all_data['Date'].isna().all():
        print("  - Trying alternative date format...")
        all_data['Date'] = pd.to_datetime(all_data['Date'], format='%d/%m/%Y', errors='coerce')

    # Count successful conversions
    valid_dates = all_data['Date'].notna().sum()
    print(f"Valid dates: {valid_dates} out of {len(all_data)}")

    # Remove any rows with invalid dates
    invalid_dates = all_data['Date'].isna().sum()
    if invalid_dates > 0:
        print(f"  - Removing {invalid_dates} rows with invalid dates")
        all_data = all_data.dropna(subset=['Date'])

    if len(all_data) == 0:
        print("  - Warning: No valid dates found after parsing!")
        return pd.DataFrame()

    # Pivot to get separate columns for each technology
    pivoted = all_data.pivot_table(
        index=['Date', 'Site_Name'],
        columns='Technology',
        values='Traffic_GB',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Ensure all technology columns exist
    for tech in ['2G', '3G', '4G']:
        if tech not in pivoted.columns:
            pivoted[tech] = 0

    # Rename columns for clarity
    pivoted.columns = ['Date', 'Site_Name', '2G_Traffic_GB', '3G_Traffic_GB', '4G_Traffic_GB']

    # Calculate total traffic
    pivoted['Total_Traffic_GB'] = pivoted['2G_Traffic_GB'] + pivoted['3G_Traffic_GB'] + pivoted['4G_Traffic_GB']

    # Extract region
    pivoted['Region'] = pivoted['Site_Name'].apply(extract_region)

    # Sort by Date and Site Name
    pivoted = pivoted.sort_values(['Date', 'Site_Name']).reset_index(drop=True)

    # Format date for display
    pivoted['Date_Formatted'] = pivoted['Date'].dt.strftime('%d-%m-%Y')

    return pivoted


def generate_summary_reports(combined_df):
    """
    Generate summary reports from combined data.
    """
    reports = {}

    if len(combined_df) == 0:
        print("  - Warning: No data to generate summary reports")
        return reports

    # 1. Daily summary
    daily_summary = combined_df.groupby('Date').agg({
        '2G_Traffic_GB': 'sum',
        '3G_Traffic_GB': 'sum',
        '4G_Traffic_GB': 'sum',
        'Total_Traffic_GB': 'sum'
    }).reset_index()
    daily_summary['Date_Formatted'] = daily_summary['Date'].dt.strftime('%d-%m-%Y')
    daily_summary = daily_summary.sort_values('Date')
    reports['Daily_Summary'] = daily_summary

    # 2. Regional summary
    regional_summary = combined_df.groupby('Region').agg({
        '2G_Traffic_GB': 'sum',
        '3G_Traffic_GB': 'sum',
        '4G_Traffic_GB': 'sum',
        'Total_Traffic_GB': 'sum'
    }).reset_index()
    regional_summary = regional_summary.sort_values('Total_Traffic_GB', ascending=False)
    reports['Regional_Summary'] = regional_summary

    # 3. Top 50 sites by total traffic
    top_sites = combined_df.groupby('Site_Name').agg({
        'Region': 'first',
        '2G_Traffic_GB': 'sum',
        '3G_Traffic_GB': 'sum',
        '4G_Traffic_GB': 'sum',
        'Total_Traffic_GB': 'sum'
    }).reset_index()
    top_sites = top_sites.nlargest(50, 'Total_Traffic_GB')
    reports['Top_50_Sites'] = top_sites

    # 4. Worst 50 sites by total traffic (excluding zero traffic sites)
    worst_sites = combined_df.groupby('Site_Name').agg({
        'Region': 'first',
        '2G_Traffic_GB': 'sum',
        '3G_Traffic_GB': 'sum',
        '4G_Traffic_GB': 'sum',
        'Total_Traffic_GB': 'sum'
    }).reset_index()
    worst_sites = worst_sites[worst_sites['Total_Traffic_GB'] > 0]
    worst_sites = worst_sites.nsmallest(50, 'Total_Traffic_GB')
    reports['Worst_50_Sites'] = worst_sites

    # 5. Technology share by date
    tech_share = combined_df.groupby('Date').agg({
        '2G_Traffic_GB': 'sum',
        '3G_Traffic_GB': 'sum',
        '4G_Traffic_GB': 'sum'
    }).reset_index()
    total_by_date = tech_share[['2G_Traffic_GB', '3G_Traffic_GB', '4G_Traffic_GB']].sum(axis=1)
    for tech in ['2G', '3G', '4G']:
        tech_share[f'{tech}_Share_%'] = (tech_share[f'{tech}_Traffic_GB'] / total_by_date * 100).round(2)
    tech_share['Date_Formatted'] = tech_share['Date'].dt.strftime('%d-%m-%Y')
    reports['Tech_Share_by_Date'] = tech_share

    # 6. Site coverage report
    if not combined_df.empty:
        site_coverage = combined_df.groupby('Site_Name').agg({
            'Region': 'first',
            '2G_Traffic_GB': lambda x: (x > 0).sum() if len(x) > 0 else 0,
            '3G_Traffic_GB': lambda x: (x > 0).sum() if len(x) > 0 else 0,
            '4G_Traffic_GB': lambda x: (x > 0).sum() if len(x) > 0 else 0,
            'Total_Traffic_GB': 'sum'
        }).reset_index()
        site_coverage.columns = ['Site_Name', 'Region', 'Days_With_2G', 'Days_With_3G', 'Days_With_4G',
                                 'Total_Traffic_GB']

        site_coverage['Has_2G'] = site_coverage['Days_With_2G'] > 0
        site_coverage['Has_3G'] = site_coverage['Days_With_3G'] > 0
        site_coverage['Has_4G'] = site_coverage['Days_With_4G'] > 0

        def get_tech_combo(row):
            combo = []
            if row['Has_2G']: combo.append('2G')
            if row['Has_3G']: combo.append('3G')
            if row['Has_4G']: combo.append('4G')
            return '+'.join(combo) if combo else 'None'

        site_coverage['Tech_Combination'] = site_coverage.apply(get_tech_combo, axis=1)
        reports['Site_Coverage'] = site_coverage

    return reports


def save_historical_archive(combined_df, reports, archive_dir, date_range):
    """
    Save processed data to historical archive with timestamp.
    """
    # Create archive directory if it doesn't exist
    if not os.path.exists(archive_dir):
        os.makedirs(archive_dir)
        print(f"  📁 Created archive directory: {archive_dir}")

    # Create timestamp for archive
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Create archive filename with date range and timestamp
    archive_filename = f"PS_Traffic_History_{date_range}_{timestamp}.xlsx"
    archive_path = os.path.join(archive_dir, archive_filename)

    # Also save a CSV version for easy import
    csv_filename = f"PS_Traffic_History_{date_range}_{timestamp}.csv"
    csv_path = os.path.join(archive_dir, csv_filename)

    print(f"  📁 Archiving to: {archive_path}")

    # Save to Excel with all sheets
    with pd.ExcelWriter(archive_path, engine='openpyxl') as writer:
        # Save main data
        if len(combined_df) > 0:
            export_df = combined_df.copy()
            export_df['Date'] = export_df['Date'].dt.strftime('%d-%m-%Y')
            export_df = export_df.drop('Date_Formatted', axis=1)
            export_df.to_excel(writer, sheet_name='All_Traffic_Data', index=False)
        else:
            empty_df = pd.DataFrame(columns=['Date', 'Site_Name', '2G_Traffic_GB', '3G_Traffic_GB',
                                             '4G_Traffic_GB', 'Total_Traffic_GB', 'Region'])
            empty_df.to_excel(writer, sheet_name='All_Traffic_Data', index=False)

        # Save summary reports
        for sheet_name, df in reports.items():
            if not df.empty:
                if 'Date' in df.columns and 'Date_Formatted' in df.columns:
                    df = df.copy()
                    df['Date'] = df['Date_Formatted']
                    df = df.drop('Date_Formatted', axis=1)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Auto-adjust column widths
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                sheet.column_dimensions[column_letter].width = adjusted_width

    # Save CSV version (main data only)
    if len(combined_df) > 0:
        export_df = combined_df.copy()
        export_df['Date'] = export_df['Date'].dt.strftime('%d-%m-%Y')
        export_df = export_df.drop('Date_Formatted', axis=1)
        export_df.to_csv(csv_path, index=False)
        print(f"  📁 CSV version: {csv_path}")

    return archive_path, csv_path


def save_to_excel(combined_df, reports, output_path):
    """
    Save all data to Excel with multiple sheets.
    """
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Main combined data
        if len(combined_df) > 0:
            export_df = combined_df.copy()
            export_df['Date'] = export_df['Date'].dt.strftime('%d-%m-%Y')
            export_df = export_df.drop('Date_Formatted', axis=1)
            export_df.to_excel(writer, sheet_name='All_Traffic_Data', index=False)
        else:
            empty_df = pd.DataFrame(columns=['Date', 'Site_Name', '2G_Traffic_GB', '3G_Traffic_GB',
                                             '4G_Traffic_GB', 'Total_Traffic_GB', 'Region'])
            empty_df.to_excel(writer, sheet_name='All_Traffic_Data', index=False)

        # Summary reports
        for sheet_name, df in reports.items():
            if not df.empty:
                if 'Date' in df.columns and 'Date_Formatted' in df.columns:
                    df = df.copy()
                    df['Date'] = df['Date_Formatted']
                    df = df.drop('Date_Formatted', axis=1)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Auto-adjust column widths
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                sheet.column_dimensions[column_letter].width = adjusted_width

    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Process PS Traffic raw exports into combined traffic reports")
    parser.add_argument("--base-dir", type=str, default=str(PS_TRAFFIC_SOURCE_DIR),
                        help="Base folder containing dated raw data folders")
    parser.add_argument("--date", type=str, default=None,
                        help="Date folder name in YYYYMMDD format")
    parser.add_argument("--output-dir", type=str, default=str(PS_TRAFFIC_OUTPUT_DIR),
                        help="Output folder for saved reports")
    parser.add_argument("--archive-dir", type=str, default=str(PS_TRAFFIC_HISTORY_DIR),
                        help="Archive folder for historical output")
    return parser.parse_args()


def main():
    """
    Main function to process PS Daily Traffic files from unzipped folder.
    """
    args = parse_args()

    base_dir = Path(args.base_dir)
    date_str = args.date if args.date else datetime.now().strftime("%Y%m%d")
    date_folder = base_dir / date_str
    output_folder = Path(args.output_dir) if args.output_dir else date_folder / "output"
    archive_dir = Path(args.archive_dir) if args.archive_dir else PS_TRAFFIC_HISTORY_DIR

    print("=" * 60)
    print("📊 PS Traffic Processing")
    print("=" * 60)
    print(f"📂 Date folder: {date_folder}")
    print(f"📁 Output folder: {output_folder}")
    print(f"📁 Archive folder: {archive_dir}")

    unzipped_folder = date_folder / "unzipped"

    # Find the PS Traffic files
    file_2g, file_3g, file_4g = find_ps_traffic_files(str(unzipped_folder))

    if not file_2g and not file_3g and not file_4g:
        print("❌ No PS Traffic files found!")
        return None

    # Load and process each file
    data_frames = {}

    df_2g_raw = load_traffic_file(file_2g, "2G")
    if df_2g_raw is not None:
        data_frames['2G'] = process_2g_data(df_2g_raw)

    df_3g_raw = load_traffic_file(file_3g, "3G")
    if df_3g_raw is not None:
        data_frames['3G'] = process_3g_data(df_3g_raw)

    df_4g_raw = load_traffic_file(file_4g, "4G")
    if df_4g_raw is not None:
        data_frames['4G'] = process_4g_data(df_4g_raw)

    if not data_frames:
        print("❌ No data loaded!")
        return None

    # Analyze site coverage
    analyze_site_coverage(data_frames)

    # Combine data
    print("\n🔄 Combining data from all technologies...")
    combined_df = combine_traffic_data(data_frames)

    if len(combined_df) == 0:
        print("❌ Error: No valid data after processing!")
        return None

    # Generate summaries
    print("\n📊 Generating summary reports...")
    reports = generate_summary_reports(combined_df)

    # Create output folder
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    # Save to Excel with timestamp
    today_str = datetime.now().strftime("%Y-%m-%d")
    output_filename = f"PS_Traffic_Combined_Report_{today_str}.xlsx"
    output_path = output_folder / output_filename

    print(f"\n💾 Saving to Excel file: {output_path}")
    save_to_excel(combined_df, reports, output_path)

    # Save historical archive
    print("\n📁 Creating historical archive...")

    # Determine date range for archive filename
    if len(combined_df) > 0 and combined_df['Date'].notna().any():
        min_date = combined_df['Date'].min()
        max_date = combined_df['Date'].max()
        if pd.notna(min_date) and pd.notna(max_date):
            date_range = f"{min_date.strftime('%Y%m%d')}_{max_date.strftime('%Y%m%d')}"
        else:
            date_range = "nodate"
    else:
        date_range = "nodata"

    # Create archive directory
    archive_dir = Path(archive_dir)
    archive_dir.mkdir(parents=True, exist_ok=True)

    # Save to archive
    archive_path, csv_path = save_historical_archive(combined_df, reports, str(archive_dir), date_range)

    # Print statistics
    print("\n" + "=" * 60)
    print("✅ Processing Complete!")
    print("=" * 60)

    if len(combined_df) > 0 and combined_df['Date'].notna().any():
        min_date = combined_df['Date'].min()
        max_date = combined_df['Date'].max()
        if pd.notna(min_date) and pd.notna(max_date):
            print(f"📅 Date Range: {min_date.strftime('%d-%m-%Y')} to {max_date.strftime('%d-%m-%Y')}")

    print(f"📍 Total Sites: {combined_df['Site_Name'].nunique()}")
    print(f"📈 Total Records: {len(combined_df)}")
    print(f"🌍 Regions Found: {combined_df['Region'].nunique()}")
    print(f"\n📊 Total Traffic Summary:")
    print(f"   2G: {combined_df['2G_Traffic_GB'].sum():,.2f} GB")
    print(f"   3G: {combined_df['3G_Traffic_GB'].sum():,.2f} GB")
    print(f"   4G: {combined_df['4G_Traffic_GB'].sum():,.2f} GB")
    print(f"   Total: {combined_df['Total_Traffic_GB'].sum():,.2f} GB")
    print(f"\n📁 Output file: {output_path}")
    print(f"📁 Archive file: {archive_path}")

    return combined_df


# Run the script
if __name__ == "__main__":
    raise SystemExit(main())
