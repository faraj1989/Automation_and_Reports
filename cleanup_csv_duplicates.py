#!/usr/bin/env python3
"""
Libyana NPM - CSV Cleanup & Deduplication Script
Removes duplicate rows from all existing CSV files.
Normalizes all dates to YYYY-MM-DD format.
Removes completely empty rows.
Generates detailed cleanup report.
"""

import os
import pandas as pd
import logging
from datetime import datetime

# Setup logging
log_file = "csv_cleanup.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

CSV_FOLDER = "output/csv"


def normalize_date(date_value):
    """Normalize any date format to YYYY-MM-DD"""
    if date_value is None or pd.isna(date_value):
        return None
    try:
        date_str = str(date_value).strip()
        if not date_str:
            return None
        dt = pd.to_datetime(date_str, errors='coerce')
        if pd.isna(dt):
            logger.warning(f"Could not parse date: {date_value}")
            return str(date_value)
        return dt.strftime('%Y-%m-%d')
    except Exception as e:
        logger.warning(f"Error normalizing date {date_value}: {e}")
        return str(date_value)


def cleanup_csv_file(filepath, key_columns=None):
    """
    Clean up a single CSV file:
    1. Remove completely empty rows (all NaN)
    2. Normalize date columns to YYYY-MM-DD
    3. Remove duplicate rows
    
    Returns: (rows_before, rows_after, changes_summary)
    """
    try:
        df = pd.read_csv(filepath, encoding='utf-8', dtype={'Date': str, 'day': str})
        rows_before = len(df)
        
        changes = {
            'empty_rows_removed': 0,
            'dates_normalized': 0,
            'duplicates_removed': 0
        }
        
        # 1. Remove completely empty rows
        df_clean = df.dropna(how='all')
        changes['empty_rows_removed'] = rows_before - len(df_clean)
        
        # 2. Normalize date columns
        date_cols = ['Date', 'day']
        for col in date_cols:
            if col in df_clean.columns:
                before = df_clean[col].nunique()
                df_clean[col] = df_clean[col].apply(normalize_date)
                after = df_clean[col].nunique()
                if before != after:
                    changes['dates_normalized'] += 1
                    logger.info(f"  Normalized {col}: unique values {before} -> {after}")
        
        # 3. Remove duplicates
        rows_before_dedup = len(df_clean)
        
        if key_columns:
            # Use specific key columns
            available_keys = [col for col in key_columns if col in df_clean.columns]
            if available_keys:
                df_clean = df_clean.drop_duplicates(subset=available_keys, keep='last')
                logger.debug(f"  Removed duplicates by: {available_keys}")
            else:
                df_clean = df_clean.drop_duplicates(keep='last')
        else:
            df_clean = df_clean.drop_duplicates(keep='last')
        
        changes['duplicates_removed'] = rows_before_dedup - len(df_clean)
        
        rows_after = len(df_clean)
        
        # Save cleaned file only if changes were made
        if changes['empty_rows_removed'] > 0 or changes['duplicates_removed'] > 0:
            df_clean.to_csv(filepath, index=False, encoding='utf-8')
            logger.info(f"✅ Saved cleaned file: {rows_before} -> {rows_after} rows")
        else:
            logger.info(f"✅ No changes needed: {rows_after} rows")
        
        return rows_before, rows_after, changes
        
    except Exception as e:
        logger.error(f"❌ Error processing {filepath}: {e}")
        return 0, 0, {'error': str(e)}


def main():
    """Main cleanup script"""
    logger.info("=" * 80)
    logger.info("🧹 LIBYANA CSV CLEANUP & DEDUPLICATION")
    logger.info(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 80)
    
    if not os.path.exists(CSV_FOLDER):
        logger.error(f"CSV folder not found: {CSV_FOLDER}")
        return
    
    # Define cleanup rules for each file type
    cleanup_rules = {
        'SiteSummary.csv': ['day'],
        'User_Summary.csv': ['Date'],
        '2G_Cell_CSBH.csv': ['Date', 'Cell Name'],
        '3G_Cell_CSBH.csv': ['Date', 'Cell Name'],
        '4G_Cell_BH.csv': ['Date', 'Cell Name'],
        '2G_NWBH.csv': ['Date'],
        '3G_NWBH.csv': ['Date'],
        '4G_NWBH.csv': ['Date'],
        '2G_NW_Daily.csv': ['Date'],
        '3G_NW_Daily.csv': ['Date'],
        '4G_NW_Daily.csv': ['Date'],
        'Traffic_2G.csv': ['Date'],
        'Traffic_3G.csv': ['Date'],
        'Traffic_4G.csv': ['Date'],
        'Traffic_Network_2G.csv': ['Date'],
        'Traffic_Network_3G.csv': ['Date'],
        'Traffic_Network_4G.csv': ['Date'],
        'User_CS_Subscribers.csv': ['Date'],
        'User_PS_Subscribers.csv': ['Date'],
        'User_CS_Roaming.csv': ['Date'],
        'User_PS_Roaming.csv': ['Date'],
        'User_VoLTE.csv': ['Date'],
        'SiteDetail.csv': None,
    }
    
    # Process each CSV file
    total_rows_before = 0
    total_rows_after = 0
    total_empty_removed = 0
    total_dups_removed = 0
    files_cleaned = 0
    
    logger.info("\n📂 Processing CSV files...")
    logger.info("-" * 80)
    
    for filename, key_cols in cleanup_rules.items():
        filepath = os.path.join(CSV_FOLDER, filename)
        
        if not os.path.exists(filepath):
            logger.debug(f"⏭️  File not found (skipping): {filename}")
            continue
        
        logger.info(f"\n📄 {filename}")
        logger.info(f"   Path: {filepath}")
        
        rows_before, rows_after, changes = cleanup_csv_file(filepath, key_cols)
        
        if 'error' in changes:
            logger.error(f"   Error: {changes['error']}")
        else:
            total_rows_before += rows_before
            total_rows_after += rows_after
            total_empty_removed += changes['empty_rows_removed']
            total_dups_removed += changes['duplicates_removed']
            
            logger.info(f"   Empty rows removed: {changes['empty_rows_removed']}")
            logger.info(f"   Duplicates removed: {changes['duplicates_removed']}")
            logger.info(f"   Dates normalized: {changes['dates_normalized']}")
            logger.info(f"   Rows: {rows_before} -> {rows_after}")
            
            if rows_before != rows_after:
                files_cleaned += 1
    
    # Print summary
    logger.info("\n" + "=" * 80)
    logger.info("📊 CLEANUP SUMMARY")
    logger.info("=" * 80)
    logger.info(f"Files processed: {len(cleanup_rules)}")
    logger.info(f"Files cleaned: {files_cleaned}")
    logger.info(f"Total rows before: {total_rows_before:,}")
    logger.info(f"Total rows after: {total_rows_after:,}")
    logger.info(f"Empty rows removed: {total_empty_removed:,}")
    logger.info(f"Duplicate rows removed: {total_dups_removed:,}")
    logger.info(f"Total rows eliminated: {total_rows_before - total_rows_after:,}")
    
    if total_rows_before > 0:
        pct_removed = ((total_rows_before - total_rows_after) / total_rows_before) * 100
        logger.info(f"Percentage cleaned: {pct_removed:.2f}%")
    
    logger.info("=" * 80)
    logger.info(f"✅ Cleanup completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"📋 Log file: {log_file}")
    logger.info("=" * 80)


if __name__ == "__main__":
    main()
