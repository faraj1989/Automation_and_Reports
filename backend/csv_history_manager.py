#!/usr/bin/env python3
"""
Libyana NPM - CSV History Manager
Manages historical data as CSV files instead of Excel.
Combines all CSVs into a single Excel file at the end.
"""

import os
import logging
import pandas as pd
from datetime import datetime
from backend.site_processor import SITE_SUMMARY_HEADER
from backend.site_detail_processor import SITE_DETAIL_HEADER

logger = logging.getLogger(__name__)

# Output folder
OUTPUT_FOLDER = "output"
CSV_FOLDER = os.path.join(OUTPUT_FOLDER, "csv")


class CSVHistoryManager:
    """
    Manages historical data as CSV files.
    Each sheet is stored as a separate CSV file.
    """

    def __init__(self, output_folder=OUTPUT_FOLDER):
        self.output_folder = output_folder
        self.csv_folder = os.path.join(output_folder, "csv")
        self._ensure_folders()

    def _normalize_date(self, date_value):
        """
        Normalize any date format to YYYY-MM-DD.
        Handles mixed formats: "8/12/2026", "2026-08-12", etc.
        Returns None if invalid.
        """
        if date_value is None or pd.isna(date_value):
            return None
        try:
            date_str = str(date_value).strip()
            if not date_str:
                return None
            # Try to parse - handles multiple formats
            dt = pd.to_datetime(date_str, errors='coerce')
            if pd.isna(dt):
                logger.warning(f"Could not parse date: {date_value}")
                return str(date_value)
            normalized = dt.strftime('%Y-%m-%d')
            if date_str != normalized:
                logger.debug(f"Normalized date: {date_str} → {normalized}")
            return normalized
        except Exception as e:
            logger.warning(f"Error normalizing date {date_value}: {e}")
            return str(date_value)

    def _ensure_folders(self):
        """Create output folders if they don't exist."""
        os.makedirs(self.csv_folder, exist_ok=True)
        os.makedirs(self.output_folder, exist_ok=True)

    def _get_csv_path(self, sheet_name):
        """Get the CSV file path for a sheet."""
        # Sanitize sheet name for filename
        safe_name = sheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
        return os.path.join(self.csv_folder, f"{safe_name}.csv")

    def _read_csv(self, sheet_name):
        """Read a CSV file, return empty DataFrame if not exists."""
        csv_path = self._get_csv_path(sheet_name)
        if os.path.exists(csv_path):
            try:
                return pd.read_csv(csv_path)
            except Exception as e:
                logger.warning(f"Could not read {csv_path}: {e}")
                return pd.DataFrame()
        return pd.DataFrame()

    def _write_csv(self, sheet_name, df):
        """Write a DataFrame to CSV."""
        if df is None or df.empty:
            logger.warning(f"No data to write to {sheet_name}")
            return

        csv_path = self._get_csv_path(sheet_name)
        try:
            df.to_csv(csv_path, index=False)
            logger.info(f"✅ Wrote {len(df)} rows to {csv_path}")
        except Exception as e:
            logger.error(f"Failed to write {csv_path}: {e}")
            raise

    def _append_with_dup_check(self, sheet_name, df, key_cols):
        """
        Append data with duplicate checking based on key columns.
        Returns (new_count, skipped_count)
        """
        if df is None or df.empty:
            return 0, 0

        try:
            existing_df = self._read_csv(sheet_name)

            if existing_df.empty:
                self._write_csv(sheet_name, df)
                logger.info(f"Created new CSV {sheet_name} with {len(df)} rows")
                return len(df), 0

            # Find available key columns
            available_keys = [col for col in key_cols if col in existing_df.columns and col in df.columns]

            if not available_keys:
                # Use 'Date' or 'day' as fallback key
                for col in ['Date', 'day']:
                    if col in existing_df.columns and col in df.columns:
                        available_keys = [col]
                        break

            if not available_keys:
                # No key columns found, append everything (warning)
                logger.warning(f"{sheet_name}: No key columns found, appending all {len(df)} rows")
                combined_df = pd.concat([existing_df, df], ignore_index=True)
                self._write_csv(sheet_name, combined_df)
                return len(df), 0

            # Create a set of existing keys
            existing_keys = set()
            for _, row in existing_df.iterrows():
                key = tuple(str(row.get(col, '')) for col in available_keys)
                existing_keys.add(key)

            # Filter new rows
            new_rows = []
            dup_count = 0
            for _, row in df.iterrows():
                key = tuple(str(row.get(col, '')) for col in available_keys)
                if key not in existing_keys:
                    new_rows.append(row)
                else:
                    dup_count += 1

            if not new_rows:
                logger.info(f"{sheet_name}: All {len(df)} rows already exist, skipping")
                return 0, len(df)

            # Append new rows
            new_df = pd.DataFrame(new_rows)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            self._write_csv(sheet_name, combined_df)
            logger.info(f"{sheet_name}: Appended {len(new_rows)} new rows (skipped {dup_count} duplicates)")

            return len(new_rows), dup_count

        except Exception as e:
            logger.error(f"Failed to append to {sheet_name}: {e}")
            raise

        
    def update_site_row(self, row_dict):
        """
        Update or append a row to SiteSummary CSV.
        Handles mixed date formats by normalizing before comparison.
        Prevents duplicates when source data has same dates in different formats.
        """
        if not row_dict:
            return

        day = row_dict.get('day')
        if not day:
            logger.warning("Row has no 'day' field")
            return

        # Normalize the input date
        day_normalized = self._normalize_date(day)
        if not day_normalized:
            logger.error(f"Could not normalize date: {day}")
            return

        try:
            df = self._read_csv('SiteSummary')
            
            if not df.empty:
                # Normalize all existing dates for comparison
                df['_day_normalized'] = df['day'].apply(self._normalize_date)
                existing_rows = df[df['_day_normalized'] == day_normalized]
                
                if not existing_rows.empty:
                    # Row exists - UPDATE it
                    idx = existing_rows.index[0]
                    for col in SITE_SUMMARY_HEADER:
                        if col != 'day':  # Preserve original date format
                            df.at[idx, col] = row_dict.get(col, 0)
                    logger.info(f"🔄 Updated SiteSummary for {day_normalized} (format: {day})")
                else:
                    # New row - APPEND it
                    row_dict['day'] = day_normalized  # Use normalized format
                    new_row = pd.DataFrame([row_dict])
                    df = pd.concat([df, new_row], ignore_index=True)
                    logger.info(f"➕ Appended SiteSummary for {day_normalized}")
                
                # Clean up temp column
                df = df.drop('_day_normalized', axis=1)
            else:
                # First entry
                row_dict['day'] = day_normalized
                df = pd.DataFrame([row_dict])
                logger.info(f"📝 Created SiteSummary with first entry: {day_normalized}")

            self._write_csv('SiteSummary', df)
        except Exception as e:
            logger.error(f"Failed to update SiteSummary: {e}")
            raise

    def update_site_detail(self, df, target_date=None):
        """
        Merge today's site detail snapshot into the stored table.

        Site configuration (bands, sectors, RAT) rarely changes day to day,
        so this keeps exactly one row per site rather than accumulating
        history, and only advances "Last Updated" for a site when its data
        actually changed (or the site is new). Unchanged sites keep their
        existing row untouched, including their existing "Last Updated"
        date, so that column tracks the last real change per site, not the
        last time the pipeline happened to run.

        A site missing from today's snapshot (e.g. an incomplete source
        file for that day) keeps its last known row rather than being
        dropped, so a bad day's data can't silently erase history.
        """
        if df is None or df.empty:
            return

        def _norm(v):
            """Normalize a value for comparison: NaN/None -> '', and
            21.0 -> '21' so a column that picked up float dtype from an
            unrelated blank cell doesn't register as 'changed' against an
            int-typed value meaning the same thing."""
            if pd.isna(v):
                return ''
            s = str(v).strip()
            try:
                f = float(s)
                return str(int(f)) if f.is_integer() else str(f)
            except (TypeError, ValueError):
                return s

        target_date = target_date or datetime.now().strftime('%Y-%m-%d')
        compare_cols = [c for c in SITE_DETAIL_HEADER if c != 'Site Name']
        final_cols = ['Site Name'] + compare_cols + ['Last Updated']

        new_df = df.set_index('Site Name')
        existing_df = self._read_csv('SiteDetail')

        if existing_df.empty:
            out = df.copy()
            out['Last Updated'] = target_date
            self._write_csv('SiteDetail', out[final_cols])
            logger.info(f"📝 Created SiteDetail with {len(out)} sites, Last Updated={target_date}")
            return

        if 'Last Updated' not in existing_df.columns:
            # First run under this tracking scheme - no real history to
            # backfill, so start the clock now rather than fabricate a date.
            existing_df['Last Updated'] = target_date
        existing_by_site = existing_df.set_index('Site Name')

        all_sites = list(existing_by_site.index) + [s for s in new_df.index if s not in existing_by_site.index]

        new_count = updated_count = unchanged_count = kept_missing_count = 0
        rows = []
        for site in all_sites:
            if site in new_df.index:
                new_row = new_df.loc[site]
                if site in existing_by_site.index:
                    old_row = existing_by_site.loc[site]
                    changed = any(_norm(old_row.get(col)) != _norm(new_row.get(col)) for col in compare_cols)
                    if changed:
                        row = {col: new_row.get(col) for col in compare_cols}
                        row['Last Updated'] = target_date
                        updated_count += 1
                    else:
                        row = {col: old_row.get(col) for col in compare_cols}
                        row['Last Updated'] = old_row.get('Last Updated')
                        unchanged_count += 1
                else:
                    row = {col: new_row.get(col) for col in compare_cols}
                    row['Last Updated'] = target_date
                    new_count += 1
            else:
                old_row = existing_by_site.loc[site]
                row = {col: old_row.get(col) for col in compare_cols}
                row['Last Updated'] = old_row.get('Last Updated')
                kept_missing_count += 1

            row['Site Name'] = site
            rows.append(row)

        result_df = pd.DataFrame(rows, columns=final_cols)
        self._write_csv('SiteDetail', result_df)
        logger.info(
            f"🔄 SiteDetail: {new_count} new, {updated_count} changed, {unchanged_count} unchanged, "
            f"{kept_missing_count} missing-from-today (kept last known) - {len(result_df)} sites total"
        )

    def update_network_kpis(self, results_dict):
        """Update all network KPI CSVs."""
        if not results_dict:
            return

        total_new = 0
        total_skipped = 0

        for sheet_name, df in results_dict.items():
            if df is not None and not df.empty:
                key_cols = ['Date', 'Whole Network']
                new_count, skipped_count = self._append_with_dup_check(sheet_name, df, key_cols)
                total_new += new_count
                total_skipped += skipped_count

        logger.info(f"Network KPIs: {total_new} new rows, {total_skipped} skipped")
        return total_new, total_skipped

    def update_cell_kpis(self, results_dict):
        """Update all cell KPI CSVs."""
        if not results_dict:
            return

        total_new = 0
        total_skipped = 0

        for sheet_name, df in results_dict.items():
            if df is not None and not df.empty:
                key_cols = ['Date', 'Cell Name']
                new_count, skipped_count = self._append_with_dup_check(sheet_name, df, key_cols)
                total_new += new_count
                total_skipped += skipped_count
                logger.info(f"Cell KPI {sheet_name}: {new_count} new rows, {skipped_count} skipped")

        logger.info(f"Cell KPIs: {total_new} new rows, {total_skipped} skipped")
        return total_new, total_skipped

    def update_traffic_kpis(self, results_dict):
        """Update all traffic KPI CSVs."""
        if not results_dict:
            return

        total_new = 0
        total_skipped = 0

        if 'per_site' in results_dict:
            for sheet_name, df in results_dict['per_site'].items():
                if df is not None and not df.empty:
                    key_cols = ['Date', 'Site']
                    new_count, skipped_count = self._append_with_dup_check(sheet_name, df, key_cols)
                    total_new += new_count
                    total_skipped += skipped_count

        if 'whole_network' in results_dict:
            for sheet_name, df in results_dict['whole_network'].items():
                if df is not None and not df.empty:
                    network_sheet_name = sheet_name.replace('Traffic_', 'Traffic_Network_')
                    key_cols = ['Date']
                    new_count, skipped_count = self._append_with_dup_check(network_sheet_name, df, key_cols)
                    total_new += new_count
                    total_skipped += skipped_count

        logger.info(f"Traffic KPIs: {total_new} new rows, {total_skipped} skipped")
        return total_new, total_skipped

    def update_user_kpis(self, row_dict):
        """Update or append a row to UserKPIs CSV."""
        if not row_dict:
            return

        try:
            df = self._read_csv('UserKPIs')
            day = row_dict.get('Day')

            if not day:
                logger.warning("User KPIs row missing Day")
                return

            if not df.empty and day in df['Day'].values:
                idx = df[df['Day'] == day].index[0]
                for col, value in row_dict.items():
                    if col in df.columns:
                        df.at[idx, col] = value
                logger.info(f"Updated UserKPIs for {day}")
            else:
                new_row = pd.DataFrame([row_dict])
                df = pd.concat([df, new_row], ignore_index=True)
                logger.info(f"Appended UserKPIs for {day}")

            self._write_csv('UserKPIs', df)
        except Exception as e:
            logger.error(f"Failed to update UserKPIs: {e}")
            raise

    def update_packet_loss(self, df):
        """Update Packet Loss CSV."""
        if df is None or df.empty:
            return

        key_cols = ['Date', 'GBSC', 'Adjacent Node Name', 'Adjacent Node Type', 'Adjacent Node ID']
        self._append_with_dup_check('Packet_Loss', df, key_cols)
        logger.info(f"Updated Packet_Loss with {len(df)} rows")

    def update_ept_config(self, df):
        """Replace the entire EPT_Config CSV."""
        if df is None or df.empty:
            return
        self._write_csv('EPT_Config', df)
        logger.info(f"Updated EPT_Config with {len(df)} rows")
    # ---------- User KPI Methods ----------
    def update_user_kpis(self, results_dict):
        """
        Update all user KPI sheets from the results dictionary.
        """
        if not results_dict:
            logger.warning("No user KPI data to update")
            return

        self._ensure_folders()
        total_new = 0
        total_skipped = 0

        for sheet_name, df in results_dict.items():
            if df is not None and not df.empty:
                # Use Date as key column
                key_cols = ['Date']
                new_count, skipped_count = self._append_with_dup_check(sheet_name, df, key_cols)
                total_new += new_count
                total_skipped += skipped_count
                logger.info(f"User KPI {sheet_name}: {new_count} new rows, {skipped_count} skipped")

        logger.info(f"User KPIs updated: {total_new} new rows, {total_skipped} duplicates skipped")
        return total_new, total_skipped

    def update_user_summary(self, df):
        """
        Update the consolidated user summary sheet.
        """
        if df is None or df.empty:
            return

        self._ensure_folders()
        key_cols = ['Date']
        self._append_with_dup_check('User_Summary', df, key_cols)
        logger.info(f"User Summary updated with {len(df)} rows")
    # ---------- Excel Export ----------
    def export_to_excel(self, excel_name="Historical_Network_Data.xlsx"):
        """
        Combine all CSV files into a single Excel workbook.
        """
        excel_path = os.path.join(self.output_folder, excel_name)
        logger.info(f"📁 Exporting all CSVs to Excel: {excel_path}")

        try:
            # Get all CSV files
            csv_files = [f for f in os.listdir(self.csv_folder) if f.endswith('.csv')]

            if not csv_files:
                logger.warning("No CSV files found to export")
                return None

            # Create Excel writer
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                for csv_file in csv_files:
                    sheet_name = csv_file.replace('.csv', '')
                    csv_path = os.path.join(self.csv_folder, csv_file)
                    try:
                        df = pd.read_csv(csv_path)
                        if not df.empty:
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                            logger.debug(f"Exported {sheet_name}: {len(df)} rows")
                        else:
                            logger.debug(f"Skipped {sheet_name}: empty")
                    except Exception as e:
                        logger.error(f"Failed to export {csv_file}: {e}")

            logger.info(f"✅ Exported {len(csv_files)} sheets to {excel_path}")
            return excel_path

        except Exception as e:
            logger.error(f"Failed to export to Excel: {e}")
            raise