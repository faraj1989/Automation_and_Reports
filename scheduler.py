#!/usr/bin/env python3
"""
Libyana NPM - Daily Scheduler
Full automation pipeline: FTP download → Processing → Health → Email
"""

import os
import sys
import logging
import time
from datetime import datetime, timedelta
import argparse

# Add backend to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from backend import (
    ConfigManager,
    SFTPDownloader,
    process_site_day,
    process_all_days,
    get_latest_day_folder,
    SITE_SUMMARY_HEADER
)
from backend.csv_history_manager import CSVHistoryManager
from backend.network_kpi_processor import process_network_kpis
from backend.cell_kpi_processor import process_cell_kpis
from backend.traffic_kpi_processor import process_traffic_with_aggregation
from backend.user_kpi_processor import process_user_kpis, aggregate_user_data
from backend.site_detail_processor import generate_site_detail, get_latest_available_day
from backend.report_generator import ReportGenerator

# Setup logging
LOG_FILE = "scheduler.log"

# Windows PowerShell may expose stdout with the legacy cp1252 encoding.  Several
# pipeline components log Unicode symbols, so use UTF-8 for both the console and
# the persistent log to prevent logging failures after an otherwise successful
# run.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class DailyScheduler:
    def __init__(self):
        self.config = ConfigManager()
        self.history_mgr = CSVHistoryManager()
        self.report_gen = ReportGenerator()
        self.today = datetime.now().strftime('%Y-%m-%d')
        self.yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    def run(self, target_date=None, auto_send=False):
        """Run the complete daily pipeline"""
        if target_date is None:
            target_date = self.yesterday

        logger.info("=" * 70)
        logger.info(f"🚀 STARTING DAILY PIPELINE - {target_date}")
        logger.info("=" * 70)
        start_time = time.time()

        try:
            # Step 1: FTP Download
            logger.info("📥 Step 1: Downloading from FTP...")
            if not self._download_ftp(target_date):
                logger.error("❌ FTP download failed, aborting")
                return False

            # Step 2-7: Process all KPIs
            logger.info("📊 Step 2: Processing Site Summary...")
            self._process_site_summary()

            logger.info("📈 Step 3: Processing Network KPIs...")
            self._process_network_kpis()

            logger.info("📊 Step 4: Processing Cell KPIs...")
            self._process_cell_kpis()

            logger.info("📊 Step 5: Processing Traffic KPIs...")
            self._process_traffic_kpis()

            logger.info("👥 Step 6: Processing User KPIs...")
            self._process_user_kpis()

            logger.info("📋 Step 7: Generating Site Detail...")
            self._process_site_detail()

            # Step 8: Export to Excel
            logger.info("💾 Step 8: Exporting to Excel...")
            excel_path = self.history_mgr.export_to_excel()

            # Step 9: Generate Report (health scoring happens inside report_gen,
            # computed strictly from target_date's rows in output/csv/)
            logger.info("📧 Step 9: Generating Report...")
            report_text, report_excel, report_word = self.report_gen.generate_report(target_date)

            # Step 10: Email (if auto_send)
            if auto_send:
                logger.info("📧 Step 10: Sending Email Report...")
                self._send_email(report_text, report_excel, report_word)

            elapsed = time.time() - start_time
            logger.info("=" * 70)
            logger.info(f"✅ PIPELINE COMPLETED SUCCESSFULLY - {elapsed:.1f} seconds")
            logger.info("=" * 70)
            return True

        except Exception as e:
            logger.error(f"❌ Pipeline failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def _download_ftp(self, target_date):
        """Step 1: Download files from FTP

        Huawei MAE uploads each report at ~05:30 the morning AFTER the day it
        covers: the file with mtime 2026-08-18 contains the finalized KPIs for
        2026-08-17. So to get complete data for target_date, we must fetch the
        file dated target_date + 1 day, not target_date itself.
        """
        host = self.config.get('host')
        port = self.config.get('port')
        username = self.config.get('username')
        password = self.config.get('password')
        remote_path = self.config.get('remote_path')
        local_root = self.config.get('local_root')

        target_dt = datetime.strptime(target_date, '%Y-%m-%d')
        ftp_date = (target_dt + timedelta(days=1)).strftime('%Y-%m-%d')

        downloader = SFTPDownloader(
            host, port, username, password,
            remote_path, local_root,
            log_callback=logger.info
        )

        try:
            if downloader.connect():
                result = downloader.download_and_organize(target_date=ftp_date)
                downloader.disconnect()
                return result
            return False
        except Exception as e:
            logger.error(f"FTP download failed: {e}")
            return False

    def _process_site_summary(self):
        """Step 2: Process Site Summary"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            result = process_site_day(day_folder, log_callback=logger.info)
            if result:
                result['day'] = self.yesterday
                self.history_mgr.update_site_row(result)

    def _process_network_kpis(self):
        """Step 3: Process Network KPIs"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            results = process_network_kpis(day_folder, log_callback=logger.info)
            if results:
                self.history_mgr.update_network_kpis(results)

    def _process_cell_kpis(self):
        """Step 4: Process Cell KPIs"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            results = process_cell_kpis(day_folder, log_callback=logger.info)
            if results:
                self.history_mgr.update_cell_kpis(results)

    def _process_traffic_kpis(self):
        """Step 5: Process Traffic KPIs"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            results = process_traffic_with_aggregation(day_folder, log_callback=logger.info)
            if results:
                self.history_mgr.update_traffic_kpis(results)

    def _process_user_kpis(self):
        """Step 6: Process User KPIs"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            raw_results = process_user_kpis(day_folder, log_callback=logger.info)
            if raw_results:
                self.history_mgr.update_user_kpis(raw_results)
                summary_df = aggregate_user_data(raw_results)
                if summary_df is not None:
                    self.history_mgr.update_user_summary(summary_df)

    def _process_site_detail(self):
        """Step 7: Generate Site Detail"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            df = generate_site_detail(day_folder, log_callback=logger.info)
            if df is not None and not df.empty:
                self.history_mgr.update_site_detail(df, target_date=self.yesterday)

    def _send_email(self, report_text, report_excel, report_word):
        """Step 11: Send email report"""
        # This would use SMTP - placeholder for now
        logger.info("📧 Email sending not implemented yet")
        logger.info(f"Report text: {report_text}")
        logger.info(f"Report Excel: {report_excel}")
        logger.info(f"Report Word: {report_word}")


def main():
    parser = argparse.ArgumentParser(description='Libyana NPM Daily Scheduler')
    parser.add_argument('--date', help='Target date (YYYY-MM-DD)', default=None)
    parser.add_argument('--auto-send', action='store_true', help='Auto-send email')
    args = parser.parse_args()

    scheduler = DailyScheduler()

    if args.date:
        target_date = args.date
    else:
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    scheduler.run(target_date=target_date, auto_send=args.auto_send)


if __name__ == "__main__":
    main()
