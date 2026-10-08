#!/usr/bin/env python3
"""
Libyana NPM - Daily Scheduler
Full automation pipeline: FTP download → Processing → Health → Email
"""

import os
import sys
import glob
import socket
import subprocess
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
from backend.transmission_kpi_processor import process_transmission_kpis
from backend.hourly_cell_processor import process_hourly_cell_kpis
from backend.interference_processor import process_interference_kpis, INTERFERENCE_REPORT_GLOB
from backend.traffic_kpi_processor import process_traffic_with_aggregation
from backend.user_kpi_processor import process_user_kpis, aggregate_user_data
from backend.site_detail_processor import generate_site_detail, get_latest_available_day
from backend.report_generator import ReportGenerator
from backend.history_integrity import run_and_log as run_history_integrity_check
import periodic_jobs

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

DASHBOARD_PORT = 8501
DASHBOARD_BAT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "run_dashboard.bat")
CELL_INFO_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reports", "cell_info_report.py")
PS_TRAFFIC_SCRIPT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "reports", "PS Traffic per site v3.py"
)


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
            # Step 0: Make sure the live dashboard is actually up - it's a
            # separate always-on web server, not a pipeline step, so nothing
            # else in this run depends on it, but it can silently die (closed
            # window, crash) between runs with nothing else noticing.
            logger.info("📊 Step 0: Checking dashboard is running...")
            self._ensure_dashboard_running()

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

            logger.info("📡 Step 5: Processing Transmission KPIs...")
            self._process_transmission_kpis()

            logger.info("📊 Step 6: Processing Traffic KPIs...")
            self._process_traffic_kpis()

            logger.info("👥 Step 7: Processing User KPIs...")
            self._process_user_kpis()

            logger.info("📋 Step 8: Generating Site Detail...")
            self._process_site_detail()

            logger.info("🚨 Step 8.5: Archiving NOC Daily Alarm Summary...")
            self._archive_noc_daily_alarm_summary(target_date)

            # Step 8.6: Validate history holds only complete business days
            # (no today/future rows, no duplicate keys, no recent gaps) -
            # report-only, never aborts the run
            logger.info("🔎 Step 8.6: Checking history integrity...")
            run_history_integrity_check()

            # Step 9: Export to Excel
            logger.info("💾 Step 9: Exporting to Excel...")
            excel_path = self.history_mgr.export_to_excel()

            # Step 10: Generate Report (health scoring happens inside report_gen,
            # computed strictly from target_date's rows in output/csv/)
            logger.info("📧 Step 10: Generating Report...")
            report_text, report_excel, report_word = self.report_gen.generate_report(target_date)

            # Step 11: Email (if auto_send)
            if auto_send:
                logger.info("📧 Step 11: Sending Email Report...")
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

    def _process_transmission_kpis(self):
        """Step 5: Process Transmission KPIs (IUB/ABIS packet loss + latency)"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            results = process_transmission_kpis(day_folder, log_callback=logger.info)
            if results:
                self.history_mgr.update_transmission_kpis(results)

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
                    self.history_mgr.update_user_summary(
                        summary_df, source=f"{os.path.relpath(day_folder, local_root)} (aggregated user reports)")

    def _process_site_detail(self):
        """Step 7: Generate Site Detail"""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            df = generate_site_detail(day_folder, log_callback=logger.info)
            if df is not None and not df.empty:
                self.history_mgr.update_site_detail(df, target_date=self.yesterday)

    def _archive_noc_daily_alarm_summary(self, target_date):
        """Step 8.5: Archive that day's down-site NOC alarm analysis (see
        backend/noc_alarm_processor.build_daily_noc_alarm_report and
        CSVHistoryManager.update_noc_daily_alarm_summary) into
        output/csv/NOC_Daily_Alarm_Summary.csv. Read-only against the
        sibling NOC Automation Suite's raw historical exports, so a missing/
        unavailable feed just skips this step rather than failing the
        pipeline - the rest of the daily report doesn't depend on it."""
        try:
            alarm_report = self.report_gen.build_daily_noc_alarm_report(target_date)
            if not alarm_report.get('available'):
                logger.info(f"NOC alarm feed not available for {target_date}, skipping archive")
                return
            down_summary = alarm_report.get('down_sites_summary')
            if down_summary is None or down_summary.empty:
                logger.info(f"No down sites recorded for {target_date}")
                return
            self.history_mgr.update_noc_daily_alarm_summary(down_summary)
        except Exception as e:
            logger.warning(f"NOC alarm summary archival failed (non-fatal): {e}")

    def _send_email(self, report_text, report_excel, report_word):
        """Step 11: Send email report"""
        # This would use SMTP - placeholder for now
        logger.info("📧 Email sending not implemented yet")
        logger.info(f"Report text: {report_text}")
        logger.info(f"Report Excel: {report_excel}")
        logger.info(f"Report Word: {report_word}")

    def _is_dashboard_running(self, port=DASHBOARD_PORT, timeout=1.5):
        """True if something is already listening on the dashboard's port
        (the dashboard is a persistent web server, not a pipeline step -
        this just checks it's alive, it doesn't validate it's actually
        streamlit_dashboard.py answering)."""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(timeout)
            return s.connect_ex(('127.0.0.1', port)) == 0

    def _ensure_dashboard_running(self):
        """Self-healing check: the dashboard should always be up, but it's
        an independent long-running process (not something this script
        starts and waits on) that can silently die - closed window, crash,
        machine woke from sleep without it - between runs with nothing
        else noticing. Relaunches run_dashboard.bat in its own detached
        console window (same as double-clicking it) if it's not listening."""
        if self._is_dashboard_running():
            logger.info(f"   ✅ Dashboard already running on port {DASHBOARD_PORT}")
            return

        if not os.path.exists(DASHBOARD_BAT):
            logger.warning(f"   ⚠️ Dashboard not running and {DASHBOARD_BAT} not found - can't auto-start")
            return

        logger.warning(f"   ⚠️ Dashboard not running on port {DASHBOARD_PORT} - restarting it")
        try:
            subprocess.Popen(
                ['cmd', '/c', 'start', 'Libyana Dashboard', DASHBOARD_BAT],
                cwd=os.path.dirname(DASHBOARD_BAT),
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
            logger.info("   ✅ Dashboard relaunch triggered")
        except Exception as e:
            logger.error(f"   ❌ Failed to relaunch dashboard: {e}")

    # ------------------------------------------------------------------
    # Hourly cells update - lightweight, independently-schedulable path.
    # Separate from run() because scheduler.py has no internal interval
    # loop (it's a one-shot CLI triggered once/day externally); re-running
    # the whole daily pipeline every ~6 hours would redundantly re-score
    # and re-email. This is meant to be pointed at by its own scheduled
    # task, every ~6 hours, independent of the once-daily run().
    # ------------------------------------------------------------------

    def _download_latest_ftp_files(self):
        """Fetch whatever's newest on the FTP server right now, no date
        filter - unlike _download_ftp's "yesterday's finalized report"
        day-offset logic, the hourly all-cells report is a live rolling
        window, so there's no target day to offset against."""
        host = self.config.get('host')
        port = self.config.get('port')
        username = self.config.get('username')
        password = self.config.get('password')
        remote_path = self.config.get('remote_path')
        local_root = self.config.get('local_root')

        downloader = SFTPDownloader(
            host, port, username, password,
            remote_path, local_root,
            log_callback=logger.info
        )
        try:
            if downloader.connect():
                result = downloader.download_and_organize()
                downloader.disconnect()
                return result
            return False
        except Exception as e:
            logger.error(f"FTP download failed: {e}")
            return False

    def _process_hourly_cells(self):
        """Process the hourly all-cells report (2G/3G/4G live cell KPIs)."""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            results = process_hourly_cell_kpis(day_folder, log_callback=logger.info)
            if results:
                self.history_mgr.update_hourly_cell_kpis(results)

    def run_hourly_cells_update(self):
        """Entry point for the ~6-hourly scheduled task: fetch whatever's
        new on the FTP and archive the hourly all-cells report."""
        logger.info("=" * 70)
        logger.info("📶 STARTING HOURLY CELLS UPDATE")
        logger.info("=" * 70)
        start_time = time.time()

        try:
            self._ensure_dashboard_running()
            self._download_latest_ftp_files()
            self._process_hourly_cells()

            elapsed = time.time() - start_time
            logger.info("=" * 70)
            logger.info(f"✅ HOURLY CELLS UPDATE COMPLETED - {elapsed:.1f} seconds")
            logger.info("=" * 70)
            return True
        except Exception as e:
            logger.error(f"❌ Hourly cells update failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    # ------------------------------------------------------------------
    # Weekly external-interference update - dedicated report Huawei
    # iMaster uploads daily (~04:00, ready ~06:00), but 3G/4G are big
    # enough (millions of hourly rows) that we only pull it once a week;
    # 2G is already daily cell-level granularity so nothing is lost by
    # not fetching more often on that side. Meant to be pointed at by its
    # own scheduled task, Sunday mornings, independent of the daily and
    # ~6-hourly jobs above.
    # ------------------------------------------------------------------

    INTERFERENCE_REPORT_BASE_NAME = '2G_3G_4G interference and PRB utilization for Automation'

    def _cleanup_local_interference_files(self):
        """Delete the raw interference+PRB zip/CSVs from every dated
        local folder once they're archived into output/csv/ - a single
        week's zip is ~90MB and unzips to ~650MB, the biggest single file
        this pipeline handles, and nothing needs the local copy once the
        data lives in the persistent archive."""
        local_root = self.config.get('local_root')
        if not local_root or not os.path.exists(local_root):
            return

        removed_bytes = 0
        for entry in os.listdir(local_root):
            day_folder = os.path.join(local_root, entry)
            if not os.path.isdir(day_folder):
                continue

            zipped_folder = os.path.join(day_folder, 'zipped')
            if os.path.exists(zipped_folder):
                for f in os.listdir(zipped_folder):
                    if f.lower().endswith('.zip') and f.startswith(self.INTERFERENCE_REPORT_BASE_NAME):
                        path = os.path.join(zipped_folder, f)
                        try:
                            removed_bytes += os.path.getsize(path)
                            os.remove(path)
                            logger.info(f"🗑️ Removed archived interference zip: {path}")
                        except Exception as e:
                            logger.warning(f"Could not remove {path}: {e}")

            unzipped_folder = os.path.join(day_folder, 'unzipped')
            if os.path.exists(unzipped_folder):
                for f in glob.glob(os.path.join(unzipped_folder, INTERFERENCE_REPORT_GLOB)):
                    try:
                        removed_bytes += os.path.getsize(f)
                        os.remove(f)
                        logger.info(f"🗑️ Removed archived interference CSV: {f}")
                    except Exception as e:
                        logger.warning(f"Could not remove {f}: {e}")

        if removed_bytes:
            logger.info(f"🧹 Freed {removed_bytes / (1024 * 1024):.1f} MB of local interference report files")

    def _process_interference_kpis(self):
        """Archive the dedicated external-interference report (2G/3G/4G
        cell-level interference, plus bonus 4G PRB/throughput columns)."""
        local_root = self.config.get('local_root')
        day_folder = get_latest_day_folder(local_root)

        if day_folder:
            results = process_interference_kpis(day_folder, log_callback=logger.info)
            if results:
                self.history_mgr.update_interference_kpis(results)

    def run_interference_weekly_update(self):
        """Entry point for the weekly (Sunday morning) scheduled task:
        fetch whatever's newest on the FTP (same rolling-window fetch as
        the hourly-cells job), archive the interference report with full
        history retained, then free the local disk space the raw report
        used."""
        logger.info("=" * 70)
        logger.info("📡 STARTING WEEKLY INTERFERENCE UPDATE")
        logger.info("=" * 70)
        start_time = time.time()

        try:
            self._ensure_dashboard_running()
            self._download_latest_ftp_files()
            self._process_interference_kpis()
            self._cleanup_local_interference_files()

            elapsed = time.time() - start_time
            logger.info("=" * 70)
            logger.info(f"✅ WEEKLY INTERFERENCE UPDATE COMPLETED - {elapsed:.1f} seconds")
            logger.info("=" * 70)
            return True
        except Exception as e:
            logger.error(f"❌ Weekly interference update failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    # ------------------------------------------------------------------
    # Cell Info (monthly) and PS Traffic per site (weekly) - both ported
    # standalone scripts under reports/ (see reports/cell_info_report.py,
    # reports/"PS Traffic per site v3.py"), each a top-level script rather
    # than an importable function, so they're run as a subprocess instead
    # of imported - importing PS Traffic per site v3.py in-process would
    # also have its own parse_args() read *this* process's sys.argv and
    # collide with scheduler.py's own CLI flags.
    # ------------------------------------------------------------------

    def _run_report_script(self, script_path, label):
        """subprocess.run's text=True/encoding="utf-8" only govern how *this*
        process decodes the child's output - the child still encodes its own
        emoji-laden print()s using its own default console codepage (cp1252
        on Windows) unless told otherwise, same UnicodeEncodeError class this
        script's own stdout/stderr .reconfigure() calls above guard against.
        PYTHONIOENCODING/PYTHONUTF8 force the child to use UTF-8 instead -
        same fix scrapers/scraper_watchdog.py already applies to its own
        subprocess launches."""
        logger.info(f"   ▶️ Running {label} ({script_path})...")
        child_env = os.environ.copy()
        child_env["PYTHONUTF8"] = "1"
        child_env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            [sys.executable, script_path],
            cwd=os.path.dirname(script_path),
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            env=child_env,
        )
        for line in result.stdout.splitlines():
            logger.info(f"   [{label}] {line}")
        if result.returncode != 0:
            for line in result.stderr.splitlines():
                logger.error(f"   [{label}] {line}")
            logger.error(f"   ❌ {label} exited with code {result.returncode}")
            return False
        return True

    def run_cell_info_monthly_update(self):
        """Entry point for the monthly (1st-of-month) scheduled task: pulls
        this month's 2G/3G/4G cell inventory + EPT from the FTPS server on
        port 21 (a separate account/protocol from the SFTP-22 creds used by
        every other job here - see reports/cell_info_report.py) and rebuilds
        the merged workbook the dashboard's Monthly Cell Info Report tab
        reads."""
        logger.info("=" * 70)
        logger.info("📶 STARTING MONTHLY CELL INFO UPDATE")
        logger.info("=" * 70)
        start_time = time.time()
        try:
            self._ensure_dashboard_running()
            ok = self._run_report_script(CELL_INFO_SCRIPT, "cell_info_report")
            elapsed = time.time() - start_time
            if ok:
                logger.info(f"✅ MONTHLY CELL INFO UPDATE COMPLETED - {elapsed:.1f} seconds")
            return ok
        except Exception as e:
            logger.error(f"❌ Monthly cell info update failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def run_ps_traffic_weekly_update(self):
        """Entry point for the weekly (Sunday morning) scheduled task -
        also callable on demand from the dashboard's "Refresh now" button
        (see streamlit_dashboard.py). Runs reports/"PS Traffic per site
        v3.py" against whatever's currently in PS_TRAFFIC_SOURCE_DIR."""
        logger.info("=" * 70)
        logger.info("📶 STARTING WEEKLY PS TRAFFIC PER SITE UPDATE")
        logger.info("=" * 70)
        start_time = time.time()
        try:
            self._ensure_dashboard_running()
            ok = self._run_report_script(PS_TRAFFIC_SCRIPT, "PS Traffic per site v3")
            elapsed = time.time() - start_time
            if ok:
                logger.info(f"✅ WEEKLY PS TRAFFIC UPDATE COMPLETED - {elapsed:.1f} seconds")
            return ok
        except Exception as e:
            logger.error(f"❌ Weekly PS traffic update failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def run_cem_monthly_report(self, month=None):
        """Entry point for the monthly (2nd-of-month) scheduled task: builds
        the previous month's SmartCare CEM Comprehensive Analysis workbook for
        Tripoli HQ (3 sheets, HQ's format) into Monthly_CEM_Reports. Returns
        False while the month still has missing days, so the watchdog
        backstop (periodic_jobs.py) retries it after the daily SmartCare runs
        have filled them - the file is still written each time, so the most
        complete version is always on disk."""
        from backend import smartcare_cem_processor as smartcare_cem
        if month is None:
            first_this_month = datetime.now().replace(day=1)
            month = (first_this_month - timedelta(days=1)).strftime('%Y-%m')
        logger.info("=" * 70)
        logger.info(f"📊 BUILDING MONTHLY CEM COMPREHENSIVE ANALYSIS - {month}")
        logger.info("=" * 70)
        try:
            path, info = smartcare_cem.save_monthly_cem_report(month)
            logger.info(f"   💾 {path}")
            logger.info(f"   📅 {len(info['present'])}/{info['days_in_month']} days")
            if info['missing']:
                logger.warning(f"   ⚠️ Missing day(s): {', '.join(info['missing'])} - will retry")
                return False
            logger.info("✅ MONTHLY CEM REPORT COMPLETE")
            return True
        except Exception as e:
            logger.error(f"❌ Monthly CEM report failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False


def run_periodic_job(job, fn):
    """Run one weekly/monthly job under its lock and record a success stamp
    (see periodic_jobs.py - the watchdog re-runs any job that has none for
    its latest due time). Returns the process exit code, so a failed run
    shows as failed in Task Scheduler instead of a silent 0."""
    with periodic_jobs.job_lock(job) as acquired:
        if not acquired:
            logger.warning(f"⏭️ {job} is already running (pid {periodic_jobs.lock_holder(job)}) - skipping this launch")
            return 0
        ok = fn()
        if ok:
            periodic_jobs.mark_success(job)
        return 0 if ok else 1


def main():
    parser = argparse.ArgumentParser(description='Libyana NPM Daily Scheduler')
    parser.add_argument('--date', help='Target date (YYYY-MM-DD)', default=None)
    parser.add_argument('--auto-send', action='store_true', help='Auto-send email')
    parser.add_argument('--hourly-cells', action='store_true',
                         help='Run only the ~6-hourly live cell KPI update, not the full daily pipeline')
    parser.add_argument('--interference-weekly', action='store_true',
                         help='Run only the weekly external-interference report update, not the full daily pipeline')
    parser.add_argument('--cell-info-monthly', action='store_true',
                         help='Run only the monthly Cell Info FTPS pull (reports/cell_info_report.py), '
                              'not the full daily pipeline')
    parser.add_argument('--ps-traffic-weekly', action='store_true',
                         help='Run only the weekly PS Traffic per site report '
                              '(reports/"PS Traffic per site v3.py"), not the full daily pipeline')
    parser.add_argument('--cem-monthly', nargs='?', const='', metavar='YYYY-MM',
                         help="Build the monthly SmartCare CEM Comprehensive Analysis workbook for Tripoli HQ "
                              "(default: previous month), not the full daily pipeline")
    args = parser.parse_args()

    scheduler = DailyScheduler()

    if args.hourly_cells:
        scheduler.run_hourly_cells_update()
        return

    if args.interference_weekly:
        sys.exit(run_periodic_job("interference-weekly", scheduler.run_interference_weekly_update))

    if args.cell_info_monthly:
        sys.exit(run_periodic_job("cell-info-monthly", scheduler.run_cell_info_monthly_update))

    if args.ps_traffic_weekly:
        sys.exit(run_periodic_job("ps-traffic-weekly", scheduler.run_ps_traffic_weekly_update))

    if args.cem_monthly is not None:
        if args.cem_monthly:  # an explicit month: a manual rebuild, not the scheduled job
            sys.exit(0 if scheduler.run_cem_monthly_report(args.cem_monthly) else 1)
        sys.exit(run_periodic_job("cem-monthly", scheduler.run_cem_monthly_report))

    if args.date:
        target_date = args.date
    else:
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    scheduler.run(target_date=target_date, auto_send=args.auto_send)


if __name__ == "__main__":
    main()
