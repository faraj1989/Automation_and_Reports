#!/usr/bin/env python3
"""
Libyana NPM - SFTP Downloader
Handles FTP connection, download and file organisation with duplicate prevention.
Uses file modification time to organize files by date.
Handles multiple CSV suffix patterns dynamically.
"""

import os
import zipfile
import logging
import re
from datetime import datetime
import paramiko

logger = logging.getLogger(__name__)


class SFTPDownloader:
    """Handles SFTP connection, download and file organisation with duplicate prevention."""

    def __init__(self, host, port, username, password, remote_path, local_root, log_callback=None):
        self.host = host
        self.port = int(port)
        self.username = username
        self.password = password
        self.remote_path = remote_path
        self.local_root = local_root
        self.log_callback = log_callback
        self.ssh = None
        self.sftp = None

    def log(self, msg, level='info'):
        if self.log_callback:
            self.log_callback(msg)
        else:
            logger.info(msg)

    def connect(self):
        try:
            self.log(f"Connecting to {self.host}:{self.port} as {self.username}...")
            self.ssh = paramiko.SSHClient()
            self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            self.ssh.connect(self.host, port=self.port, username=self.username, password=self.password, timeout=30)
            self.sftp = self.ssh.open_sftp()
            self.log("SFTP connection established.")
            return True
        except Exception as e:
            self.log(f"SFTP connection failed: {e}", 'error')
            return False

    def disconnect(self):
        if self.sftp:
            self.sftp.close()
        if self.ssh:
            self.ssh.close()
        self.log("SFTP disconnected.")

    def list_remote_files(self):
        """List files in the remote directory."""
        if not self.sftp:
            self.log("SFTP session not active.", 'error')
            return []

        try:
            self.sftp.chdir(self.remote_path)
            files = self.sftp.listdir()
            self.log(f"Found {len(files)} files in remote directory.")
            return files
        except Exception as e:
            self.log(f"Failed to list remote files: {e}", 'error')
            return []

    def get_file_date_from_mtime(self, filename):
        """
        Get the date from file modification time (mtime).
        Returns date in YYYY-MM-DD format.
        """
        try:
            file_stat = self.sftp.stat(filename)
            mtime = file_stat.st_mtime
            mtime_dt = datetime.fromtimestamp(mtime)
            result = mtime_dt.strftime('%Y-%m-%d')
            return result
        except (FileNotFoundError, OSError) as e:
            if getattr(e, 'errno', None) == 2 or isinstance(e, FileNotFoundError):
                # Expected, not a fault: this filename came from a listdir()
                # snapshot taken at the start of this pass, but several
                # rolling reports (hourly all-cells, VoLTE, Gi Interface)
                # get regenerated and their old copy deleted - including by
                # our own cleanup_old_remote_duplicates() right after a
                # fresh download - before this loop reaches every entry.
                # The current copy still gets caught (either later in this
                # same pass, or the next run a few hours on), so this is
                # self-healing, not data loss - see download_and_organize's
                # caller, which counts this as skipped, not an error.
                self.log(f"ℹ️ {filename} no longer on server (likely rotated/cleaned up since the directory "
                         f"listing was taken) - will pick up its replacement instead")
            else:
                self.log(f"⚠️ Could not get mtime for {filename}: {e}")
            return None
        except Exception as e:
            self.log(f"⚠️ Could not get mtime for {filename}: {e}")
            return None

    def get_base_name_from_filename(self, filename):
        """
        Extract the base name of the file (without timestamp and extension).

        Examples:
        - "234G Cell info Daily2026_20260721121337-20260813000336.zip"
          -> "234G Cell info Daily2026"
        - "PS CS Daily Traffic per site_2G_3G_4G_20260727121322-20260813053022.zip"
          -> "PS CS Daily Traffic per site_2G_3G_4G"
        - "CS Roaming users_20260813.zip"
          -> "CS Roaming users"
        """
        # Remove extension
        name = filename.rsplit('.', 1)[0]

        # Remove a trailing timestamp suffix: _YYYYMMDDHHMMSS-YYYYMMDDHHMMSS,
        # a lone _YYYYMMDD, or anything in between - confirmed live
        # 2026-09-13 that the weekly interference report's own first half is
        # truncated to 7 digits ("_2026090-20260913052430"), not the full 14
        # every other report uses, so a rigid {14}/{8}-only match silently
        # left its timestamp attached to the "base name" forever. That made
        # every week's file a permanently-unique base name to
        # cleanup_old_remote_duplicates(), which never once fired for this
        # report - each week's ~40-57MB file just accumulated on the SFTP
        # server instead of the newest one replacing the last. 6-14 digits
        # per side comfortably covers every real timestamp width seen
        # (YYYYMM through YYYYMMDDHHMMSS) without risking a false match on
        # a report name that happens to end in a short, non-date number.
        pattern = r'(_\d{6,14}(-\d{6,14})?)$'
        name = re.sub(pattern, '', name)

        # Remove trailing underscores and spaces
        name = name.rstrip('_').strip()

        return name

    def extract_csv_suffix(self, filename):
        """
        Extract the suffix from CSV filename.
        Returns the part in parentheses after the base name.

        Examples:
        - "234G Cell info Daily2026_...(2G).csv" -> "(2G)"
        - "234G Cell info Daily2026_...(3G).csv" -> "(3G)"
        - "234G Cell info Daily2026_...(4G).csv" -> "(4G)"
        - "PS CS Daily Traffic per site_2G_3G_4G_...(PS Traffic 2G).csv" -> "(PS Traffic 2G)"
        - "PS CS Daily Traffic per site_2G_3G_4G_...(PS Traffic 3G).csv" -> "(PS Traffic 3G)"
        - "PS CS Daily Traffic per site_2G_3G_4G_...(PS Traffic 4G).csv" -> "(PS Traffic 4G)"
        - "CS Roaming users_20260813.csv" -> None (no parentheses)
        """
        # Look for pattern: (something) at the end before .csv
        match = re.search(r'\([^)]+\)\.csv$', filename)
        if match:
            # Return just the parentheses part (including parentheses)
            suffix = match.group(0)[:-4]  # Remove .csv
            return suffix
        return None

    def get_csv_identifier(self, filename):
        """
        Get a unique identifier for a CSV file (base_name + suffix).
        Used to group CSV files that belong together.

        Example:
        - Base: "234G Cell info Daily2026"
        - Suffix: "(2G)"
        - Identifier: "234G Cell info Daily2026(2G)"
        """
        suffix = self.extract_csv_suffix(filename)
        if suffix:
            # get_base_name_from_filename only strips the trailing
            # _TIMESTAMP-TIMESTAMP when it's at the very end of the name -
            # true for zip files, but every extracted CSV ends in
            # "(suffix).csv" instead, so the timestamp sits before the
            # suffix and never gets stripped. Drop "(suffix).csv" first so
            # the timestamp is back at the true end before stripping it -
            # otherwise every run's new end-timestamp makes each CSV look
            # like a brand new identifier and old copies never get cleaned.
            stem = filename[: -(len(suffix) + len('.csv'))]
            base = self.get_base_name_from_filename(stem + '.csv')
            return f"{base}{suffix}"
        # If no suffix, use full filename (without extension)
        return filename.rsplit('.', 1)[0]

    def check_file_exists_in_zipped(self, zipped_folder, base_name):
        """Check if a file with the same base name already exists in zipped folder."""
        if not os.path.exists(zipped_folder):
            return None

        for existing in os.listdir(zipped_folder):
            if existing.startswith(base_name) and existing.endswith('.zip'):
                return os.path.join(zipped_folder, existing)
        return None

    def cleanup_old_zipped_duplicates(self, zipped_folder, base_name):
        """
        Keep only the most recent zip file with the given base name.
        Delete older duplicates.
        """
        if not os.path.exists(zipped_folder):
            return

        files = []
        for f in os.listdir(zipped_folder):
            if f.startswith(base_name) and f.endswith('.zip'):
                file_path = os.path.join(zipped_folder, f)
                files.append((file_path, os.path.getmtime(file_path)))

        if len(files) <= 1:
            return

        # Sort by modification time (newest first)
        files.sort(key=lambda x: x[1], reverse=True)

        # Keep the newest, delete the rest
        for file_path, _ in files[1:]:
            try:
                os.remove(file_path)
                self.log(f"🗑️ Deleted duplicate zip: {os.path.basename(file_path)}")
            except Exception as e:
                self.log(f"⚠️ Could not delete {file_path}: {e}")

    def cleanup_old_csv_duplicates(self, unzipped_folder, base_name):
        """
        Keep only the most recent version of EACH CSV type.
        Groups CSVs by their unique identifier (base_name + suffix).

        This ensures that (2G), (3G), (4G), (PS Traffic 2G), etc.
        are all kept separately.
        """
        if not os.path.exists(unzipped_folder):
            return

        # Group CSVs by their unique identifier
        csv_groups = {}
        for f in os.listdir(unzipped_folder):
            if f.startswith(base_name) and f.endswith('.csv'):
                # Get unique identifier
                identifier = self.get_csv_identifier(f)
                if identifier:
                    if identifier not in csv_groups:
                        csv_groups[identifier] = []
                    csv_groups[identifier].append((f, os.path.getmtime(os.path.join(unzipped_folder, f))))

        # For each group, keep only the newest
        for identifier, files in csv_groups.items():
            if len(files) <= 1:
                continue
            files.sort(key=lambda x: x[1], reverse=True)
            for f, _ in files[1:]:
                try:
                    os.remove(os.path.join(unzipped_folder, f))
                    self.log(f"🗑️ Deleted duplicate CSV: {f}")
                except Exception as e:
                    self.log(f"⚠️ Could not delete {f}: {e}")

    def cleanup_old_remote_duplicates(self, base_name, keep_filename):
        """
        Delete older remote zip files sharing `base_name`, keeping only
        `keep_filename` (the one just downloaded and processed locally).
        Best-effort: a failed remote delete is logged and skipped rather
        than raised, so it can't take down the rest of the run.
        """
        try:
            remote_files = self.sftp.listdir()
        except Exception as e:
            self.log(f"⚠️ Could not list remote files for cleanup: {e}")
            return

        for fname in remote_files:
            if fname == keep_filename or not fname.lower().endswith('.zip'):
                continue
            if self.get_base_name_from_filename(fname) != base_name:
                continue
            try:
                self.sftp.remove(fname)
                self.log(f"🗑️ Deleted old remote file: {fname}")
            except Exception as e:
                self.log(f"⚠️ Could not delete remote {fname}: {e}")

    def download_and_organize(self, target_date=None):
        """
        Download and organize files from FTP.
        If target_date is specified (YYYY-MM-DD), only downloads files for that date.
        Uses file MODIFICATION TIME for folder organization.
        """
        if not self.sftp:
            self.log("SFTP session not active.", 'error')
            return False

        try:
            self.sftp.chdir(self.remote_path)
        except Exception as e:
            self.log(f"Failed to change to {self.remote_path}: {e}", 'error')
            return False

        try:
            files = self.sftp.listdir()
        except Exception as e:
            self.log(f"Failed to list remote files: {e}", 'error')
            return False

        zip_files = [f for f in files if f.lower().endswith('.zip')]
        self.log(f"📁 Found {len(zip_files)} zip files in remote directory.")

        if target_date:
            self.log(f"🎯 Target date: {target_date}")

        processed_count = 0
        skipped_count = 0
        error_count = 0

        for fname in zip_files:
            # ============================================================
            # STEP 1: Get file date from MODIFICATION TIME
            # ============================================================
            file_date = self.get_file_date_from_mtime(fname)

            if not file_date:
                # get_file_date_from_mtime already logged why (typically the
                # file was rotated/deleted server-side between listdir() and
                # here, or a transient SFTP hiccup) - counted as skipped, not
                # an error, since every observed case of this has been
                # self-healing (the replacement file gets caught later in
                # this same pass or the next run).
                skipped_count += 1
                continue

            # ============================================================
            # STEP 2: Check if target_date matches
            # ============================================================
            if target_date and file_date != target_date:
                self.log(f"⏭️ Skipping {fname} (date {file_date} != {target_date})")
                skipped_count += 1
                continue

            # ============================================================
            # STEP 3: Get base name (without timestamp)
            # ============================================================
            base_name = self.get_base_name_from_filename(fname)

            # ============================================================
            # STEP 4: Create local folders
            # ============================================================
            day_folder = os.path.join(self.local_root, file_date)
            zipped_folder = os.path.join(day_folder, 'zipped')
            unzipped_folder = os.path.join(day_folder, 'unzipped')
            os.makedirs(zipped_folder, exist_ok=True)
            os.makedirs(unzipped_folder, exist_ok=True)

            # ============================================================
            # STEP 5: Check if zip already exists in zipped folder
            # ============================================================
            existing_zip = self.check_file_exists_in_zipped(zipped_folder, base_name)

            if existing_zip:
                try:
                    remote_mtime = self.sftp.stat(fname).st_mtime
                    local_mtime = os.path.getmtime(existing_zip)

                    if remote_mtime <= local_mtime:
                        self.log(f"⏭️ Skipping {fname} - already exists in zipped (local is newer or same)")
                        skipped_count += 1
                        continue
                    else:
                        # Remote is newer, delete old and download new
                        try:
                            os.remove(existing_zip)
                            self.log(f"🗑️ Removed old zip: {os.path.basename(existing_zip)}")
                        except:
                            pass
                except Exception as e:
                    self.log(f"⚠️ Could not compare files: {e}")

            # ============================================================
            # STEP 6: Download the zip file
            # ============================================================
            local_zip_path = os.path.join(zipped_folder, fname)
            self.log(f"📥 Downloading: {fname}")
            self.log(f"   📅 Date (mtime): {file_date}")
            self.log(f"   📁 Zipped: {local_zip_path}")

            try:
                # Download the file
                self.sftp.get(fname, local_zip_path)
                self.log(f"✅ Downloaded to {local_zip_path}")

                # Clean up duplicates in zipped folder
                self.cleanup_old_zipped_duplicates(zipped_folder, base_name)

                # ============================================================
                # STEP 7: Unzip the file
                # ============================================================
                self.log(f"📂 Unzipping {local_zip_path}...")
                self._unzip_file(local_zip_path, unzipped_folder)

                # ============================================================
                # STEP 8: Clean up duplicate CSVs (keep one per technology)
                # ============================================================
                self.cleanup_old_csv_duplicates(unzipped_folder, base_name)

                # ============================================================
                # STEP 9: Now that the new file is safely downloaded and
                # unzipped, delete older remote versions of this same report
                # to keep FTP storage flat (rolling-window reports otherwise
                # accumulate one new zip per fetch forever).
                # ============================================================
                self.cleanup_old_remote_duplicates(base_name, fname)

                processed_count += 1
                self.log(f"✅ Done: {fname}")

            except Exception as e:
                self.log(f"❌ Failed to download/unzip {fname}: {e}", 'error')
                error_count += 1
                continue

        # ============================================================
        # SUMMARY
        # ============================================================
        self.log(f"\n{'=' * 50}")
        self.log(f"📊 DOWNLOAD SUMMARY")
        self.log(f"{'=' * 50}")
        self.log(f"   ✅ Downloaded: {processed_count} files")
        self.log(f"   ⏭️ Skipped: {skipped_count} files")
        self.log(f"   ❌ Errors: {error_count} files")
        self.log(f"   📁 Local root: {self.local_root}")
        self.log(f"{'=' * 50}")

        return True

    def _unzip_file(self, zip_path, extract_to):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_to)
            self.log(f"✅ Unzipped to {extract_to}")
        except Exception as e:
            self.log(f"❌ Unzip failed: {e}", 'error')
            raise