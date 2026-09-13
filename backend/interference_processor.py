#!/usr/bin/env python3
"""
Libyana NPM - External Interference Processor
Processes the dedicated "2G_3G_4G interference and PRB utilization for
Automation" report (uploaded by Huawei iMaster daily at ~04:00, ready to
fetch by ~06:00): cell-level external interference for 2G (one row per
cell per day already) and 3G/4G (one row per cell per hour), plus bonus
4G PRB/throughput columns.

Same fix as hourly_cell_processor.py: the 3 per-technology CSVs inside
the report's zip have non-breaking-space suffixes in their filenames
("(2G\xa0Interference).csv" etc.), so files are identified by a
distinctive metric column in their header instead of by filename.
"""

import os
import glob
import logging

import pandas as pd

from backend.csv_loader import read_csv_skip_metadata

logger = logging.getLogger(__name__)

INTERFERENCE_REPORT_GLOB = '*interference*PRB*utilization*Automation*.csv'

# A column that uniquely identifies which technology's file this is.
TECH_IDENTITY_COLUMN = {
    'Interference Band Proportion (4~5)(%)': '2G_Interference',
    'VS.MeanRTWP': '3G_Interference_Hourly',
    'L.UL.Interference.Avg(dBm)': '4G_Interference_Hourly',
}

# 3G DL frequency -> band, and 4G "Frequency band" (an index, not the raw
# EARFCN channel) -> MHz - duplicated from special_reports_processor.py's
# _3G_BAND_MAP/_4G_BAND_MAP (not imported, to avoid a circular import: that
# module imports build_nq_template_report-adjacent helpers that don't need
# this rollup logic, and this module is imported by scheduler.py before
# special_reports_processor.py's own dependencies are needed).
_3G_BAND_MAP = {3054: 900, 3062: 900, 3075: 900, 10562: 2100, 10587: 2100}
_4G_BAND_MAP = {1: 2100, 3: 1800, 8: 900, 28: 700}

# Per-tech spec for the daily interfered-hours rollup: which raw archive to
# read, which small rollup sheet to write, and the same metric/threshold/
# band mapping build_external_interference_report uses to decide "was this
# hour interfered".
ROLLUP_SPECS = {
    '3G': {
        'raw_sheet': '3G_Interference_Hourly', 'rollup_sheet': '3G_Interference_Daily',
        'metric_col': 'VS.MeanRTWP', 'threshold': -95,
        'band_col': 'DL frequency', 'band_map': _3G_BAND_MAP,
    },
    '4G': {
        'raw_sheet': '4G_Interference_Hourly', 'rollup_sheet': '4G_Interference_Daily',
        'metric_col': 'L.UL.Interference.Avg(dBm)', 'threshold': -100,
        'band_col': 'Frequency band', 'band_map': _4G_BAND_MAP,
    },
}


def _classify_sheet(columns) -> str:
    for col, sheet_name in TECH_IDENTITY_COLUMN.items():
        if col in columns:
            return sheet_name
    return None


def aggregate_interfered_hours(df, spec) -> pd.DataFrame:
    """One row per (Cell Name, Band, Date) that had at least one hourly
    reading in `df`: BadHours (count of hours where the metric breached
    spec['threshold']) and TotalHoursReported. This is exactly the
    intermediate build_external_interference_report used to compute
    inline from the full raw hourly file - factored out so it can run
    once per chunk (see build_daily_rollup) or once per weekly batch
    (see update_daily_rollup), instead of the report re-deriving it from
    scratch out of a multi-hundred-MB file on every view."""
    if df is None or df.empty:
        return pd.DataFrame(columns=['Cell Name', 'Band', 'Date', 'BadHours', 'TotalHoursReported'])

    band_col, metric_col = spec['band_col'], spec['metric_col']
    if band_col not in df.columns or metric_col not in df.columns:
        return pd.DataFrame(columns=['Cell Name', 'Band', 'Date', 'BadHours', 'TotalHoursReported'])

    work = df[['Time', 'Cell Name', band_col, metric_col]].copy()
    work['Date'] = pd.to_datetime(work['Time'], errors='coerce').dt.floor('D')
    work = work.dropna(subset=['Date'])
    work['_metric'] = pd.to_numeric(work[metric_col], errors='coerce')
    work['Band'] = pd.to_numeric(work[band_col], errors='coerce').map(spec['band_map'])
    work = work.dropna(subset=['Band'])
    work['_bad'] = work['_metric'] > spec['threshold']

    grouped = work.groupby(['Cell Name', 'Band', 'Date'])['_bad'].agg(
        BadHours='sum', TotalHoursReported='count'
    ).reset_index()
    return grouped


def build_daily_rollup(csv_folder, tech, chunksize=500_000, log_callback=None):
    """One-time (re)build of the small BadHours/day rollup from the FULL
    raw hourly archive, reading it in chunks so peak memory stays bounded
    to one chunk's worth of rows regardless of the raw file's total size -
    the raw 3G/4G interference archives are multi-hundred-MB (unbounded,
    full history kept by design) and a plain pd.read_csv() on them has
    been observed to raise MemoryError on a 16GB machine already under
    normal desktop load. The accumulator's size is bounded by
    cells x bands x days, not by raw row count, so this stays small
    (tens of thousands of keys) even over months of hourly data."""

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            logger.info(msg)

    spec = ROLLUP_SPECS[tech]
    path = os.path.join(csv_folder, f"{spec['raw_sheet']}.csv")
    if not os.path.exists(path):
        log(f"⚠️ {path} not found, skipping rollup build")
        return pd.DataFrame()

    # This machine's free memory has been observed to swing by several GB
    # within minutes (other desktop apps, Windows memory compression), so
    # even a chunked read can hit a bad moment - halve chunksize and retry
    # from scratch rather than failing outright on a transient dip.
    min_chunksize = 10_000
    while True:
        try:
            accum = {}  # (Cell Name, Band, Date) -> [BadHours, TotalHoursReported]
            chunks_seen = 0
            for chunk in pd.read_csv(path, usecols=['Time', 'Cell Name', spec['band_col'], spec['metric_col']],
                                      chunksize=chunksize):
                chunks_seen += 1
                part = aggregate_interfered_hours(chunk, spec)
                for row in part.itertuples(index=False):
                    key = (row[0], row[1], row[2])  # (Cell Name, Band, Date)
                    bad, total = accum.get(key, (0, 0))
                    accum[key] = (bad + row.BadHours, total + row.TotalHoursReported)
                log(f"   chunk {chunks_seen}: {len(chunk):,} raw rows -> {len(accum):,} cumulative (Cell,Band,Date) keys")
            break
        except (MemoryError, pd.errors.ParserError) as e:
            if chunksize <= min_chunksize:
                raise
            chunksize = max(chunksize // 4, min_chunksize)
            log(f"⚠️ Ran out of memory mid-read ({e}) - retrying with chunksize={chunksize:,}")

    if not accum:
        return pd.DataFrame()

    rollup = pd.DataFrame(
        [{'Cell Name': k[0], 'Band': k[1], 'Date': k[2], 'BadHours': v[0], 'TotalHoursReported': v[1]}
         for k, v in accum.items()]
    )
    log(f"✅ Built {tech} rollup: {len(rollup):,} (Cell,Band,Date) rows from {chunks_seen} chunk(s)")
    return rollup


def update_daily_rollup(batch_df, existing_rollup, tech) -> pd.DataFrame:
    """Incremental version of build_daily_rollup for the weekly archival
    job: aggregates only the just-fetched batch (already in memory, at
    most ~2.4M rows per Huawei's own export cap - no chunking needed) and
    merges it into the existing small rollup, replacing any (Cell, Band,
    Date) key the batch also covers. Safe against the batch's date range
    overlapping the previous rollup (every weekly pull re-sends 2-3 weeks
    of trailing data) because raw archival already unions all hourly rows
    ever seen for a given date before this runs, so the batch's count for
    an overlapping date is always the complete one - never a partial
    double-count."""
    spec = ROLLUP_SPECS[tech]
    new_part = aggregate_interfered_hours(batch_df, spec)
    if new_part.empty:
        return existing_rollup if existing_rollup is not None else pd.DataFrame()

    if existing_rollup is None or existing_rollup.empty:
        return new_part

    # existing_rollup comes back from a plain CSV read (CSVHistoryManager.
    # _read_csv), so its Date column is still a string ("2026-09-01");
    # new_part's Date is a real Timestamp (from aggregate_interfered_hours'
    # .dt.floor('D')). Left unnormalized, the same calendar day compares as
    # two different values in drop_duplicates below and BOTH rows survive -
    # the stale existing row never gets replaced, and it silently keeps
    # accumulating duplicate/stale (Cell, Band, Date) rows forever.
    existing_rollup = existing_rollup.copy()
    existing_rollup['Date'] = pd.to_datetime(existing_rollup['Date'], errors='coerce')

    combined = pd.concat([existing_rollup, new_part], ignore_index=True)
    return combined.drop_duplicates(subset=['Cell Name', 'Band', 'Date'], keep='last').reset_index(drop=True)


def process_interference_kpis(day_folder, log_callback=None):
    """
    Find and process the 3 per-technology files in the dedicated
    interference+PRB report. Returns a dictionary: sheet_name -> DataFrame,
    at full native grain (2G: one row/cell/day; 3G/4G: one row/cell/hour -
    no aggregation here, that happens in build_external_interference_report).
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("=" * 60)
    log("📡 PROCESSING EXTERNAL INTERFERENCE REPORT (2G/3G/4G)")
    log("=" * 60)

    results = {}

    candidates = glob.glob(os.path.join(day_folder, INTERFERENCE_REPORT_GLOB))
    if not candidates:
        log("   ⚠️ No interference+PRB report files found")
        return results

    for file_path in candidates:
        log(f"📄 Reading: {os.path.basename(file_path)}")
        try:
            df = read_csv_skip_metadata(file_path)
            if df is None or df.empty:
                log("   ⚠️ File is empty or could not be read")
                continue

            sheet_name = _classify_sheet(df.columns)
            if sheet_name is None:
                log("   ⚠️ Could not classify technology (no recognized identity column), skipping")
                continue

            rows_before = len(df)
            df = df.dropna(how='all')
            empty_rows_removed = rows_before - len(df)
            if empty_rows_removed > 0:
                log(f"   🗑️  Removed {empty_rows_removed} completely empty rows")

            results[sheet_name] = df
            log(f"   ✅ Classified as {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
        except Exception as e:
            log(f"   ❌ Error reading {os.path.basename(file_path)}: {e}")

    log("=" * 60)
    return results


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 2:
        print("Usage: python interference_processor.py <day_folder_path>")
        sys.exit(1)

    test_folder = sys.argv[1]
    results = process_interference_kpis(test_folder)

    print("\n" + "=" * 60)
    print("EXTERNAL INTERFERENCE PROCESSING RESULTS")
    print("=" * 60)

    for sheet_name, df in results.items():
        print(f"\n{sheet_name}:")
        if df is not None and not df.empty:
            print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")
            print(f"  Sample:\n{df.head(3)}")
        else:
            print("  No data")
