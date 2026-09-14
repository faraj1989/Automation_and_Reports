#!/usr/bin/env python3
"""
Libyana NPM - Special (HQ) Reports Processor
Builds the recurring HQ/Tripoli report sheets (starting with the "NQ Data
Collection Template") directly from the pipeline's own output/csv/ history,
instead of the old manual per-sheet scripts (raw SFTP zip -> pandas by
hand). Each build_* function returns one sheet's DataFrame, EAST branch
only, matching the template's column names as closely as the template
itself allows.
"""

import os
import re
import json
import math
import logging
import functools
import pandas as pd

logger = logging.getLogger(__name__)

BRANCH = 'East'

CITY_MAP_FILE = 'config/site_arabic_city_map.json'
CELL_INFO_FILE = 'config/LIBYANA Cell Info.xlsx'
HQ_TRAFFIC_TEMPLATE_FILE = 'config/Traffic and network availability Needed from Tripoli HQ.xlsx'

_SITE_RE = re.compile(r'^(L[A-Z]+\d+)(-\d+)?$')


def _site_of(cell_name: str):
    m = _SITE_RE.match(str(cell_name).strip())
    return m.group(1) if m else None


@functools.lru_cache(maxsize=1)
def _load_city_map():
    """Site -> Arabic city name, built from historical "Cells with High DL
    PRB (EAST)" rows in the NQ template (ground truth) plus a reliable
    (>=90%-consistent) letter-prefix fallback - see
    config/site_arabic_city_map.json and the chat history for how this was
    derived (coordinate-based fallback for the handful of sites with
    neither, verified against Libya's real bounding box to reject bad
    source coordinates)."""
    if not os.path.exists(CITY_MAP_FILE):
        return {}, {}
    with open(CITY_MAP_FILE, encoding='utf-8') as f:
        d = json.load(f)
    return d.get('site_city', {}), d.get('prefix_city', {})


def _city_for_site(site: str) -> str:
    site_city, prefix_city = _load_city_map()
    if site in site_city:
        return site_city[site]
    m = re.match(r'^(L[A-Z]+)', site)
    if m and m.group(1) in prefix_city:
        return prefix_city[m.group(1)]
    return ''


@functools.lru_cache(maxsize=1)
def _load_cell_coords():
    """Cell Name -> (lat, lon) from config/LIBYANA Cell Info.xlsx (East
    sheet), plus a per-site average for cells not found individually."""
    if not os.path.exists(CELL_INFO_FILE):
        return {}, {}
    import openpyxl
    wb = openpyxl.load_workbook(CELL_INFO_FILE, data_only=True, read_only=True)
    ws = wb['East']
    cell_coords = {}
    site_points = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        cell_name, lat, lon = row[1], row[10], row[11]
        if not cell_name or lat is None or lon is None:
            continue
        try:
            lat, lon = float(lat), float(lon)
        except (TypeError, ValueError):
            continue
        cell_coords[str(cell_name).strip()] = (lat, lon)
        site = _site_of(cell_name) or _site_of('L' + str(cell_name).strip())
        if site:
            site_points.setdefault(site, []).append((lat, lon))
    site_coords = {s: (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts))
                   for s, pts in site_points.items()}
    return cell_coords, site_coords


def _coords_for_cell(cell_name: str, site: str):
    cell_coords, site_coords = _load_cell_coords()
    if cell_name in cell_coords:
        return cell_coords[cell_name]
    if site in site_coords:
        return site_coords[site]
    return (None, None)

# Template's 4 DL PRB utilization buckets, in low-to-high order (used for
# the tie-break: equal day-counts in two buckets -> higher bucket wins,
# matching the original script's TIE_BREAKER='higher' default). Lower
# bound of the first bucket is exclusive (0% itself is treated as
# invalid/"Other" and dropped, not "no load"), matching the original
# script's `if 0 < prb_value < 70`.
PRB_BUCKETS = [
    (0, False, 70, 'DL PRB  UT 0%<X<70%'),
    (70, True, 80, 'DL PRB  UT 70%<X<80%'),
    (80, True, 90, 'DL PRB  UT 80%<X<90%'),
    (90, True, 100.0001, 'DL PRB  UT 90%<X<100%'),
]


def _load_csv(csv_folder, name, usecols=None):
    path = os.path.join(csv_folder, f"{name}.csv")
    if not os.path.exists(path):
        logger.warning(f"Special report source not found: {path}")
        return None
    if usecols is not None:
        # The interference archives are multi-hundred-MB/millions-of-rows
        # hourly files with far more columns than this report needs - the
        # pyarrow engine reads only the requested columns roughly 10x
        # faster than the default C engine here (measured: ~8s -> ~0.9s on
        # the 685MB 4G file). Falls back to the default engine if pyarrow
        # isn't installed on a given machine.
        try:
            df = pd.read_csv(path, usecols=usecols, engine='pyarrow')
        except Exception:
            df = pd.read_csv(path, usecols=usecols)
    else:
        df = pd.read_csv(path)
    return df if not df.empty else None


def build_subscribers_report(csv_folder='output/csv') -> pd.DataFrame:
    """Subscribers sheet: one row per ISO week (Wednesday's daily figure -
    output/csv/User_Summary.csv is already one row/day, so "peak of the
    week" simplifies to "that week's Wednesday reading", matching the
    original per-hour-peak script's weekday==3 filter without needing to
    hunt through hourly data that no longer exists at this grain)."""
    df = _load_csv(csv_folder, 'User_Summary')
    if df is None:
        return pd.DataFrame()

    df = df.copy()
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    df = df[df['Date'].dt.weekday == 2]  # Wednesday

    if df.empty:
        return pd.DataFrame()

    rows = []
    for _, r in df.iterrows():
        iu = r.get('Roaming 3G PS (Iu)')
        s1 = r.get('Roaming 4G PS (S1)')
        rows.append({
            'year': r['Date'].year,
            'Week no': f"W{int(r['Date'].isocalendar().week):02d}",
            'Branch': BRANCH,
            'Maximum number of attached subscribers(GSM In SGSN)': r.get('2G PS user'),
            'Maximum number of attached subscribers(UMTS in SGSN )': r.get('3G PS user'),
            'Number of subscribers in VLR (Connected to BSC)': r.get('2G CS user'),
            'Number of subscribers in VLR (Connected to RNC)': r.get('3G CS user'),
            'Max Number of EPS Attach subscribers in MME': r.get('4G PS user'),
            'Number of Registered Subscribers (Almadar in Libyana Metwork)': r.get('Roaming CS (Almadar)'),
            'Number of Registered Subscribers (Almadar in Libyana PS Network) for(3G,4G)':
                f"3G={iu},4G={s1}" if pd.notna(iu) and pd.notna(s1) else '',
        })

    return pd.DataFrame(rows).sort_values(['year', 'Week no']).reset_index(drop=True)


def build_prb_bucket_report(csv_folder='output/csv', min_days_per_cell_month=10, df=None) -> pd.DataFrame:
    """4G Cell Prb Dl ut(%) sheet: majority DL-PRB-utilization bucket per
    cell per month, EAST only, from output/csv/4G_Cell_BH.csv (already
    busy-hour, one value per cell per day - the original script's own
    "already matches the criteria" note). Criteria: Availability >99% (or
    missing - kept per the original script's NIL-handling), >=10 valid
    days in the month, majority bucket wins ties toward the higher bucket.

    `df` lets build_nq_template_report() pass an already-loaded
    4G_Cell_BH frame instead of re-reading it from disk."""
    df = _load_csv(csv_folder, '4G_Cell_BH') if df is None else df
    if df is None or 'Cell Name' not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    df['Year'] = df['Date'].dt.year
    df['Month'] = df['Date'].dt.month
    df['Month_Name'] = df['Date'].dt.strftime('%B')

    avail = pd.to_numeric(df['Radio Network Availability Rate(%)'], errors='coerce')
    prb = pd.to_numeric(df['DL PRB Utilizing Rate(%)'], errors='coerce')

    df = df[(avail.isna()) | (avail > 99)]
    df = df.assign(_prb=prb.loc[df.index]).dropna(subset=['_prb'])

    if df.empty:
        return pd.DataFrame()

    day_counts = df.groupby(['Cell Name', 'Year', 'Month']).size().reset_index(name='Days')
    valid_cells = day_counts[day_counts['Days'] >= min_days_per_cell_month][['Cell Name', 'Year', 'Month']]
    df = df.merge(valid_cells, on=['Cell Name', 'Year', 'Month'], how='inner')

    if df.empty:
        return pd.DataFrame()

    def bucket_of(v):
        for lo, lo_inclusive, hi, label in PRB_BUCKETS:
            if (v >= lo if lo_inclusive else v > lo) and v < hi:
                return label
        return None

    df['Bucket'] = df['_prb'].apply(bucket_of)
    df = df.dropna(subset=['Bucket'])

    bucket_order = {label: i for i, (_, _, _, label) in enumerate(PRB_BUCKETS)}
    counts = df.groupby(['Cell Name', 'Year', 'Month', 'Bucket']).size().reset_index(name='Day_Count')
    counts['Bucket_Order'] = counts['Bucket'].map(bucket_order)
    counts = counts.sort_values(
        ['Cell Name', 'Year', 'Month', 'Day_Count', 'Bucket_Order'],
        ascending=[True, True, True, False, False],
    )
    majority = counts.groupby(['Cell Name', 'Year', 'Month']).first().reset_index()

    monthly_counts = majority.groupby(['Year', 'Month', 'Bucket']).size().reset_index(name='Count')
    month_names = df[['Year', 'Month', 'Month_Name']].drop_duplicates()
    monthly_counts = monthly_counts.merge(month_names, on=['Year', 'Month'], how='left')

    rows = []
    for (year, month), grp in monthly_counts.groupby(['Year', 'Month']):
        month_name = grp['Month_Name'].iloc[0]
        counts_by_bucket = dict(zip(grp['Bucket'], grp['Count']))
        total = 0
        for _, _, _, label in PRB_BUCKETS:
            c = counts_by_bucket.get(label, 0)
            rows.append({'year': year, 'Month': month_name, 'Branch': label, BRANCH: c})
            total += c
        rows.append({'year': year, 'Month': month_name, 'Branch': 'Total', BRANCH: total})

    return pd.DataFrame(rows)


def build_high_prb_cells_report(csv_folder='output/csv', min_days_per_cell_week=4, df=None) -> pd.DataFrame:
    """Cells with High DL PRB (EAST) sheet: weekly (Sun-Sat, matching the
    template's existing date-range format) listing of individual cells
    whose majority DL-PRB bucket that week is 90-100%, with coordinates
    and Arabic city name. Same availability/PRB filter as
    build_prb_bucket_report(), just grouped by week instead of month, and
    listing cells instead of counting them.

    `min_days_per_cell_week` (default 4 of a possible 7) is a judgment
    call, not from your original script (which only defined a monthly
    threshold) - adjust if you want a different bar.

    `df` lets build_nq_template_report() pass an already-loaded
    4G_Cell_BH frame instead of re-reading it from disk."""
    df = _load_csv(csv_folder, '4G_Cell_BH') if df is None else df
    if df is None or 'Cell Name' not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])

    avail = pd.to_numeric(df['Radio Network Availability Rate(%)'], errors='coerce')
    prb = pd.to_numeric(df['DL PRB Utilizing Rate(%)'], errors='coerce')
    df = df[(avail.isna()) | (avail > 99)]
    df = df.assign(_prb=prb.loc[df.index]).dropna(subset=['_prb'])
    if df.empty:
        return pd.DataFrame()

    # Sunday-Saturday week window, matching the template's existing
    # "17-08-2025 to 23-08-2025" date-range format.
    week_start = df['Date'] - pd.to_timedelta((df['Date'].dt.weekday + 1) % 7, unit='D')
    df['_week_start'] = week_start

    def bucket_of(v):
        for lo, lo_inclusive, hi, label in PRB_BUCKETS:
            if (v >= lo if lo_inclusive else v > lo) and v < hi:
                return label
        return None

    df['Bucket'] = df['_prb'].apply(bucket_of)
    df = df.dropna(subset=['Bucket'])

    day_counts = df.groupby(['Cell Name', '_week_start']).size().reset_index(name='Days')
    valid_cells = day_counts[day_counts['Days'] >= min_days_per_cell_week][['Cell Name', '_week_start']]
    df_valid = df.merge(valid_cells, on=['Cell Name', '_week_start'], how='inner')
    if df_valid.empty:
        return pd.DataFrame()

    bucket_order = {label: i for i, (_, _, _, label) in enumerate(PRB_BUCKETS)}
    counts = df_valid.groupby(['Cell Name', '_week_start', 'Bucket']).size().reset_index(name='Day_Count')
    counts['Bucket_Order'] = counts['Bucket'].map(bucket_order)
    counts = counts.sort_values(
        ['Cell Name', '_week_start', 'Day_Count', 'Bucket_Order'],
        ascending=[True, True, False, False],
    )
    majority = counts.groupby(['Cell Name', '_week_start']).first().reset_index()
    high_prb = majority[majority['Bucket'] == PRB_BUCKETS[-1][3]]
    if high_prb.empty:
        return pd.DataFrame()

    # Weekly average PRB (all valid days that week, not just the
    # majority-bucket days) for the cells that made the cut.
    week_avg = df_valid.groupby(['Cell Name', '_week_start'])['_prb'].mean().reset_index(name='avg_prb')
    high_prb = high_prb.merge(week_avg, on=['Cell Name', '_week_start'], how='left')

    rows = []
    for _, r in high_prb.iterrows():
        cell_name = r['Cell Name']
        site = _site_of(cell_name)
        lat, lon = _coords_for_cell(cell_name, site)
        week_end = r['_week_start'] + pd.Timedelta(days=6)
        rows.append({
            'Date': f"{r['_week_start'].strftime('%d-%m-%Y')} to {week_end.strftime('%d-%m-%Y')}",
            'Cell Name': cell_name,
            'Area': BRANCH,
            'DL PRB (%)': r['avg_prb'],
            'Long': lon,
            'Lat': lat,
            'City': _city_for_site(site) if site else '',
        })

    return pd.DataFrame(rows).sort_values(['Date', 'Cell Name']).reset_index(drop=True)


def build_cell_data_report(csv_folder='output/csv', throughput_threshold=3, prb_threshold=70, df=None) -> pd.DataFrame:
    """Cell Data sheet: daily count of EAST 4G cells above the DL
    throughput / PRB utilization thresholds, plus total cell count, from
    output/csv/4G_Cell_BH.csv.

    `df` lets build_nq_template_report() pass an already-loaded
    4G_Cell_BH frame instead of re-reading it from disk."""
    df = _load_csv(csv_folder, '4G_Cell_BH') if df is None else df
    if df is None or 'Cell Name' not in df.columns:
        return pd.DataFrame()

    df = df.copy()
    throughput = pd.to_numeric(df['User Downlink Average Throughput (Mbps)'], errors='coerce')
    prb = pd.to_numeric(df['DL PRB Utilizing Rate(%)'], errors='coerce')

    rows = []
    for date, grp in df.groupby('Date'):
        idx = grp.index
        high_throughput_cells = grp.loc[throughput.loc[idx] > throughput_threshold, 'Cell Name'].nunique()
        high_prb_cells = grp.loc[prb.loc[idx] > prb_threshold, 'Cell Name'].nunique()
        total_cells = grp['Cell Name'].nunique()
        rows.append({
            'Day': date,
            'Region': BRANCH,
            f'Number Of Cells with Average DL Throughput per User >{throughput_threshold}Mbps @cell BH': high_throughput_cells,
            f'Number Of Cells with L PRB Utilization Rate(%t)>{prb_threshold}% @cell BH': high_prb_cells,
            'Total Number of cells': total_cells,
        })

    return pd.DataFrame(rows).sort_values('Day').reset_index(drop=True)


# 3G DL frequency -> band (matches the original script's map_dl_freq)
_3G_BAND_MAP = {3054: 900, 3062: 900, 3075: 900, 10562: 2100, 10587: 2100}
# 4G "Frequency band" (a band INDEX, not the raw EARFCN channel number) -> MHz
_4G_BAND_MAP = {1: 2100, 3: 1800, 8: 900, 28: 700}


INTERFERENCE_PERIODS = ('day', 'week', 'month', 'quarter')

# "More than 5 bad days" in a ~30-day month is really a ~1-in-5-days (20%)
# recurrence bar, meant to separate persistent external interference from a
# one-off blip. Scaling that same duty cycle to whatever window the report
# is grouped by - instead of leaving the day-count frozen at "5" - keeps a
# flagged cell meaning the same thing at every granularity: a "day" view
# degenerates to "was it bad that day", a "quarter" view still demands
# roughly 1-in-5 days bad rather than becoming toothless over ~90 days.
_GSM_BAD_DAY_DUTY_CYCLE = 0.2


def _min_bad_days(period_days) -> int:
    return max(1, math.ceil(period_days * _GSM_BAD_DAY_DUTY_CYCLE))


def _period_cols(dates: pd.Series, period: str):
    """Returns (year_col, label_col, sort_key_col) for grouping at the
    requested granularity. 'month' reproduces the original Year/Month
    behaviour (label = full month name); 'day'/'week'/'quarter' are the
    dashboard's additional zoom levels. sort_key is always a real
    chronological ordinal, separate from label - sorting by the 'month'
    label string directly would put "August" before "July" before
    "September" (plain alphabetical), not calendar order.

    Computed once per UNIQUE date, then broadcast back over the full
    `dates` Series via a dict lookup, instead of running .dt.strftime()
    (or any other per-row accessor) over every row directly. The 3G/4G
    interference archives are multi-million-row hourly data with at most
    a few hundred distinct calendar days in them, and .dt.strftime() on
    the full row count (not the groupby chain, not even reading the
    underlying multi-hundred-MB CSVs) was measured to be THE dashboard
    bottleneck - ~40s per call on a 6.5M-row file just to label rows with
    a month name that's identical for ~200,000 of them at a time."""
    unique_dates = pd.Series(dates.unique())
    u_year = unique_dates.dt.year
    if period == 'day':
        u_label = unique_dates.dt.strftime('%Y-%m-%d')
        u_sort_key = u_label
    elif period == 'week':
        iso = unique_dates.dt.isocalendar()
        u_sort_key = iso.week.astype(int)
        u_label = 'W' + iso.week.astype(str).str.zfill(2)
    elif period == 'quarter':
        u_sort_key = unique_dates.dt.quarter
        u_label = 'Q' + u_sort_key.astype(str)
    else:
        u_sort_key = unique_dates.dt.month
        u_label = unique_dates.dt.strftime('%B')

    year = dates.map(dict(zip(unique_dates, u_year)))
    label = dates.map(dict(zip(unique_dates, u_label)))
    sort_key = dates.map(dict(zip(unique_dates, u_sort_key)))
    return year, label, sort_key


def build_external_interference_report(csv_folder='output/csv', period='month',
                                        gsm_daily_threshold=5,
                                        umts_lte_hourly_threshold_hours=6) -> pd.DataFrame:
    """External Interference sheet, EAST only, one section per technology,
    sourced from the dedicated interference report archived weekly by
    backend/interference_processor.py (2G_Interference.csv /
    3G_Interference_Hourly.csv / 4G_Interference_Hourly.csv), grouped at
    the requested `period` ('day'/'week'/'month'/'quarter'; default
    'month' matches the original behaviour):

    - 2G: from 2G_Interference.csv - already one row per cell per day, no
      hourly averaging needed. Band from 'DL frequency' (GSM900/DCS1800
      text). A cell/day counts as "bad" if that day's Interference Band
      Proportion (4~5)(%) > 5 (gsm_daily_threshold); a cell counts toward
      the period's total if its number of bad days clears a duty-cycle
      bar scaled to the period's length (see _min_bad_days).
    - 3G: from 3G_Interference_Hourly.csv. VS.MeanRTWP > -95dBm counts as
      an "interfered hour"; a cell counts if ANY single day within the
      period had >=6 such hours (exactly the original
      process_3g_interference rule) - this is already a per-day event so
      it needs no duty-cycle scaling, just a different grouping window.
      Band from 'DL frequency' via the same frequency->band table.
    - 4G: from 4G_Interference_Hourly.csv. L.UL.Interference.Avg(dBm) >
      -100dBm counts as an "interfered hour", same >=6-hours-in-a-day
      rule. Band from 'Frequency band' (a band index, not the raw EARFCN
      channel number)."""
    if period not in INTERFERENCE_PERIODS:
        raise ValueError(f"period must be one of {INTERFERENCE_PERIODS}, got {period!r}")

    rows = []

    # ---- 2G ----
    df2 = _load_csv(csv_folder, '2G_Interference')
    if df2 is not None and 'DL frequency' in df2.columns:
        df2 = df2.copy()
        df2['Date'] = pd.to_datetime(df2['Date'], errors='coerce').dt.floor('D')
        df2 = df2.dropna(subset=['Date'])
        df2['_interf'] = pd.to_numeric(df2['Interference Band Proportion (4~5)(%)'], errors='coerce')

        year, label, sort_key = _period_cols(df2['Date'], period)
        df2['Year'], df2['Period'], df2['SortKey'] = year, label, sort_key

        total_per_band = df2.groupby(['Year', 'Period', 'DL frequency'])['Cell Name'].nunique()
        period_days = df2.groupby(['Year', 'Period'])['Date'].nunique()
        sort_key_map = df2.groupby(['Year', 'Period'])['SortKey'].first()

        bad_days = df2[df2['_interf'] > gsm_daily_threshold]
        bad_day_counts = bad_days.groupby(['Cell Name', 'DL frequency', 'Year', 'Period']).size().reset_index(name='BadDays')
        if not bad_day_counts.empty:
            bad_day_counts['MinBadDays'] = [
                _min_bad_days(period_days.get((y, p), 1))
                for y, p in zip(bad_day_counts['Year'], bad_day_counts['Period'])
            ]
            interfered = bad_day_counts[bad_day_counts['BadDays'] >= bad_day_counts['MinBadDays']]
            interfered_per_band = interfered.groupby(['Year', 'Period', 'DL frequency'])['Cell Name'].nunique()
        else:
            interfered_per_band = pd.Series(dtype=int)

        for (year, label_, band), total in total_per_band.items():
            count = interfered_per_band.get((year, label_, band), 0)
            rows.append({'year': year, 'week': label_, 'SortKey': sort_key_map[(year, label_)],
                         'Tech Type': '2G', 'Branch': BRANCH,
                         'Band': band, 'Count of Cells with External interference': count,
                         'Total count of cells': total})

    # ---- 3G / 4G (shared shape: any day with >= N interfered hours flags the cell) ----
    # Reads the small BadHours-per-(Cell,Band,Date) rollup
    # (backend/interference_processor.py's build_daily_rollup/
    # update_daily_rollup, kept current by CSVHistoryManager.
    # update_interference_kpis), NOT the raw hourly archives directly -
    # those are multi-hundred-MB and grow every week with no cap by
    # design, and a plain read of them has been observed to raise
    # MemoryError on a machine already under normal desktop load. The
    # rollup already has "was this cell/band bad, and how many hours,
    # on this day" precomputed, so nothing here needs to touch a metric
    # threshold or an hourly timestamp at all.
    for tech, rollup_sheet in [('3G', '3G_Interference_Daily'), ('4G', '4G_Interference_Daily')]:
        df = _load_csv(csv_folder, rollup_sheet)
        if df is None or 'BadHours' not in df.columns:
            continue
        df = df.copy()
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date'])

        year, label, sort_key = _period_cols(df['Date'], period)
        df['Year'], df['Period'], df['SortKey'] = year, label, sort_key

        total_per_band = df.groupby(['Year', 'Period', 'Band'])['Cell Name'].nunique()
        sort_key_map = df.groupby(['Year', 'Period'])['SortKey'].first()

        bad_cell_days = df[df['BadHours'] >= umts_lte_hourly_threshold_hours]
        interfered_cells = bad_cell_days[['Cell Name', 'Band', 'Year', 'Period']].drop_duplicates()
        interfered_per_band = interfered_cells.groupby(['Year', 'Period', 'Band'])['Cell Name'].nunique()

        for (year, label_, band), total in total_per_band.items():
            count = interfered_per_band.get((year, label_, band), 0)
            rows.append({'year': year, 'week': label_, 'SortKey': sort_key_map[(year, label_)],
                         'Tech Type': tech, 'Branch': BRANCH,
                         'Band': int(band), 'Count of Cells with External interference': count,
                         'Total count of cells': total})

    result = pd.DataFrame(rows).sort_values(['year', 'SortKey', 'Tech Type', 'Band']).reset_index(drop=True)
    result = result.drop(columns=['SortKey'])
    # 2G's Band is text ('GSM900'/'DCS1800') while 3G/4G's is an int MHz value (900/1800/2100/700) -
    # mixing both in one object column is what Streamlit's Arrow conversion above was choking on.
    result['Band'] = result['Band'].astype(str)
    return result


# Sheet1 of the Tripoli-HQ template is 5 side-by-side (Month, Value) column
# pairs, one row per calendar month (row N+1 = month N), EAST branch only.
# Column letter -> tidy field name (None = the repeated "Month" index column).
HQ_TRAFFIC_COLUMNS = {
    'A': None, 'B': '4G PS Traffic (TB)',
    'C': None, 'D': 'Gi Interface Traffic (TB)',
    'E': None, 'F': 'DL PRB Utilization (%)',
    'G': None, 'H': '2G Network Availability (%)',
    'I': None, 'J': 'Avg DL Throughput per User (Mbps)',
}
HQ_TRAFFIC_METRICS = [v for v in HQ_TRAFFIC_COLUMNS.values() if v]


def load_hq_traffic_history(path=HQ_TRAFFIC_TEMPLATE_FILE) -> pd.DataFrame:
    """Reads Sheet1 of the Tripoli-HQ 'Traffic and network availability'
    template as a tidy one-row-per-month table (percent columns converted
    from the file's 0-1 fraction to plain 0-100, matching how percentages
    are shown everywhere else in this dashboard). Months not yet filled in
    (all 4 values blank, e.g. the current month before compute_hq_traffic_month
    has been run for it) are dropped."""
    empty = pd.DataFrame(columns=['Month'] + HQ_TRAFFIC_METRICS)
    if not os.path.exists(path):
        return empty
    try:
        raw = pd.read_excel(path, sheet_name='Sheet1', header=0)
    except Exception:
        logger.warning(f"Could not read HQ traffic template: {path}", exc_info=True)
        return empty
    if raw.shape[1] < len(HQ_TRAFFIC_COLUMNS):
        return empty

    out = pd.DataFrame({'Month': raw.iloc[:, 0]})
    for col_letter, field in HQ_TRAFFIC_COLUMNS.items():
        if field is None:
            continue
        idx = ord(col_letter) - ord('A')
        values = pd.to_numeric(raw.iloc[:, idx], errors='coerce')
        out[field] = values * 100 if '%' in field else values

    out = out.dropna(subset=['Month'])
    out = out[out[HQ_TRAFFIC_METRICS].notna().any(axis=1)]
    out['Month'] = out['Month'].astype(int)
    return out.sort_values('Month').reset_index(drop=True)


def compute_hq_traffic_month(csv_folder='output/csv', year=None, month=None) -> dict:
    """Computes the 5 Tripoli-HQ 'Traffic and network availability' metrics
    for one calendar month, EAST branch, straight from the pipeline's own
    output/csv/ history - the same 5 metrics as Sheet1 of the HQ template:

      - 4G PS Traffic (TB): sum of 'PS 4G Sgi_traffic (TB)' (Gi_Interface_Traffic.csv)
      - Gi Interface Traffic (TB): sum of 'PS 234G traffic (TB)', the combined
        2G/3G/4G Gi/Sgi volume (same file)
      - DL PRB Utilization (%): mean of 'DL PRB Utilizing Rate(%)' (4G_NW_Daily.csv)
      - 2G Network Availability (%): mean of 'RR307:TCH Availability(%)'
        (2G_NW_Daily.csv) - the GSM availability KPI per config/kpi_thresholds.csv
      - Avg DL Throughput per User (Mbps): mean of 'User Downlink Average
        Throughput (Mbps)' (4G_NWBH.csv - whole-network busy-hour figure;
        the per-cell 4G_Cell_BH.csv average runs ~2x higher because it's an
        unweighted mean across cells of very different load, so it doesn't
        match this template's existing Jan-Jul scale the way the network-level
        busy-hour KPI does)

    A metric is left as None when its source has no rows that month at all.
    `coverage` reports (days_with_data, days_in_month) per source so a partial
    month (e.g. ingestion only started mid-month) is visible rather than
    silently averaged/summed as if the month were complete."""
    import calendar as _calendar
    days_in_month = _calendar.monthrange(year, month)[1]

    def _month_slice(df):
        d = pd.to_datetime(df['Date'], errors='coerce')
        return df[(d.dt.year == year) & (d.dt.month == month)]

    result = {'Month': month, 'Year': year, 'coverage': {}}
    for field in HQ_TRAFFIC_METRICS:
        result[field] = None

    gi = _load_csv(csv_folder, 'Gi_Interface_Traffic')
    if gi is not None and 'Date' in gi.columns:
        gi_m = _month_slice(gi)
        if not gi_m.empty:
            result['4G PS Traffic (TB)'] = pd.to_numeric(gi_m['PS 4G Sgi_traffic (TB)'], errors='coerce').sum()
            result['Gi Interface Traffic (TB)'] = pd.to_numeric(gi_m['PS 234G traffic (TB)'], errors='coerce').sum()
            result['coverage']['Gi_Interface_Traffic.csv'] = (len(gi_m), days_in_month)

    nw4g = _load_csv(csv_folder, '4G_NW_Daily')
    if nw4g is not None and 'Date' in nw4g.columns:
        nw4g_m = _month_slice(nw4g)
        if not nw4g_m.empty:
            result['DL PRB Utilization (%)'] = pd.to_numeric(nw4g_m['DL PRB Utilizing Rate(%)'], errors='coerce').mean()
            result['coverage']['4G_NW_Daily.csv'] = (len(nw4g_m), days_in_month)

    nw2g = _load_csv(csv_folder, '2G_NW_Daily')
    if nw2g is not None and 'Date' in nw2g.columns:
        nw2g_m = _month_slice(nw2g)
        if not nw2g_m.empty:
            result['2G Network Availability (%)'] = pd.to_numeric(nw2g_m['RR307:TCH Availability(%)'], errors='coerce').mean()
            result['coverage']['2G_NW_Daily.csv'] = (len(nw2g_m), days_in_month)

    nwbh4g = _load_csv(csv_folder, '4G_NWBH')
    if nwbh4g is not None and 'Date' in nwbh4g.columns:
        nwbh4g_m = _month_slice(nwbh4g)
        if not nwbh4g_m.empty:
            result['Avg DL Throughput per User (Mbps)'] = pd.to_numeric(
                nwbh4g_m['User Downlink Average Throughput (Mbps)'], errors='coerce').mean()
            result['coverage']['4G_NWBH.csv'] = (len(nwbh4g_m), days_in_month)

    return result


def save_hq_traffic_month(computed: dict, path=HQ_TRAFFIC_TEMPLATE_FILE) -> None:
    """Writes one month's computed values into Sheet1 of the Tripoli-HQ
    template in place - values only, so Sheet2 and all existing formatting
    survive - the same file then accumulates real monthly history (row N+1
    = month N) the way it already holds Jan-Jul, and next month only the
    newest row needs computing. Percent fields are stored back as the
    file's native 0-1 fraction. Skips any metric that is None (nothing
    computed for it) rather than blanking out a previously-saved value."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb['Sheet1']
    row = computed['Month'] + 1  # header is row 1, month N is row N+1

    for col_letter, field in HQ_TRAFFIC_COLUMNS.items():
        cell = ws[f'{col_letter}{row}']
        if field is None:
            cell.value = computed['Month']
        else:
            value = computed.get(field)
            if value is None:
                continue
            cell.value = (value / 100) if '%' in field else value
        above = ws[f'{col_letter}{row - 1}']
        if above.number_format and above.number_format != 'General':
            cell.number_format = above.number_format

    wb.save(path)


def build_network_daily_kpis_report(csv_folder='output/csv') -> pd.DataFrame:
    """Network Daily KPI's sheet, EAST branch only. The template also
    carries North/South/Middle/West columns and ~14 "Libyana" (nationwide)
    totals per KPI - this pipeline only ever has FTP/OSS access to the
    East branch's own Huawei exports, so those columns are structurally
    impossible to produce here and are intentionally left out (confirmed
    with the user 2026-09-14: "we only use east area").

    Column sourcing:
    - PS Traffic (TB) per tech: 4G_NW_Daily (DL+UL Traffic Volume, matching
      cell_info_report.py's own DL+UL convention - NOT the DL-only figure
      Traffic_Network_4G.csv uses elsewhere in this dashboard, so this
      column will read slightly higher than that one for the same dates),
      3G_NW_Daily ('PS traffic (UL+DL)(GB)'), 2G_NW_Daily ('PS Traffic
      (RLC)(MB)') - GB/MB converted to TB. "PS Traffic (TB)" is the sum of
      the three.
    - Gi Interface Traffic Volume (TB): Gi_Interface_Traffic.csv's own
      'PS 234G traffic (TB)' - this file has a ~1-2 day reporting lag baked
      into Huawei's own report generation (confirmed 2026-09-14; see
      FRESHNESS_EXPECTED_LAG in report_generator.py), unrelated to this
      pipeline, so this column will run behind the others here until the
      user's upstream fix lands.
    - VoLTE/3G/2G CS traffic and their "Voice Traffic (Erl)" sum: direct
      NW_Daily columns, same counters cell_info_report.py already uses for
      the equivalent per-tech traffic figures.
    - Network Availability (%): 4G from 4G_NW_Daily's own dedicated
      availability column. 3G from 3G_NWBH's 'Availability_all level' -
      NOT 3G_NW_Daily/3G_NWBH's plain 'Availability' column, which despite
      the name is some unrelated unbounded metric (checked directly:
      consistently large negative values, not a 0-100 percentage at all).
      2G uses 'RR307:TCH Availability(%)' as the primary metric per the
      user's explicit 2026-09-14 direction (no dedicated "network
      availability" counter exists for 2G in this export) - same counter
      compute_hq_traffic_month() already uses for the Tripoli-HQ template.
    - Maximum Number of RRC Connection Users: 4G_NW_Daily's
      'L.Traffic.User.Max' (same counter as the Overview tab's "LTE
      maximum attached users").
    - DL/UL Throughput, DL/UL PRB, E-RAB/RRC Setup Success Rate: 4G_NWBH
      (busy-hour), matching how these are used elsewhere in this project.
    - E-RAB Drop Rate / RRC Drop Rate: counters the user added to the
      Huawei export config 2026-09-14 ('E-RAB Drop Rate of QCI1(CMCC
      Cell)-ZM' and 'RRC Drop Rate (%)') - not present in historical
      exports yet, so these read all-NaN until a future pipeline run
      picks them up. Checked in both 4G_NWBH and 4G_NW_Daily since which
      report they land in isn't confirmed yet - if they still read NaN
      once the user says fresh data has the new counters, check the other
      raw report/sheet Huawei actually put them in.
    - 3G/2G CSSR and CDR: 'Call Setup Success Rate(%)' (2G) / '..._EFD'
      (3G) variants, 'TCH Drop Rate(%)' (2G) / 'CS Call Drop Rate(%)' (3G).
    - Latency / Packet Loss Rate: Transmission_KPIs.csv's 'Avg Delay(ms)'/
      'Avg Packet Loss(%)', averaged network-wide per day - that file is
      per Site+Adjacent-Node, not already one row per day, so this is a
      network-wide daily mean rather than a single dedicated counter.
    - VoLTE Success Rate: 4G_NW_Daily's 'VoLTE Setup Success Rate-ZM(%)'.
    """
    def _load_dated(name):
        df = _load_csv(csv_folder, name)
        if df is None or 'Date' not in df.columns:
            return None
        df = df.copy()
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.normalize()
        df = df.dropna(subset=['Date']).drop_duplicates(subset='Date', keep='last')
        return df.set_index('Date')

    nw2g = _load_dated('2G_NW_Daily')
    nw3g = _load_dated('3G_NW_Daily')
    nw4g = _load_dated('4G_NW_Daily')
    bh3g = _load_dated('3G_NWBH')
    bh4g = _load_dated('4G_NWBH')
    gi = _load_dated('Gi_Interface_Traffic')

    trans_daily = None
    trans = _load_csv(csv_folder, 'Transmission_KPIs')
    if trans is not None and 'Date' in trans.columns:
        t = trans.copy()
        t['Date'] = pd.to_datetime(t['Date'], errors='coerce').dt.normalize()
        t = t.dropna(subset=['Date'])
        trans_daily = t.groupby('Date')[['Avg Delay(ms)', 'Avg Packet Loss(%)']].mean()
        trans_daily = trans_daily.rename(columns={
            'Avg Delay(ms)': 'Latency (ms)', 'Avg Packet Loss(%)': 'Packet Loss Rate (%)',
        })

    frames = [df for df in (nw2g, nw3g, nw4g, bh3g, bh4g, gi, trans_daily) if df is not None]
    if not frames:
        return pd.DataFrame()
    all_dates = sorted(set().union(*[df.index for df in frames]))
    out = pd.DataFrame(index=pd.Index(all_dates, name='Date'))

    def g(df, col):
        if df is None or col not in df.columns:
            return None
        return pd.to_numeric(df[col], errors='coerce').reindex(out.index)

    def g_any(*candidates):
        for df, col in candidates:
            s = g(df, col)
            if s is not None:
                return s
        return pd.Series(float('nan'), index=out.index, dtype='float64')

    traffic_4g_tb = (g(nw4g, 'DL Traffic  Volume(GB)') + g(nw4g, 'UL Traffic  Volume(GB)')) / 1024
    traffic_3g_tb = g(nw3g, 'PS traffic (UL+DL)(GB)') / 1024
    traffic_2g_tb = g(nw2g, 'PS Traffic (RLC)(MB)') / 1024 / 1024
    out['4G PS Traffic (TB)'] = traffic_4g_tb
    out['3G PS Traffic (TB)'] = traffic_3g_tb
    out['2G PS Traffic (TB)'] = traffic_2g_tb
    out['PS Traffic (TB)'] = traffic_4g_tb.add(traffic_3g_tb, fill_value=0).add(traffic_2g_tb, fill_value=0)

    out['Gi Interface Traffic Volume (TB)'] = g(gi, 'PS 234G traffic (TB)')

    volte_erl = g(nw4g, 'VoLTE Traffic Volume (Erl)')
    cs3g_erl = g(nw3g, 'CS Traffic(Erl)')
    cs2g_erl = g(nw2g, 'K3014:Traffic Volume on TCH(Erl)')
    out['VoLTE Voice Traffic (Erl)'] = volte_erl
    out['3G CS Traffic (Erl)'] = cs3g_erl
    out['2G CS Traffic (Erl)'] = cs2g_erl
    out['Voice Traffic (Erl)'] = volte_erl.add(cs3g_erl, fill_value=0).add(cs2g_erl, fill_value=0)

    out['4G Network Availability (%)'] = g(nw4g, 'Radio Network Availability Rate(%)')
    out['3G Network Availability (%)'] = g(bh3g, 'Availability_all level')
    out['2G Network Availability (%)'] = g(nw2g, 'RR307:TCH Availability(%)')

    out['Maximum Number of RRC Connection Users'] = g(nw4g, 'L.Traffic.User.Max')
    out['Average DL Throughput per User (Mbps)'] = g(bh4g, 'User Downlink Average Throughput (Mbps)')
    out['Average UL Throughput per User (Mbps)'] = g(bh4g, 'User Uplink Average Throughput (Mbps)')
    out['DL PRB Utilization Rate (%)'] = g(bh4g, 'DL PRB Utilizing Rate(%)')
    out['UL PRB Utilization Rate (%)'] = g(bh4g, 'UL PRB Utilizing Rate(%)')
    out['E-RAB Setup Success Rate (%)'] = g(bh4g, 'E-RAB Setup Success Rate')
    out['RRC Setup Success Rate (%)'] = g(bh4g, 'RRC Setup Success Rate(%)')
    out['E-RAB Drop Rate (%)'] = g_any(
        (bh4g, 'E-RAB Drop Rate of QCI1(CMCC Cell)-ZM'), (nw4g, 'E-RAB Drop Rate of QCI1(CMCC Cell)-ZM'),
    )
    out['RRC Drop Rate (%)'] = g_any((bh4g, 'RRC Drop Rate (%)'), (nw4g, 'RRC Drop Rate (%)'))

    out['3G CSSR (%)'] = g(nw3g, 'Call Setup Success Rate(%)_EFD')
    out['2G CSSR (%)'] = g(nw2g, 'Call Setup Success Rate(%)')
    out['3G CDR (%)'] = g(nw3g, 'CS Call Drop Rate(%)')
    out['2G CDR (%)'] = g(nw2g, 'TCH Drop Rate(%)')

    if trans_daily is not None:
        out['Latency (ms)'] = trans_daily['Latency (ms)'].reindex(out.index)
        out['Packet Loss Rate (%)'] = trans_daily['Packet Loss Rate (%)'].reindex(out.index)
    else:
        out['Latency (ms)'] = float('nan')
        out['Packet Loss Rate (%)'] = float('nan')
    out['VoLTE Success Rate (%)'] = g(nw4g, 'VoLTE Setup Success Rate-ZM(%)')

    out.insert(0, 'Branch', BRANCH)
    out = out.reset_index().sort_values('Date').reset_index(drop=True)
    out['Date'] = out['Date'].dt.strftime('%Y-%m-%d')
    return out


def build_nq_template_report(csv_folder='output/csv', interference_period='month') -> dict:
    """All NQ Data Collection Template sheets, EAST only."""
    bh_4g = _load_csv(csv_folder, '4G_Cell_BH')
    return {
        'Subscribers': build_subscribers_report(csv_folder),
        "Network Daily KPI's": build_network_daily_kpis_report(csv_folder),
        '4G Cell Prb Dl ut(%)': build_prb_bucket_report(csv_folder, df=bh_4g),
        'Cell Data': build_cell_data_report(csv_folder, df=bh_4g),
        'Cells with High DL PRB (EAST)': build_high_prb_cells_report(csv_folder, df=bh_4g),
        'External Interference': build_external_interference_report(csv_folder, period=interference_period),
    }
