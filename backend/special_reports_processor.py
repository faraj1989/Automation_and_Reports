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


def _load_csv(csv_folder, name):
    path = os.path.join(csv_folder, f"{name}.csv")
    if not os.path.exists(path):
        logger.warning(f"Special report source not found: {path}")
        return None
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
    "September" (plain alphabetical), not calendar order."""
    year = dates.dt.year
    if period == 'day':
        label = dates.dt.strftime('%Y-%m-%d')
        sort_key = label
    elif period == 'week':
        iso = dates.dt.isocalendar()
        sort_key = iso.week.astype(int)
        label = 'W' + iso.week.astype(str).str.zfill(2)
    elif period == 'quarter':
        sort_key = dates.dt.quarter
        label = 'Q' + sort_key.astype(str)
    else:
        sort_key = dates.dt.month
        label = dates.dt.strftime('%B')
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

    # ---- 3G / 4G (shared shape: hourly threshold -> interfered-hours/day -> any day >= N hours) ----
    for tech, sheet, metric_col, threshold, band_col, band_map in [
        ('3G', '3G_Interference_Hourly', 'VS.MeanRTWP', -95, 'DL frequency', _3G_BAND_MAP),
        ('4G', '4G_Interference_Hourly', 'L.UL.Interference.Avg(dBm)', -100, 'Frequency band', _4G_BAND_MAP),
    ]:
        df = _load_csv(csv_folder, sheet)
        if df is None or band_col not in df.columns or metric_col not in df.columns:
            continue
        df = df.copy()
        df['Date'] = pd.to_datetime(df['Time'], errors='coerce').dt.floor('D')
        df = df.dropna(subset=['Date'])
        df['_metric'] = pd.to_numeric(df[metric_col], errors='coerce')
        df['Band'] = pd.to_numeric(df[band_col], errors='coerce').map(band_map)
        df = df.dropna(subset=['Band'])
        df['_is_interfered_hour'] = df['_metric'] > threshold

        year, label, sort_key = _period_cols(df['Date'], period)
        df['Year'], df['Period'], df['SortKey'] = year, label, sort_key

        total_per_band = df.groupby(['Year', 'Period', 'Band'])['Cell Name'].nunique()
        sort_key_map = df.groupby(['Year', 'Period'])['SortKey'].first()

        daily_hours = df.groupby(['Cell Name', 'Band', 'Year', 'Period', 'Date'])['_is_interfered_hour'].sum().reset_index(name='Hours')
        bad_cell_days = daily_hours[daily_hours['Hours'] >= umts_lte_hourly_threshold_hours]
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


def build_nq_template_report(csv_folder='output/csv', interference_period='month') -> dict:
    """All currently-ready NQ Data Collection Template sheets, EAST only.
    Network Daily KPI's still needs column-by-column source confirmation
    for a few columns (Gi Interface is settled; E-RAB/RRC Drop Rate, Max
    RRC Connection User, and Packet Loss Rate meaning are still open) and
    isn't included yet."""
    bh_4g = _load_csv(csv_folder, '4G_Cell_BH')
    return {
        'Subscribers': build_subscribers_report(csv_folder),
        '4G Cell Prb Dl ut(%)': build_prb_bucket_report(csv_folder, df=bh_4g),
        'Cell Data': build_cell_data_report(csv_folder, df=bh_4g),
        'Cells with High DL PRB (EAST)': build_high_prb_cells_report(csv_folder, df=bh_4g),
        'External Interference': build_external_interference_report(csv_folder, period=interference_period),
    }
