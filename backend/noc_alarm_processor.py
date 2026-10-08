#!/usr/bin/env python3
"""
Libyana NPM - NOC Alarm Feed Reader

Reads (read-only) raw MAE/NetEco/NCE alarm exports (current + historical)
written by scrapers/ in this same project (mae_scraper.py,
neteco_continuous all alrams.py, nce_active_alarms_scraper.py and their
historical counterparts, all managed by scrapers/scraper_watchdog.py) to
<DATA_ROOT>/Output/...

Originally this read a separate NOC Automation Suite's own pre-computed
rollups (Live_NOC_Report.xlsx from its Merge Reports service,
historical_insights_snapshot.json from its historical_noc_analysis.py
ledger). Once scraping moved into this project (2026-09-10) those two
sibling services had no guarantee of still running, so build_live_
disconnected_sites() and build_historical_insights_native() replaced them -
computed directly from the raw exports below, with no dependency on any
process outside this project. (The sibling ledger also had a history of
crashing itself with MemoryError and going stale for days, so this is a
reliability improvement too, not just an independence one.)

Every function here returns a not-loaded result rather than raising if its
raw exports aren't present yet - same convention as topology_processor.py's
handling of a missing FN-HUB Excel file.

build_daily_noc_alarm_report and its helpers (_attach_alarm_reasons,
_nce_reason_for_site, _is_power_reason) answer "how many sites went down on
2026-09-08, for how many total hours, and what did NetEco/NCE each report
as the reason" for one specific calendar day; build_live_disconnected_sites
and build_historical_insights_native reuse the exact same per-site logic
for "right now" and "rolled up over the last N days" respectively, so all
three views agree on what counts as a Power vs. Transmission cause. Column
parsing is a straight port of the original sibling suite's own
historical_alarm_parser.py / site_matching.py logic (pure, dependency-free
modules - no runtime dependency on that project).
"""
import csv
import os
import logging
import re
import zipfile
from datetime import datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from backend.topology_processor import build_site_ancestor_map

logger = logging.getLogger(__name__)

# Root of the sibling NOC Automation Suite's output tree (its own
# DATA_ROOT/Output). Overridable via env var for a different machine/install
# without touching code.
NOC_ALARM_DATA_ROOT = os.environ.get(
    "NOC_ALARM_DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data\Output"
)

# Current (live) alarm export dirs/globs - same layout the continuous
# scrapers themselves write to (scrapers/mae_scraper.py,
# neteco_continuous all alrams.py, nce_active_alarms_scraper.py). Each
# keeps only its single latest export per day (keep_only_latest_export),
# so "most recently modified file under here" is always "right now."
CURRENT_ALARMS_DIR = os.path.join(NOC_ALARM_DATA_ROOT, "Current_Alarms")
NCE_CURRENT_ALARMS_DIR = os.path.join(NOC_ALARM_DATA_ROOT, "NCE_Current_Alarms")
MAE_CURRENT_GLOB = "CurrentAlarms_MAE_*"
NETECO_CURRENT_GLOB = "NetEco_All_Current_Alarm_*"
NCE_CURRENT_GLOB = "CurrentAlarms_NCE_*"

# Mirrors the sibling suite's own Telegram bot (bots/telegram_noc_bot.py's
# POWER_ALARMS) so a disconnected site's root cause reads the same way here
# as it does there - a site with none of these terms in its Power Reason is
# the bot's "Check TX/Link (No Power Alarm)" case, i.e. transmission/other.
POWER_ALARM_TERMS = ("Mains Failure", "BLVD", "LLVD")

# Judgment calls, not derived from the sibling suite's own config: its
# scrapers cycle every ~5 min (live feed) and its historical analysis every
# few minutes to hours, so a live feed older than 30 min or historical
# insights older than 2 days means a scraper is very likely stalled, not
# just running a bit behind - flag it rather than presenting stale alarm
# counts as current.
LIVE_STALE_MINUTES = 30
HISTORICAL_STALE_MINUTES = 60 * 24 * 2


def is_available() -> bool:
    return os.path.isdir(NOC_ALARM_DATA_ROOT)


def _find_latest_export(base_dir: str, glob_pattern: str) -> Optional[Path]:
    """Most-recently-modified file under base_dir/<dated-subfolder>/
    matching glob_pattern - the current-alarm scrapers only keep their
    single latest export per day, so this is always "the current snapshot"
    regardless of which dated folder it landed in (matters right after
    midnight, when today's folder may not have a file yet)."""
    base_path = Path(base_dir)
    if not base_path.exists():
        return None
    candidates = list(base_path.glob(f"*/{glob_pattern}"))
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def get_site_alarm_status(site_names: List[str]) -> Dict:
    """For the given site(s): any currently-active alarm (from the live
    feed) plus historical chronic-offender/downtime rows (from the
    insights snapshot). Each piece is independently optional - a site can
    have one, the other, both, or neither, and either source can be
    entirely unavailable without blocking the other."""
    empty = pd.DataFrame()
    if not site_names:
        return {'current': empty, 'chronic': empty, 'downtime': empty}

    live = build_live_disconnected_sites()
    current = empty
    if live.get('loaded') and 'Site Name' in live['alarms'].columns:
        current = live['alarms'][live['alarms']['Site Name'].isin(site_names)]

    hist = build_historical_insights_native()
    chronic = empty
    downtime = empty
    if hist.get('loaded'):
        co = hist['chronic_offenders']
        if not co.empty and 'Site' in co.columns:
            chronic = co[co['Site'].isin(site_names)]
        sd = hist['site_downtime']
        if not sd.empty and 'Site' in sd.columns:
            downtime = sd[sd['Site'].isin(site_names)]

    return {'current': current, 'chronic': chronic, 'downtime': downtime}


# ============================================================
# Daily NOC alarm analysis - per-day MAE/NetEco/NCE merge from raw
# historical exports (see module docstring for why this bypasses the
# sibling suite's own ledger/analysis pipeline).
# ============================================================

# Historical scrapers write <base_dir>/<YYYY-MM-DD>/<file>, retained only a
# few days (HISTORICAL_RAW_EXPORT_RETENTION_DAYS in the sibling suite) -
# MAE and NetEco share one base dir, NCE has its own.
MAE_HISTORICAL_DIR = os.path.join(NOC_ALARM_DATA_ROOT, "Historical_Alarms")
NETECO_HISTORICAL_DIR = os.path.join(NOC_ALARM_DATA_ROOT, "Historical_Alarms")
NCE_HISTORICAL_DIR = os.path.join(NOC_ALARM_DATA_ROOT, "NCE_Historical_Alarms")

MAE_HISTORICAL_GLOB = "HistoricalAlarms_MAE_*"
NETECO_HISTORICAL_GLOB = "NetEco_Historical_Alarm_*"
NCE_HISTORICAL_GLOB = "HistoricalAlarms_NCE_*"

# The definitive "site is down" signal - same alarm name the live feed's
# Current Alarms Summary and the sibling suite's own triage
# (enhanced_noc_analysis.py's ne_disconnected check) both key off.
SITE_DOWN_ALARM_NAME = "NE Is Disconnected"

# canonical field -> candidate column names, in preference order. Ported
# from the sibling suite's historical_alarm_parser.py (MAE/NetEco/NCE
# genuinely use different column sets for the same fields - see that
# module's docstring for the confirmed-against-real-exports column list
# per source).
_HISTORICAL_FIELD_ALIASES = {
    "site": ("MO Name", "Site Name", "Alarm Source"),
    "name": ("Name",),
    "occurred": ("Occurred On (NT)", "Last Occurred (NT)", "Last Occurred", "Last Occurred (ST)"),
    "cleared": ("Cleared On (NT)", "Cleared On", "Cleared On (ST)"),
    "severity": ("Severity",),
    "alarm_id": ("Alarm ID",),
}

# Canonical site code out of a decorated MAE/NetEco/NCE field (e.g.
# "SLOG002_ATN950D_A" -> "SLOG002", "BGZ087(FTTS)" -> "BGZ087") - ported
# from the sibling suite's site_matching.extract_site_code(); see there
# for the full reasoning behind each rule.
_SITE_CODE_RE = re.compile(r"[A-Z]{2,8}\d{1,4}")


def _extract_site_code(raw) -> Optional[str]:
    """Canonical site code out of a decorated MAE/NetEco/NCE field.

    Two genuinely different raw formats share this one function:
      - Historical/NCE "Alarm Source": SITECODE_DeviceModel_Port (e.g.
        "SLOG002_ATN950D_A") - the segment before the first "_" IS the site
        by construction, digits or not.
      - Current Alarms' "MO Name": "Key=Value, Key2=Value2, ..." (e.g.
        "eNodeB Function Name=LBGZ049, Local Cell ID=2, ... indication=
        CELL_FDD") - an incidental "_" inside a VALUE (like "CELL_FDD")
        must not be mistaken for the first format's site-delimiter, or the
        site truncates to a lowercase-starting fragment and is dropped.
    The presence of "=" reliably distinguishes the two - only the
    underscore-split rule needs to be skipped for key=value strings; the
    general regex works for both.

    A second decoration sits on top of that: Huawei's per-RAT device name
    prefixes the real site code with one technology letter - "NodeB
    Name=U<code>" / "Label=U<code>-N" for 3G, "eNodeB Function Name=L<code>"
    for 4G - while the 2G/OSS-level MO Name uses the bare code untouched
    (confirmed against output/csv/SiteDetail.csv, which only lists the bare
    form, e.g. site "SAR001" shows up as NodeB Name "USAR001"). Left
    unstripped, a real single-site outage reads as non-overlapping "down"
    populations across NE Is Disconnected (bare) vs NodeB
    Unavailable/UMTS Cell Unavailable (U-prefixed) - confirmed live
    2026-09-10 on site SAR001, where one NE Is Disconnected event had
    GSM Cell out of Service + CSL Fault (bare "SAR001") and UMTS Cell
    Unavailable + NodeB Unavailable (both "USAR001") fire as its own
    MAE-flagged Correlative alarms within the same 5-minute window - one
    physical event, wrongly split into two "sites" without this stripping.
    Only strip when the decorating key is actually present, not on every
    leading U/L: a handful of real sites (UMSF001, UQUB001, LAHB001,
    LAQN001, LMLD001) already start with U/L on their own and appear as
    bare MO Name tokens (NE Type BTS3900/OSS) that must pass through
    unchanged. Unverified edge case: if one of those sites has its own
    3G/4G radio, its NodeB/eNodeB name is assumed to double the letter
    (e.g. "UUMSF001") the same way the bare-code sites single-prefix - no
    live example has actually shown up yet to confirm that.
    """
    if raw is None:
        return None
    text = str(raw).strip()
    if not text or text.lower() == "nan":
        return None
    if "_" in text and "=" not in text:
        candidate = text.split("_", 1)[0].strip()
        return candidate if candidate and candidate[0].isupper() and candidate.isalnum() else None
    match = _SITE_CODE_RE.search(text)
    if not match:
        return None
    code = match.group(0)
    if code[0] == "U" and ("NodeB Name=" in text or re.search(r"\bLabel=U", text)):
        code = code[1:]
    elif code[0] == "L" and "eNodeB Function Name=" in text:
        code = code[1:]
    return code or None


def _is_historical_header_row(cells) -> bool:
    stripped = [str(c).strip() for c in cells]
    has_site = "MO Name" in stripped or "Site Name" in stripped or "Alarm Source" in stripped
    has_marker = "Severity" in stripped or "Alarm ID" in stripped
    return "Name" in stripped and has_marker and has_site


def _iter_historical_row_batches(export_path: Path):
    """Yield rows (already CSV-split) for every tabular payload in an
    export - a direct .csv, or a .zip of 1+ "HistoricalAlarmsN.csv" parts
    (the historical scrapers' own output shape)."""
    suffix = export_path.suffix.lower()
    if suffix == ".zip":
        with zipfile.ZipFile(export_path) as zf:
            names = sorted(n for n in zf.namelist() if n.lower().endswith(".csv"))
            for name in names:
                text = zf.read(name).decode("utf-8-sig", errors="replace")
                yield list(csv.reader(text.splitlines()))
    else:
        text = export_path.read_text(encoding="utf-8-sig", errors="replace")
        yield list(csv.reader(text.splitlines()))


_HISTORICAL_COLUMNS = ["Source", "Site", "Name", "Occurred On", "Cleared On", "Severity", "Alarm ID"]


def _parse_historical_export(export_path: Path, source: str) -> pd.DataFrame:
    """Cached front for _parse_historical_export_uncached - keyed on the
    file's mtime too, so a re-written export is re-parsed. The all-days
    range view (build_daily_noc_alarm_range) asks for the same handful of
    rolling-window exports once per day it covers; without this each of
    them was re-parsed dozens of times. Returns a fresh copy each call."""
    path = Path(export_path)
    return _parse_historical_export_cached(str(path), path.stat().st_mtime).assign(Source=source)


@lru_cache(maxsize=96)
def _parse_historical_export_cached(path_str: str, mtime: float) -> pd.DataFrame:
    return _parse_historical_export_uncached(Path(path_str), "")


def _parse_historical_export_uncached(export_path: Path, source: str) -> pd.DataFrame:
    """Normalize one MAE/NetEco/NCE Historical Alarms export (.csv or .zip
    of parts) into a common schema. The header isn't at a fixed row (a
    preamble of title/save-time/username precedes it, and an embedded
    multi-line quoted field makes line-counting unreliable), so it's
    located by content match instead."""
    all_records = []
    for rows in _iter_historical_row_batches(export_path):
        header = None
        header_idx = None
        for idx, row in enumerate(rows):
            if row and _is_historical_header_row(row):
                header = [str(c).strip() for c in row]
                header_idx = idx
                break
        if header is None:
            continue
        col_index = {name: i for i, name in enumerate(header) if name}
        field_index = {}
        for field, aliases in _HISTORICAL_FIELD_ALIASES.items():
            for alias in aliases:
                if alias in col_index:
                    field_index[field] = col_index[alias]
                    break

        def get(row, field):
            idx = field_index.get(field)
            if idx is None or idx >= len(row):
                return None
            value = str(row[idx]).strip()
            return value if value and value.lower() != "nan" else None

        for row in rows[header_idx + 1:]:
            if not row or all(not str(c).strip() for c in row):
                continue
            all_records.append({
                "Site": get(row, "site"), "Name": get(row, "name"),
                "Occurred On": get(row, "occurred"), "Cleared On": get(row, "cleared"),
                "Severity": get(row, "severity"), "Alarm ID": get(row, "alarm_id"),
            })

    df = pd.DataFrame.from_records(all_records, columns=_HISTORICAL_COLUMNS[1:])
    df.insert(0, "Source", source)
    if df.empty:
        return df
    df["Occurred On"] = pd.to_datetime(df["Occurred On"], errors="coerce")
    df["Cleared On"] = pd.to_datetime(df["Cleared On"], errors="coerce")
    df["Site Code"] = df["Site"].apply(_extract_site_code)
    return df


def _find_best_historical_export(base_dir: str, glob_pattern: str, target_date: str):
    """(export path, covers_full_day) for target_date, or (None, False) if
    no export from target_date's folder onward exists.

    Each historical export is a rolling window, not full history - checked
    2026-10-08: MAE ~7 days back, NetEco/NCE ~3 days. The newest export
    has the most complete Cleared-On picture, but for an older target_date
    its window can start partway through that day (today's 09:19 MAE
    export started at 10-01 09:20), silently dropping that day's morning
    outages. So: walk dated folders newest-first, and take the latest
    export of the first folder whose window (earliest Occurred On) reaches
    back to target_date's midnight. If none does (the day falls in a
    scraping gap), fall back to whichever export reaches back furthest,
    flagged as partial."""
    base_path = Path(base_dir)
    if not base_path.exists():
        return None, False
    try:
        target = datetime.strptime(target_date, "%Y-%m-%d").date()
    except ValueError:
        return None, False
    day_start = pd.Timestamp(target)

    folders = []
    for entry in base_path.iterdir():
        if not entry.is_dir() or len(entry.name) != 10:
            continue
        try:
            folder_date = datetime.strptime(entry.name, "%Y-%m-%d").date()
        except ValueError:
            continue
        if folder_date >= target:
            folders.append((folder_date, entry))

    best_partial, best_partial_start = None, None
    for _, folder in sorted(folders, reverse=True):
        files = list(folder.glob(glob_pattern))
        if not files:
            continue
        latest = max(files, key=lambda p: p.stat().st_mtime)
        parsed = _parse_historical_export(latest, "")
        window_start = parsed['Occurred On'].min() if not parsed.empty else pd.NaT
        if pd.isna(window_start):
            continue
        if window_start <= day_start:
            return latest, True
        if best_partial_start is None or window_start < best_partial_start:
            best_partial, best_partial_start = latest, window_start
    return best_partial, False


# Current-alarm export per historical source - see _merge_still_active.
_CURRENT_EXPORT_FOR_SOURCE = {
    'MAE': (CURRENT_ALARMS_DIR, MAE_CURRENT_GLOB),
    'NetEco': (CURRENT_ALARMS_DIR, NETECO_CURRENT_GLOB),
    'NCE': (NCE_CURRENT_ALARMS_DIR, NCE_CURRENT_GLOB),
}


_ALARM_EVENT_KEY = ['Site', 'Name', 'Occurred On']


def _dedupe_alarm_events(df: pd.DataFrame) -> pd.DataFrame:
    """One row per alarm event (site/name/occurred time) across several
    overlapping exports, keeping a row that has a Cleared On over one that
    doesn't - distinct events at the same site (several outages in one
    day) all survive, since their Occurred On differs."""
    if df.empty:
        return df
    has_clear = df['Cleared On'].notna()
    return (df.assign(_has_clear=has_clear).sort_values('_has_clear', ascending=False)
            .drop_duplicates(subset=_ALARM_EVENT_KEY).drop(columns='_has_clear')
            .reset_index(drop=True))


def _merge_still_active(historical: pd.DataFrame, source: str):
    """(merged frame, live snapshot time or None).

    Historical exports only list CLEARED alarms (verified 2026-10-08:
    every NE Is Disconnected row in them has a Cleared On), so a site that
    went down on a given day and is still down now was missing from that
    day's report entirely - e.g. AJDB009/SLOG008/BGZ197 on 2026-10-07.
    Append the source's latest current-alarm export (Cleared On left
    empty = not cleared), dropping any alarm the historical export already
    has as cleared (same site/name/occurred time - it cleared after the
    live snapshot was taken). Those rows carry 'Live Snapshot' (the
    export's mtime) so _event_status can tell "still active per a fresh
    snapshot" from "was active in a stale one" - the latter is Unknown,
    not proof the site is still down."""
    base_dir, pattern = _CURRENT_EXPORT_FOR_SOURCE[source]
    current_path = _find_latest_export(base_dir, pattern)
    historical = historical.assign(**{'Live Snapshot': pd.NaT})
    if current_path is None:
        return historical, None
    snapshot_time = pd.Timestamp(datetime.fromtimestamp(current_path.stat().st_mtime))
    current = _parse_historical_export(current_path, source)
    if current.empty:
        return historical, snapshot_time
    current = current.assign(**{'Cleared On': pd.NaT, 'Live Snapshot': snapshot_time})
    if not historical.empty:
        seen = pd.MultiIndex.from_frame(historical[_ALARM_EVENT_KEY])
        current = current[~pd.MultiIndex.from_frame(current[_ALARM_EVENT_KEY]).isin(seen)]
        return pd.concat([historical, current], ignore_index=True), snapshot_time
    return current, snapshot_time


def _day_overlap_hours(df: pd.DataFrame, target_date: str) -> pd.DataFrame:
    """Clip each alarm's [Occurred On, Cleared On] interval to target_date's
    24h window and add an 'Overlap Hours' column - an alarm spanning
    several days only contributes the portion that actually fell on
    target_date, not its full multi-day duration. A still-open alarm (no
    Cleared On) is treated as ongoing through end-of-day, or "now" if
    target_date is today (so it can't claim hours that haven't happened
    yet). Rows with zero overlap (e.g. an alarm entirely on a different day
    but still present because of the export's rolling window) are dropped."""
    if df.empty or 'Occurred On' not in df.columns:
        df = df.copy()
        df['Overlap Hours'] = pd.Series(dtype=float)
        return df

    day_start = pd.Timestamp(target_date)
    day_end = day_start + pd.Timedelta(days=1)
    day_end_clip = min(day_end, pd.Timestamp.now())

    df = df.copy()
    effective_end = df['Cleared On'].fillna(day_end_clip)
    overlap_start = df['Occurred On'].clip(lower=day_start)
    overlap_end = effective_end.clip(upper=day_end)
    overlap_seconds = (overlap_end - overlap_start).dt.total_seconds().clip(lower=0)
    df['Overlap Hours'] = overlap_seconds / 3600.0
    return df[df['Occurred On'].notna() & (df['Overlap Hours'] > 0)].reset_index(drop=True)


def load_historical_alarms_for_date(target_date: str) -> Dict:
    """MAE/NetEco/NCE historical alarm rows whose interval overlaps
    target_date, each clipped to that day, from whichever raw historical
    export best covers that date. Each source is independently optional -
    one whose raw exports have already aged out of the sibling suite's
    retention window (a few days) is reported unavailable for that date
    rather than blocking the other two."""
    sources = {
        'MAE': (MAE_HISTORICAL_DIR, MAE_HISTORICAL_GLOB),
        'NetEco': (NETECO_HISTORICAL_DIR, NETECO_HISTORICAL_GLOB),
        'NCE': (NCE_HISTORICAL_DIR, NCE_HISTORICAL_GLOB),
    }
    result = {}
    for source, (base_dir, pattern) in sources.items():
        export_path, covers_full_day = _find_best_historical_export(base_dir, pattern, target_date)
        if export_path is None:
            result[source] = {'available': False, 'alarms': pd.DataFrame(), 'covers_full_day': False}
            continue
        raw = _parse_historical_export(export_path, source)
        # The full-day export can be days older than the newest one (it's
        # picked for reaching back to midnight), so an event from that day
        # which cleared after it was taken would look still-open there -
        # layer the newest export on top for the freshest Cleared On.
        newest_path = _find_latest_export(base_dir, pattern)
        if newest_path is not None and newest_path != export_path:
            raw = pd.concat([raw, _parse_historical_export(newest_path, source)], ignore_index=True)
        raw, snapshot_time = _merge_still_active(_dedupe_alarm_events(raw), source)
        result[source] = {
            'available': True,
            'alarms': _day_overlap_hours(raw, target_date),
            'export_file': export_path.name,
            'export_time': datetime.fromtimestamp(export_path.stat().st_mtime),
            'covers_full_day': covers_full_day,
            'live_snapshot_time': snapshot_time,
        }
    return result


def _reasons_for_site(source_df: pd.DataFrame, site: str) -> str:
    if source_df.empty or 'Site Code' not in source_df.columns:
        return ""
    rows = source_df[source_df['Site Code'] == site]
    return " | ".join(sorted(set(rows['Name'].dropna())))


def _nce_reason_for_site(nce_df: pd.DataFrame, site: str, ancestor_map: dict) -> str:
    """NCE (transmission) evidence for site: its own node first, then - if
    the site itself shows nothing - every FN/HUB node upstream of it
    (config/site_topology.csv via build_site_ancestor_map). Confirmed live
    2026-09-10: a single upstream FN (KUFRAGP) being down explained 8
    separate MAE-down sites at once that had zero NCE alarm on their own
    node - without this, those sites fell through to "Investigating" even
    though the transmission-side cause was already visible one hop up."""
    own = _reasons_for_site(nce_df, site)
    if own:
        return own
    for hub in sorted(ancestor_map.get(site, ())):
        hub_reason = _reasons_for_site(nce_df, hub)
        if hub_reason:
            return f"(via hub {hub}) {hub_reason}"
    return ""


def _is_power_reason(reason: str) -> bool:
    """A NetEco reason string only counts as Power evidence if it names an
    actual power fault (POWER_ALARM_TERMS) - not just any NetEco alarm at
    all. Confirmed live 2026-09-10: 28% of MAE 'NE Is Disconnected' sites
    had no NetEco evidence beyond NetEco's own 'Communication Between NMS
    And NE Is Abnormal' - the same link/power loss MAE already reports for
    that site, not independent power evidence - so bucketing any non-empty
    NetEco reason as Power over-attributes; those sites belong in
    Investigating instead."""
    return any(term in reason for term in POWER_ALARM_TERMS)


def _attach_alarm_reasons(per_site: pd.DataFrame, neteco: pd.DataFrame, nce: pd.DataFrame) -> pd.DataFrame:
    """Adds NetEco Reason (Power)/NCE Reason (Transmission)/Root Cause
    Bucket columns to per_site (which must have a 'Site' column) - shared
    by the daily historical report, the live disconnected-sites view, and
    the multi-day chronic-offender/downtime rollup, so the same hub-aware
    Power/Transmission bucketing logic (_is_power_reason, _nce_reason_for_site)
    answers "why is this site down" identically everywhere, not just in the
    one place it was first written."""
    ancestor_map = build_site_ancestor_map()
    per_site = per_site.copy()
    per_site['NetEco Reason (Power)'] = per_site['Site'].apply(lambda s: _reasons_for_site(neteco, s))
    per_site['NCE Reason (Transmission)'] = per_site['Site'].apply(
        lambda s: _nce_reason_for_site(nce, s, ancestor_map))
    per_site['Root Cause Bucket'] = per_site.apply(
        lambda r: "/".join(b for b, present in
                            (("Power", _is_power_reason(r['NetEco Reason (Power)'])),
                             ("Transmission", bool(r['NCE Reason (Transmission)']))) if present)
        or "Investigating", axis=1)
    return per_site


def _power_reason_and_time(neteco_df: pd.DataFrame, site: str):
    """(Power Reason (NOC), Mains Failure Time) for site, formatted to
    match the original sibling suite's Live_NOC_Report.xlsx labels: matched
    POWER_ALARM_TERMS joined in a fixed order (e.g. "Mains Failure - LLVD"),
    or "Check TX/Link (No Power Alarm)" when none of them fired for this
    site even if some other NetEco alarm (e.g. the ambiguous 'Communication
    Between NMS And NE Is Abnormal') did - see _is_power_reason. Mains
    Failure Time is specifically the Mains Failure alarm's own most recent
    Occurred timestamp - "-" if that alarm didn't fire even if BLVD/LLVD did."""
    if neteco_df.empty or 'Site Code' not in neteco_df.columns:
        return "Check TX/Link (No Power Alarm)", "-"
    rows = neteco_df[neteco_df['Site Code'] == site]
    if rows.empty:
        return "Check TX/Link (No Power Alarm)", "-"
    names = set(rows['Name'].dropna())
    matched = [term for term in POWER_ALARM_TERMS if any(term in n for n in names)]
    reason = " - ".join(matched) if matched else "Check TX/Link (No Power Alarm)"
    mains_time = "-"
    if 'Occurred On' in rows.columns:
        mains_rows = rows[rows['Name'].astype(str).str.contains('Mains Failure', na=False)]
        if not mains_rows.empty:
            latest = mains_rows['Occurred On'].max()
            if pd.notna(latest):
                mains_time = latest.strftime('%Y-%m-%d %H:%M:%S')
    return reason, mains_time


def _transmission_reason_and_time(nce_df: pd.DataFrame, site: str, ancestor_map: dict):
    """(Transmission Reason (NOC), NCE Last Occurred) for site - own
    transmission node first, then (annotated) its upstream FN/HUB, same
    fallback order as _nce_reason_for_site, but also returning the most
    recent Occurred timestamp from whichever alarm set was actually used."""
    if nce_df.empty or 'Site Code' not in nce_df.columns:
        return "-", "-"

    def _summarize(rows: pd.DataFrame):
        # A busy transmission node during a real outage can carry a dozen+
        # symptom alarms at once (BFD/BGP/LDP/interface flaps, etc.) - pipe-
        # joining all of them makes the cell unreadable. NCE's own "The
        # Device is offline" is the definitive down-signal (NCE's
        # equivalent of MAE's "NE Is Disconnected" - see _nce_reason_for_site),
        # so show just that when present; otherwise fall back to the first
        # few distinct symptom alarms rather than the full list.
        distinct = sorted(set(rows['Name'].dropna()))
        if "The Device is offline" in distinct:
            names = "The Device is offline"
        else:
            names = " - ".join(distinct[:3]) + (" - ..." if len(distinct) > 3 else "")
        latest_str = "-"
        if 'Occurred On' in rows.columns:
            latest = rows['Occurred On'].max()
            if pd.notna(latest):
                latest_str = latest.strftime('%Y-%m-%d %H:%M:%S')
        return names, latest_str

    own = nce_df[nce_df['Site Code'] == site]
    if not own.empty:
        return _summarize(own)
    for hub in sorted(ancestor_map.get(site, ())):
        hub_rows = nce_df[nce_df['Site Code'] == hub]
        if not hub_rows.empty:
            names, latest_str = _summarize(hub_rows)
            return f"(via hub {hub}) {names}", latest_str
    return "-", "-"


_RICH_ALARM_TABLE_COLUMNS = [
    'Site Name', 'MO Name', 'Alarm Name', 'Last Occurred', 'Down Hours',
    'Power Reason (NOC)', 'Mains Failure Time', 'Transmission Reason (NOC)', 'NCE Last Occurred',
]

# Extra per-site columns the historical (day) view adds after Last Occurred
# - see _build_rich_alarm_table.
_DAY_STATUS_COLUMNS = ['Cleared On', 'Status', 'Outage Class', 'Outage Events']

# An outage event lasting at least this long (occurred -> cleared, or ->
# now if still open) is classed "Long-Term Outage" rather than "Normal".
# Judgment call, not from any upstream config: a week is past any
# ordinary power/TX restoration cycle (the longest normal events seen
# 2026-09/10 were ~4 days, e.g. AJDB007), while the sites that motivated
# this (ECV001 since 2026-05, BYDA001 since 2026-03, BGZ162/UECV001 since
# late 2025) are months in - they look decommissioned or parked, and at
# 24h/day each they otherwise add ~96 hours to every day's Total Downtime.
LONG_TERM_OUTAGE_DAYS = 7

STATUS_CLEARED = "Cleared"
STATUS_NOT_CLEARED = "Not Cleared"
STATUS_UNKNOWN = "Unknown"
CLASS_NORMAL = "Normal"
CLASS_LONG_TERM = "Long-Term Outage"


def _fmt_ts(ts) -> str:
    return ts.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(ts) else "-"


def _event_status(cleared_on, live_snapshot, now: pd.Timestamp) -> str:
    """Cleared when the event has its own Cleared On. Otherwise it only
    came from a live current-alarm snapshot (_merge_still_active): Not
    Cleared if that snapshot is fresh (LIVE_STALE_MINUTES), Unknown if
    it's old - a stale snapshot showing the alarm is not proof the site is
    still down now."""
    if pd.notna(cleared_on):
        return STATUS_CLEARED
    if pd.notna(live_snapshot) and (now - live_snapshot) <= pd.Timedelta(minutes=LIVE_STALE_MINUTES):
        return STATUS_NOT_CLEARED
    return STATUS_UNKNOWN


def _build_outage_events(down: pd.DataFrame, now: pd.Timestamp) -> pd.DataFrame:
    """Every NE Is Disconnected event in down (already day-clipped by
    _day_overlap_hours), one row each - a site with three separate outages
    that day keeps all three here, while the per-site table rolls them up."""
    cols = ['Site Name', 'MO Name', 'Occurred On', 'Cleared On', 'Status', 'Outage Class',
            'Hours This Day', 'Event Duration (Hours)']
    if down.empty:
        return pd.DataFrame(columns=cols)
    live_snap = down['Live Snapshot'] if 'Live Snapshot' in down.columns else pd.Series(pd.NaT, index=down.index)
    end = down['Cleared On'].fillna(now)
    duration_h = ((end - down['Occurred On']).dt.total_seconds() / 3600.0).clip(lower=0)
    events = pd.DataFrame({
        'Site Name': down['Site Code'],
        'MO Name': down['Site'].fillna(down['Site Code']),
        'Occurred On': down['Occurred On'],
        'Cleared On': down['Cleared On'],
        'Status': [_event_status(c, s, now) for c, s in zip(down['Cleared On'], live_snap)],
        'Outage Class': (duration_h >= LONG_TERM_OUTAGE_DAYS * 24).map({True: CLASS_LONG_TERM, False: CLASS_NORMAL}),
        'Hours This Day': down['Overlap Hours'].round(2) if 'Overlap Hours' in down.columns else 0.0,
        'Event Duration (Hours)': duration_h.round(2),
    })
    return events.sort_values(['Site Name', 'Occurred On']).reset_index(drop=True)[cols]


def _site_status(statuses: pd.Series) -> str:
    """Site-level rollup of its events' statuses: any still-open event
    makes the site Not Cleared; failing that, any unverifiable one makes it
    Unknown; only all-cleared is Cleared."""
    values = set(statuses)
    if STATUS_NOT_CLEARED in values:
        return STATUS_NOT_CLEARED
    if STATUS_UNKNOWN in values:
        return STATUS_UNKNOWN
    return STATUS_CLEARED


def _build_rich_alarm_table(mae: pd.DataFrame, neteco: pd.DataFrame, nce: pd.DataFrame,
                             hours_mode: str = 'day', events: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Per-disconnected-site table matching the original sibling suite's
    Live_NOC_Report.xlsx column layout (Site Name/MO Name/Alarm Name/Last
    Occurred/Power Reason (NOC)/Mains Failure Time/Transmission Reason
    (NOC)/NCE Last Occurred), plus a Down Hours column - built natively
    from raw MAE/NetEco/NCE exports (mae/neteco/nce are whatever
    _parse_historical_export produced; works for current or historical
    alike, since both now populate the same Site/Site Code/Name/Occurred On
    schema). hours_mode='day' sums each site's already-day-clipped
    'Overlap Hours' (for one historical calendar day); hours_mode='live'
    instead reports elapsed time since the alarm's own Last Occurred, since
    a live snapshot has no day-clipped Overlap Hours column at all.

    hours_mode='day' also adds _DAY_STATUS_COLUMNS, rolled up from events
    (_build_outage_events): Status via _site_status, Cleared On = latest
    clear time when Cleared, Outage Class = Long-Term if any of the site's
    events is, and how many events the site had that day. The live view
    adds only Outage Class - every row there is uncleared by definition."""
    columns = list(_RICH_ALARM_TABLE_COLUMNS)
    if hours_mode == 'day':
        columns[columns.index('Last Occurred') + 1:1] = _DAY_STATUS_COLUMNS
    else:
        columns.insert(columns.index('Last Occurred') + 1, 'Outage Class')
    if mae.empty or 'Name' not in mae.columns:
        return pd.DataFrame(columns=columns)

    down = mae[mae['Name'] == SITE_DOWN_ALARM_NAME].dropna(subset=['Site Code'])
    if down.empty:
        return pd.DataFrame(columns=columns)

    hours_by_site = None
    if hours_mode == 'day' and 'Overlap Hours' in down.columns:
        hours_by_site = down.groupby('Site Code')['Overlap Hours'].sum()
    events_by_site = None
    if hours_mode == 'day' and events is not None and not events.empty:
        events_by_site = {site: grp for site, grp in events.groupby('Site Name')}

    if 'Occurred On' in down.columns:
        down = down.sort_values('Occurred On', ascending=False).drop_duplicates(subset=['Site Code'])
    else:
        down = down.drop_duplicates(subset=['Site Code'])

    ancestor_map = build_site_ancestor_map()
    now = datetime.now()
    rows = []
    for _, r in down.iterrows():
        site = r['Site Code']
        occurred = r.get('Occurred On')
        has_occurred = pd.notna(occurred) if occurred is not None else False
        occurred_str = occurred.strftime('%Y-%m-%d %H:%M:%S') if has_occurred else "-"
        if hours_by_site is not None:
            down_hours = float(hours_by_site.get(site, 0.0))
        elif has_occurred:
            down_hours = max((now - occurred).total_seconds() / 3600.0, 0.0)
        else:
            down_hours = 0.0
        power_reason, mains_time = _power_reason_and_time(neteco, site)
        trans_reason, nce_time = _transmission_reason_and_time(nce, site, ancestor_map)
        row = {
            'Site Name': site,
            'MO Name': r.get('Site') or site,
            'Alarm Name': r['Name'],
            'Last Occurred': occurred_str,
            'Down Hours': round(down_hours, 2),
            'Power Reason (NOC)': power_reason,
            'Mains Failure Time': mains_time,
            'Transmission Reason (NOC)': trans_reason,
            'NCE Last Occurred': nce_time,
        }
        if hours_mode == 'day' and events_by_site is not None:
            site_events = events_by_site.get(site)
            status = _site_status(site_events['Status'])
            row['Status'] = status
            row['Cleared On'] = _fmt_ts(site_events['Cleared On'].max()) if status == STATUS_CLEARED else "-"
            row['Outage Class'] = (CLASS_LONG_TERM if (site_events['Outage Class'] == CLASS_LONG_TERM).any()
                                   else CLASS_NORMAL)
            row['Outage Events'] = len(site_events)
        elif hours_mode != 'day':
            row['Outage Class'] = (CLASS_LONG_TERM if down_hours >= LONG_TERM_OUTAGE_DAYS * 24
                                   else CLASS_NORMAL)
        rows.append(row)
    return pd.DataFrame(rows, columns=columns).sort_values(
        'Last Occurred', ascending=False).reset_index(drop=True)


def noc_alarm_kpis(sites: pd.DataFrame, exclude_long_term: bool = False) -> dict:
    """The Daily NOC Alarm Analysis headline tiles, derived from the
    per-site table itself so they can never drift out of sync with what
    the table shows. Counts are distinct sites - identical to row counts
    for one day, and "sites affected at least once" over a multi-day
    stack. exclude_long_term drops Long-Term Outage sites first (they stay
    in the table/exports; this is only for operational KPIs)."""
    keys = ('ne_disconnected', 'total_down_hours', 'mains_failure', 'nce_transmission',
            'cleared', 'not_cleared', 'unknown', 'long_term')
    if sites is None or sites.empty:
        return dict.fromkeys(keys, 0)
    long_term_mask = (sites['Outage Class'] == CLASS_LONG_TERM) if 'Outage Class' in sites.columns \
        else pd.Series(False, index=sites.index)
    long_term = int(sites.loc[long_term_mask, 'Site Name'].nunique())
    if exclude_long_term:
        sites = sites[~long_term_mask]

    def distinct(mask):
        return int(sites.loc[mask, 'Site Name'].nunique())

    status = sites['Status'] if 'Status' in sites.columns else pd.Series("", index=sites.index)
    return {
        'ne_disconnected': int(sites['Site Name'].nunique()),
        'total_down_hours': float(sites['Down Hours'].sum()),
        'mains_failure': distinct(sites['Mains Failure Time'] != '-'),
        'nce_transmission': distinct(sites['Transmission Reason (NOC)'] != '-'),
        'cleared': distinct(status == STATUS_CLEARED),
        'not_cleared': distinct(status == STATUS_NOT_CLEARED),
        'unknown': distinct(status == STATUS_UNKNOWN),
        'long_term': long_term,
    }


def _alarm_table_metrics(table: pd.DataFrame) -> dict:
    """Live view's 3 tiles - kept as its own name since the live section
    only ever showed these three."""
    k = noc_alarm_kpis(table)
    return {key: k[key] for key in ('ne_disconnected', 'mains_failure', 'nce_transmission')}


def _data_quality_row(target_date: str, by_source: Dict) -> dict:
    """One Data Quality row for target_date: overall coverage plus each
    source's export file/time and whether it spans the whole day."""
    partial = [src for src, info in by_source.items()
               if not (info['available'] and info.get('covers_full_day'))]
    row = {'Date': target_date, 'Data Coverage': f"Partial: {', '.join(partial)}" if partial else "Full"}
    for src, info in by_source.items():
        row[f'{src} Export'] = info.get('export_file') or "missing"
        row[f'{src} Export Time'] = _fmt_ts(pd.Timestamp(info['export_time'])) if info.get('export_time') else "-"
        row[f'{src} Full Day'] = "Yes" if info['available'] and info.get('covers_full_day') else "No"
    snap = by_source['MAE'].get('live_snapshot_time')
    row['Live Snapshot (MAE)'] = _fmt_ts(snap) if snap is not None else "missing"
    return row


def build_daily_noc_alarm_report(target_date: str) -> Dict:
    """Per-day NOC alarm analysis for target_date: how many sites went
    down (MAE's SITE_DOWN_ALARM_NAME), total summed downtime hours across
    them that day, and - for each down site - the NetEco (power) and NCE
    (transmission) alarm evidence for that same day merged into one row.
    Mirrors the sibling suite's own current-alarm triage logic
    (enhanced_noc_analysis.build_triage's per-site evidence-joining) but
    computed from raw historical exports for one specific calendar day
    instead of "right now." 'outage_events' keeps every individual
    NE Is Disconnected event behind the per-site rows."""
    by_source = load_historical_alarms_for_date(target_date)
    mae, neteco, nce = by_source['MAE']['alarms'], by_source['NetEco']['alarms'], by_source['NCE']['alarms']

    down = mae[mae['Name'] == SITE_DOWN_ALARM_NAME].dropna(subset=['Site Code']) \
        if not mae.empty and 'Name' in mae.columns else pd.DataFrame()
    events = _build_outage_events(down, pd.Timestamp.now())
    down_summary = _build_rich_alarm_table(mae, neteco, nce, hours_mode='day', events=events)
    down_summary.insert(0, 'Date', target_date)
    events.insert(0, 'Date', target_date)
    metrics = noc_alarm_kpis(down_summary)

    neteco_display = neteco.rename(columns={'Site Code': 'Site Canonical'}) if not neteco.empty else neteco
    nce_display = nce.rename(columns={'Site Code': 'Site Canonical'}) if not nce.empty else nce

    return {
        'date': target_date,
        'metrics': metrics,
        'available': any(info['available'] for info in by_source.values()),
        'sites_down': metrics['ne_disconnected'],
        'total_down_hours': metrics['total_down_hours'],
        'down_sites_summary': down_summary,
        'outage_events': events,
        'neteco_alarms': neteco_display,
        'nce_alarms': nce_display,
        'data_quality': _data_quality_row(target_date, by_source),
        'sources': {src: {'available': info['available'], 'export_file': info.get('export_file'),
                           'export_time': info.get('export_time'),
                           'covers_full_day': info.get('covers_full_day', False)}
                    for src, info in by_source.items()},
    }


def available_daily_noc_alarm_dates() -> List[str]:
    """Every calendar day from the earliest alarm the oldest surviving MAE
    historical export reaches back to, through today - the range the
    dashboard's date picker allows. Based on file content, not folder
    names: retention purges old exports but leaves their dated folders
    behind empty (2026-08-25..09-15 as of 2026-10-08). Days in a scraping
    gap stay in the range; they just come back flagged Partial."""
    base_path = Path(MAE_HISTORICAL_DIR)
    if not base_path.exists():
        return []
    for folder in sorted(p for p in base_path.iterdir() if p.is_dir()):
        files = list(folder.glob(MAE_HISTORICAL_GLOB))
        if not files:
            continue
        parsed = _parse_historical_export(max(files, key=lambda p: p.stat().st_mtime), 'MAE')
        if parsed.empty or parsed['Occurred On'].isna().all():
            continue
        oldest = parsed['Occurred On'].min().date()
        today = datetime.now().date()
        return [(oldest + timedelta(days=i)).isoformat() for i in range((today - oldest).days + 1)]
    return []


def build_daily_noc_alarm_range(start_date: str, end_date: str) -> Dict:
    """build_daily_noc_alarm_report for every day start_date..end_date
    (inclusive), stacked: 'sites' (per-site rows, Date column first),
    'events' (every outage event), 'data_quality' (one row per day - also
    the list of days the range covers, so a day with zero outages still
    gets a summary row). Days with no MAE export at all are skipped."""
    site_frames, event_frames, dq_rows = [], [], []
    for day in pd.date_range(start_date, end_date, freq='D').strftime('%Y-%m-%d'):
        report = build_daily_noc_alarm_report(day)
        if not report['sources']['MAE']['available']:
            continue
        dq_rows.append(report['data_quality'])
        if not report['down_sites_summary'].empty:
            site_frames.append(report['down_sites_summary'])
        if not report['outage_events'].empty:
            event_frames.append(report['outage_events'])
    return {
        'sites': pd.concat(site_frames, ignore_index=True) if site_frames else pd.DataFrame(),
        'events': pd.concat(event_frames, ignore_index=True) if event_frames else pd.DataFrame(),
        'data_quality': pd.DataFrame(dq_rows),
    }


def summarize_noc_alarm_days(sites: pd.DataFrame, data_quality: pd.DataFrame,
                             exclude_long_term: bool = False) -> pd.DataFrame:
    """One row per day in data_quality: noc_alarm_kpis over that day's
    rows of sites (already filtered by the caller, if it filters), plus the
    day's Data Coverage. Newest day first."""
    rows = []
    for _, dq in data_quality.iterrows():
        day_sites = sites[sites['Date'] == dq['Date']] if not sites.empty else sites
        k = noc_alarm_kpis(day_sites, exclude_long_term=exclude_long_term)
        rows.append({
            'Date': dq['Date'],
            'Count of NE Is Disconnected': k['ne_disconnected'],
            'Total Downtime (Hours)': round(k['total_down_hours'], 2),
            'Count of Mains Failure': k['mains_failure'],
            'Count of NCE Transmission Alarm Sites': k['nce_transmission'],
            'Cleared': k['cleared'],
            'Not Cleared': k['not_cleared'],
            'Unknown': k['unknown'],
            'Long-Term Outage Sites': k['long_term'],
            'Data Coverage': dq['Data Coverage'],
        })
    summary = pd.DataFrame(rows)
    return summary.sort_values('Date', ascending=False).reset_index(drop=True) if not summary.empty else summary


# Power Reason filter choices -> substring each matches in 'Power Reason (NOC)'.
POWER_REASON_FILTERS = {
    'Mains Failure': 'Mains Failure', 'BLVD': 'BLVD', 'LLVD': 'LLVD',
    'No Power Alarm': 'No Power Alarm',
}
TRANSMISSION_FILTERS = ('Own NCE alarm', 'Via upstream hub', 'No NCE alarm')


def apply_noc_alarm_filters(sites: pd.DataFrame, data_quality: pd.DataFrame, site_codes=None,
                            statuses=None, outage_classes=None, power_reasons=None,
                            transmission=None, min_hours=None, max_hours=None,
                            coverage=None) -> pd.DataFrame:
    """Filter the per-site table. Each argument left None/empty means "no
    filter on that field"; multi-value ones OR within a field and AND
    across fields. site_codes match as case-insensitive substrings
    (so "BGZ" catches every Benghazi site). coverage is a subset of
    {'Full', 'Partial'}, matched against each row's day in data_quality."""
    if sites.empty:
        return sites
    mask = pd.Series(True, index=sites.index)
    if site_codes:
        pattern = '|'.join(re.escape(c.strip()) for c in site_codes if c.strip())
        if pattern:
            mask &= sites['Site Name'].astype(str).str.contains(pattern, case=False, na=False)
    if statuses:
        mask &= sites['Status'].isin(statuses)
    if outage_classes:
        mask &= sites['Outage Class'].isin(outage_classes)
    if power_reasons:
        reason = sites['Power Reason (NOC)'].astype(str)
        mask &= pd.concat([reason.str.contains(POWER_REASON_FILTERS[p], regex=False)
                           for p in power_reasons], axis=1).any(axis=1)
    if transmission:
        trans = sites['Transmission Reason (NOC)'].astype(str)
        kinds = pd.Series('Own NCE alarm', index=sites.index)
        kinds[trans.str.startswith('(via hub')] = 'Via upstream hub'
        kinds[trans == '-'] = 'No NCE alarm'
        mask &= kinds.isin(transmission)
    if min_hours is not None:
        mask &= sites['Down Hours'] >= min_hours
    if max_hours is not None:
        mask &= sites['Down Hours'] <= max_hours
    if coverage and not data_quality.empty:
        day_cov = data_quality.set_index('Date')['Data Coverage'].map(
            lambda c: 'Full' if c == 'Full' else 'Partial')
        mask &= sites['Date'].map(day_cov).isin(coverage)
    return sites[mask].reset_index(drop=True)


def active_outages(sites: pd.DataFrame) -> pd.DataFrame:
    """Sites whose most recent day in sites is still Not Cleared/Unknown -
    one row per site (its latest day), longest-down first."""
    if sites.empty:
        return sites
    latest = sites.sort_values('Date').drop_duplicates(subset=['Site Name'], keep='last')
    active = latest[latest['Status'].isin([STATUS_NOT_CLEARED, STATUS_UNKNOWN])]
    return active.sort_values('Last Occurred').reset_index(drop=True)


_STATUS_FILLS = {STATUS_CLEARED: 'C6EFCE', STATUS_NOT_CLEARED: 'FFC7CE', STATUS_UNKNOWN: 'FFEB9C',
                 CLASS_LONG_TERM: 'D9D2E9'}


def build_noc_alarm_workbook(sheets: List, info: List) -> bytes:
    """Formatted .xlsx for the Daily NOC Alarm Analysis export: an Info
    sheet (info = [(field, value), ...] - report scope and active filters),
    then one sheet per (name, DataFrame) in sheets, each with a styled
    frozen header, autofilter, auto-sized columns, and Status/Outage Class
    cells colour-coded the same way as the dashboard."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    import io

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(bold=True, color='FFFFFF')
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        pd.DataFrame(info, columns=['Field', 'Value']).to_excel(writer, sheet_name='Info', index=False)
        for name, df in sheets:
            out = df if df is not None and not df.empty else pd.DataFrame([{'Note': 'No rows for this selection'}])
            out.to_excel(writer, sheet_name=name[:31], index=False)
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.fill, cell.font = header_fill, header_font
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            ws.freeze_panes = 'A2'
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for col_idx, col_cells in enumerate(ws.columns, start=1):
                header = str(col_cells[0].value)
                width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width + 2, 10), 60)
                if header in ('Status', 'Outage Class'):
                    for c in col_cells[1:]:
                        fill = _STATUS_FILLS.get(c.value)
                        if fill:
                            c.fill = PatternFill('solid', fgColor=fill)
    return buf.getvalue()


def build_live_disconnected_sites() -> Dict:
    """Native replacement for the old sibling-suite-dependent
    load_live_alarm_summary(): "which sites are down right now and why,"
    computed directly from each source's own latest current-alarm export
    (CURRENT_ALARMS_DIR/NCE_CURRENT_ALARMS_DIR) instead of Live_NOC_Report.xlsx
    - which was built by a separate Merge Reports service that has no
    guarantee of still running now that scraping happens in this project.
    Uses the identical site-extraction/reason/bucketing logic as
    build_daily_noc_alarm_report (_attach_alarm_reasons), just against
    "right now" instead of one historical day, so the two views never
    disagree on what counts as a power vs. transmission cause."""
    mae_path = _find_latest_export(CURRENT_ALARMS_DIR, MAE_CURRENT_GLOB)
    if mae_path is None:
        return {'loaded': False}
    neteco_path = _find_latest_export(CURRENT_ALARMS_DIR, NETECO_CURRENT_GLOB)
    nce_path = _find_latest_export(NCE_CURRENT_ALARMS_DIR, NCE_CURRENT_GLOB)

    mae = _parse_historical_export(mae_path, 'MAE')
    neteco = _parse_historical_export(neteco_path, 'NetEco') if neteco_path else pd.DataFrame()
    nce = _parse_historical_export(nce_path, 'NCE') if nce_path else pd.DataFrame()

    alarms_df = _build_rich_alarm_table(mae, neteco, nce, hours_mode='live')
    metrics = _alarm_table_metrics(alarms_df)

    mtimes = [p.stat().st_mtime for p in (mae_path, neteco_path, nce_path) if p]
    oldest_mtime = min(mtimes)
    age_minutes = (datetime.now().timestamp() - oldest_mtime) / 60.0
    return {
        'loaded': True,
        'alarms': alarms_df,
        'metrics': metrics,
        'age_minutes': age_minutes,
        'is_stale': age_minutes > LIVE_STALE_MINUTES,
        'last_updated': datetime.fromtimestamp(oldest_mtime),
    }


def build_historical_insights_native(lookback_days: int = 14) -> Dict:
    """Native replacement for the old sibling-suite-dependent
    load_historical_insights(): chronic-offender/site-downtime/daily-trend/
    category rollups computed by running build_daily_noc_alarm_report's
    same per-day logic (_attach_alarm_reasons) across a rolling window of
    raw historical exports, instead of reading historical_insights_snapshot.json
    - built by a separate historical_noc_analysis.py ledger process with no
    guarantee of still running now that scraping happens in this project
    (and which had a history of crashing itself with MemoryError - see
    module docstring). Days with no raw export available (aged out of
    retention, or simply not scraped yet) are silently skipped rather than
    failing the whole window."""
    per_day_rows = []
    today = datetime.now().date()
    for offset in range(lookback_days):
        day = (today - timedelta(days=offset)).isoformat()
        by_source = load_historical_alarms_for_date(day)
        mae, neteco, nce = by_source['MAE']['alarms'], by_source['NetEco']['alarms'], by_source['NCE']['alarms']
        if mae.empty or 'Name' not in mae.columns:
            continue
        down = mae[mae['Name'] == SITE_DOWN_ALARM_NAME].dropna(subset=['Site Code'])
        if down.empty:
            continue
        per_site = down.groupby('Site Code').agg(
            Occurrences=('Name', 'size'), **{'Down Hours': ('Overlap Hours', 'sum')},
        ).reset_index().rename(columns={'Site Code': 'Site'})
        per_site = _attach_alarm_reasons(per_site, neteco, nce)
        per_site.insert(0, 'Date', day)
        per_day_rows.append(per_site)

    if not per_day_rows:
        return {'loaded': False}

    all_days = pd.concat(per_day_rows, ignore_index=True)

    chronic_offenders = all_days.groupby('Site').agg(
        Occurrences=('Date', 'nunique'), **{'Total Down Hours': ('Down Hours', 'sum')},
    ).reset_index().sort_values('Occurrences', ascending=False).reset_index(drop=True)

    site_downtime = all_days.groupby('Site')['Down Hours'].sum().reset_index()
    site_downtime['Total Outage Minutes'] = site_downtime['Down Hours'] * 60
    site_downtime = site_downtime.drop(columns=['Down Hours']).sort_values(
        'Total Outage Minutes', ascending=False).reset_index(drop=True)

    daily_trend = all_days.groupby('Date').agg(
        **{'Sites Down': ('Site', 'nunique'), 'Total Down Hours': ('Down Hours', 'sum')},
    ).reset_index().sort_values('Date').reset_index(drop=True)

    category_rollup = all_days.groupby('Root Cause Bucket').agg(
        **{'Site-Days': ('Site', 'size')},
    ).reset_index().sort_values('Site-Days', ascending=False).reset_index(drop=True)

    return {
        'loaded': True,
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'age_minutes': 0.0,
        'is_stale': False,
        'chronic_offenders': chronic_offenders,
        'site_downtime': site_downtime,
        'daily_trend': daily_trend,
        'category_rollup': category_rollup,
    }


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    live_summary = build_live_disconnected_sites()
    print("Live loaded:", live_summary.get('loaded'), "| age (min):", live_summary.get('age_minutes'))
    if live_summary.get('loaded'):
        print(live_summary['alarms'].head(10).to_string())

    insights = build_historical_insights_native()
    print("\nHistorical loaded:", insights.get('loaded'), "| generated_at:", insights.get('generated_at'))
    if insights.get('loaded'):
        print("Chronic offenders:", len(insights['chronic_offenders']))
        print("Site downtime:", len(insights['site_downtime']))
