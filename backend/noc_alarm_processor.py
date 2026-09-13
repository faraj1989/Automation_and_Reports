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


def _find_best_historical_export(base_dir: str, glob_pattern: str, target_date: str) -> Optional[Path]:
    """The rolling-window historical exports mean the MOST RECENT export
    that still covers target_date has the most complete Cleared-On picture
    for that day's alarms (an alarm from target_date that only cleared the
    next day needs an export taken after it cleared to show that). Search
    target_date's own dated folder and every later one up to today, and
    return the single most-recently-modified matching file found - or None
    if the raw export has already aged out of the retention window."""
    base_path = Path(base_dir)
    if not base_path.exists():
        return None
    try:
        target = datetime.strptime(target_date, "%Y-%m-%d").date()
    except ValueError:
        return None

    candidates = []
    for entry in base_path.iterdir():
        if not entry.is_dir() or len(entry.name) != 10:
            continue
        try:
            folder_date = datetime.strptime(entry.name, "%Y-%m-%d").date()
        except ValueError:
            continue
        if folder_date >= target:
            candidates.extend(entry.glob(glob_pattern))
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


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
        export_path = _find_best_historical_export(base_dir, pattern, target_date)
        if export_path is None:
            result[source] = {'available': False, 'alarms': pd.DataFrame()}
            continue
        raw = _parse_historical_export(export_path, source)
        result[source] = {
            'available': True,
            'alarms': _day_overlap_hours(raw, target_date),
            'export_file': export_path.name,
            'export_time': datetime.fromtimestamp(export_path.stat().st_mtime),
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


def _build_rich_alarm_table(mae: pd.DataFrame, neteco: pd.DataFrame, nce: pd.DataFrame,
                             hours_mode: str = 'day') -> pd.DataFrame:
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
    a live snapshot has no day-clipped Overlap Hours column at all."""
    if mae.empty or 'Name' not in mae.columns:
        return pd.DataFrame(columns=_RICH_ALARM_TABLE_COLUMNS)

    down = mae[mae['Name'] == SITE_DOWN_ALARM_NAME].dropna(subset=['Site Code'])
    if down.empty:
        return pd.DataFrame(columns=_RICH_ALARM_TABLE_COLUMNS)

    hours_by_site = None
    if hours_mode == 'day' and 'Overlap Hours' in down.columns:
        hours_by_site = down.groupby('Site Code')['Overlap Hours'].sum()

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
        rows.append({
            'Site Name': site,
            'MO Name': r.get('Site') or site,
            'Alarm Name': r['Name'],
            'Last Occurred': occurred_str,
            'Down Hours': round(down_hours, 2),
            'Power Reason (NOC)': power_reason,
            'Mains Failure Time': mains_time,
            'Transmission Reason (NOC)': trans_reason,
            'NCE Last Occurred': nce_time,
        })
    return pd.DataFrame(rows, columns=_RICH_ALARM_TABLE_COLUMNS).sort_values(
        'Last Occurred', ascending=False).reset_index(drop=True)


def _alarm_table_metrics(table: pd.DataFrame) -> dict:
    """The 3 headline counts from the original Live_NOC_Report.xlsx metric
    tiles, derived from the rich alarm table itself so they can never drift
    out of sync with what the table actually shows."""
    if table.empty:
        return {'ne_disconnected': 0, 'mains_failure': 0, 'nce_transmission': 0}
    return {
        'ne_disconnected': len(table),
        'mains_failure': int((table['Mains Failure Time'] != '-').sum()),
        'nce_transmission': int((table['Transmission Reason (NOC)'] != '-').sum()),
    }


def build_daily_noc_alarm_report(target_date: str) -> Dict:
    """Per-day NOC alarm analysis for target_date: how many sites went
    down (MAE's SITE_DOWN_ALARM_NAME), total summed downtime hours across
    them that day, and - for each down site - the NetEco (power) and NCE
    (transmission) alarm evidence for that same day merged into one row.
    Mirrors the sibling suite's own current-alarm triage logic
    (enhanced_noc_analysis.build_triage's per-site evidence-joining) but
    computed from raw historical exports for one specific calendar day
    instead of "right now."""
    by_source = load_historical_alarms_for_date(target_date)
    mae, neteco, nce = by_source['MAE']['alarms'], by_source['NetEco']['alarms'], by_source['NCE']['alarms']

    down_summary = _build_rich_alarm_table(mae, neteco, nce, hours_mode='day')
    sites_down = len(down_summary)
    total_down_hours = float(down_summary['Down Hours'].sum()) if not down_summary.empty else 0.0
    if not down_summary.empty:
        down_summary.insert(0, 'Date', target_date)
    metrics = _alarm_table_metrics(down_summary)

    neteco_display = neteco.rename(columns={'Site Code': 'Site Canonical'}) if not neteco.empty else neteco
    nce_display = nce.rename(columns={'Site Code': 'Site Canonical'}) if not nce.empty else nce

    return {
        'date': target_date,
        'metrics': metrics,
        'available': any(info['available'] for info in by_source.values()),
        'sites_down': sites_down,
        'total_down_hours': total_down_hours,
        'down_sites_summary': down_summary,
        'neteco_alarms': neteco_display,
        'nce_alarms': nce_display,
        'sources': {src: {'available': info['available'], 'export_file': info.get('export_file'),
                           'export_time': info.get('export_time')}
                    for src, info in by_source.items()},
    }


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
