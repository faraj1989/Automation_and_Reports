#!/usr/bin/env python3
"""
Libyana NPM - EPT (Engineering Parameters Table) Manager

Loads, validates, and exports the manually-maintained EPT reference file
(config/Libyana MS EPT_*-Whole Network.xlsx) - per-cell physical/RF
engineering parameters (coordinates, azimuth, tilt, antenna height,
PCI/PSC, band, etc.), one sheet per technology (GSM/UMTS/LTE).

Unlike the daily KPI pipeline, EPT is hand-maintained by RF engineers and
not generated from the system, so it can contain typos/stale rows -
validate_ept() cross-checks it against the authoritative per-cell KPI
sheets in output/csv/ to surface those.
"""

import os
import glob
import math
import logging
from datetime import datetime
from typing import Dict, List, Optional
from xml.sax.saxutils import escape

import pandas as pd

logger = logging.getLogger(__name__)

EPT_DIR = "config"
EPT_FILE_PATTERN = "Libyana MS EPT_*-Whole Network.xlsx"
EPT_SHEETS = ['GSM', 'UMTS', 'LTE']

# System (KPI pipeline) cell-level sheets used to cross-validate EPT against
# what's actually reporting in production.
SYSTEM_CELL_SHEETS = {'GSM': '2G_Cell_CSBH', 'UMTS': '3G_Cell_CSBH', 'LTE': '4G_Cell_BH'}

# KML icon colors per technology, KML format is aabbggrr (alpha, blue, green, red)
TECH_KML_COLOR = {'GSM': 'ffff0000', 'UMTS': 'ff00a5ff', 'LTE': 'ff00ff00'}


def find_ept_file(ept_dir: str = EPT_DIR) -> Optional[str]:
    """Latest EPT workbook by filename (there's normally only one)."""
    matches = sorted(glob.glob(os.path.join(ept_dir, EPT_FILE_PATTERN)))
    return matches[-1] if matches else None


def get_ept_last_updated(ept_dir: str = EPT_DIR) -> Optional[datetime]:
    path = find_ept_file(ept_dir)
    if not path:
        return None
    return datetime.fromtimestamp(os.path.getmtime(path))


def load_ept(tech: str, ept_dir: str = EPT_DIR) -> Optional[pd.DataFrame]:
    path = find_ept_file(ept_dir)
    if not path or tech not in EPT_SHEETS:
        return None
    try:
        return pd.read_excel(path, sheet_name=tech)
    except Exception as e:
        logger.warning(f"Could not read EPT sheet {tech}: {e}")
        return None


def load_ept_all(ept_dir: str = EPT_DIR) -> Dict[str, pd.DataFrame]:
    result = {}
    for tech in EPT_SHEETS:
        df = load_ept(tech, ept_dir)
        if df is not None:
            result[tech] = df
    return result


def get_ept_file_bytes(ept_dir: str = EPT_DIR) -> Optional[bytes]:
    """Raw bytes of the current EPT workbook, for a straight download button
    (no reformatting - RF engineers get exactly what's on disk)."""
    path = find_ept_file(ept_dir)
    if not path:
        return None
    with open(path, 'rb') as f:
        return f.read()


def validate_ept(csv_folder: str = "output/csv", ept_dir: str = EPT_DIR) -> pd.DataFrame:
    """Cross-check EPT against the system's authoritative cell-level KPI
    sheets. Returns one summary row per technology."""
    rows = []
    for tech, sheet in SYSTEM_CELL_SHEETS.items():
        sys_path = os.path.join(csv_folder, f"{sheet}.csv")
        sys_names = set()
        if os.path.exists(sys_path):
            try:
                sys_names = set(pd.read_csv(sys_path)['Cell Name'].dropna().unique())
            except Exception:
                pass

        ept_df = load_ept(tech, ept_dir)
        if ept_df is None or 'Cell Name' not in ept_df.columns:
            rows.append({'Technology': tech, 'EPT Rows': 0, 'EPT Duplicate Cell Names': 0,
                         'System Cells': len(sys_names), 'Missing from EPT': len(sys_names),
                         'EPT-only (not in system)': 0, 'Status': '🔴 EPT sheet not found'})
            continue

        ept_names = set(ept_df['Cell Name'].dropna().unique())
        dupes = int(ept_df['Cell Name'].duplicated().sum())
        missing = len(sys_names - ept_names)
        ept_only = len(ept_names - sys_names)
        status = '🟢 Clean' if dupes == 0 and missing == 0 and ept_only == 0 else '🟡 Review needed'
        rows.append({
            'Technology': tech,
            'EPT Rows': len(ept_df),
            'EPT Duplicate Cell Names': dupes,
            'System Cells': len(sys_names),
            'Missing from EPT': missing,
            'EPT-only (not in system)': ept_only,
            'Status': status,
        })
    return pd.DataFrame(rows)


def get_ept_duplicates(tech: str, ept_dir: str = EPT_DIR) -> pd.DataFrame:
    """Full rows for every Cell Name that appears more than once in the EPT
    sheet - these are real data-entry conflicts that need an RF engineer's
    judgment on which value is correct."""
    df = load_ept(tech, ept_dir)
    if df is None or 'Cell Name' not in df.columns:
        return pd.DataFrame()
    dupes = df[df['Cell Name'].duplicated(keep=False)].sort_values('Cell Name')
    return dupes


def get_ept_only_rows(tech: str, csv_folder: str = "output/csv", ept_dir: str = EPT_DIR) -> pd.DataFrame:
    """Full EPT rows for cells that don't appear in the system's authoritative
    KPI sheet - candidates for retiring from EPT (decommissioned/renamed) or
    documenting why they're legitimately not reporting yet."""
    df = load_ept(tech, ept_dir)
    if df is None or 'Cell Name' not in df.columns:
        return pd.DataFrame()
    sys_path = os.path.join(csv_folder, f"{SYSTEM_CELL_SHEETS[tech]}.csv")
    sys_names = set()
    if os.path.exists(sys_path):
        try:
            sys_names = set(pd.read_csv(sys_path)['Cell Name'].dropna().unique())
        except Exception:
            pass
    return df[~df['Cell Name'].isin(sys_names)].sort_values('Cell Name')


def get_missing_from_ept(tech: str, csv_folder: str = "output/csv", ept_dir: str = EPT_DIR) -> pd.DataFrame:
    """Cell Names reporting KPIs in the system but absent from EPT entirely -
    these need a brand-new row added, not just a correction."""
    sys_path = os.path.join(csv_folder, f"{SYSTEM_CELL_SHEETS[tech]}.csv")
    sys_names = set()
    if os.path.exists(sys_path):
        try:
            sys_names = set(pd.read_csv(sys_path)['Cell Name'].dropna().unique())
        except Exception:
            pass
    df = load_ept(tech, ept_dir)
    ept_names = set(df['Cell Name'].dropna().unique()) if df is not None and 'Cell Name' in df.columns else set()
    missing = sorted(sys_names - ept_names)
    return pd.DataFrame({'Cell Name': missing})


def generate_review_list_excel(csv_folder: str = "output/csv", ept_dir: str = EPT_DIR) -> bytes:
    """One workbook, three sheets per technology (Duplicates / EPT-only /
    Missing-from-EPT) - a working list an RF engineer can review row-by-row
    while editing the EPT file directly."""
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        summary = validate_ept(csv_folder, ept_dir)
        summary.to_excel(writer, sheet_name='Summary', index=False)
        for tech in EPT_SHEETS:
            dupes = get_ept_duplicates(tech, ept_dir)
            if not dupes.empty:
                dupes.to_excel(writer, sheet_name=f'{tech}_Duplicates', index=False)
            ept_only = get_ept_only_rows(tech, csv_folder, ept_dir)
            if not ept_only.empty:
                ept_only.to_excel(writer, sheet_name=f'{tech}_EPT_only', index=False)
            missing = get_missing_from_ept(tech, csv_folder, ept_dir)
            if not missing.empty:
                missing.to_excel(writer, sheet_name=f'{tech}_Missing', index=False)
    buf.seek(0)
    return buf.read()


# ------------------------------------------------------------------
# KML export
# ------------------------------------------------------------------

def _kml_placemark_point(name: str, lon: float, lat: float, description: str, color: str) -> str:
    return (
        f"<Placemark><name>{escape(str(name))}</name>"
        f"<Style><IconStyle><color>{color}</color><scale>0.8</scale>"
        f"<Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon>"
        f"</IconStyle></Style>"
        f"<description><![CDATA[{description}]]></description>"
        f"<Point><coordinates>{lon},{lat},0</coordinates></Point></Placemark>"
    )


def _kml_azimuth_line(name: str, lon: float, lat: float, azimuth, color: str, length_m: float = 150) -> str:
    """Short line from the cell's location pointing in its azimuth direction
    - a lightweight way to visualize sector orientation without computing a
    full coverage-sector wedge polygon."""
    try:
        az_rad = math.radians(float(azimuth))
        lat_f, lon_f = float(lat), float(lon)
    except (TypeError, ValueError):
        return ""
    dlat = (length_m * math.cos(az_rad)) / 111320.0
    dlon = (length_m * math.sin(az_rad)) / (111320.0 * math.cos(math.radians(lat_f)) or 1e-9)
    lat2, lon2 = lat_f + dlat, lon_f + dlon
    return (
        f"<Placemark><name>{escape(str(name))} (az {azimuth}°)</name>"
        f"<Style><LineStyle><color>{color}</color><width>2</width></LineStyle></Style>"
        f"<LineString><coordinates>{lon_f},{lat_f},0 {lon2},{lat2},0</coordinates></LineString></Placemark>"
    )


EPT_KML_DESC_FIELDS = ['Site Name', 'Sector Name', 'Azimuth', 'Mechanical Tilt',
                        'Electric Tilt', 'Antenna Height (m)', 'Type', 'Active Status']


def generate_ept_kml(techs: Optional[List[str]] = None, ept_dir: str = EPT_DIR) -> bytes:
    """One folder per technology; each cell gets a point placemark (its
    location) plus a short line showing its azimuth pointing direction, for
    opening in Google Earth / GIS tools."""
    techs = techs or EPT_SHEETS
    generated = datetime.now().strftime('%Y-%m-%d %H:%M')
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2"><Document>',
        f'<name>Libyana EPT - Whole Network (generated {generated})</name>',
    ]

    for tech in techs:
        df = load_ept(tech, ept_dir)
        if df is None or df.empty:
            continue
        color = TECH_KML_COLOR.get(tech, 'ffffffff')
        parts.append(f'<Folder><name>{escape(tech)} ({len(df)} cells)</name>')
        desc_fields = [f for f in EPT_KML_DESC_FIELDS if f in df.columns]
        for _, row in df.iterrows():
            lon, lat = row.get('Longitude'), row.get('Latitude')
            if pd.isna(lon) or pd.isna(lat):
                continue
            cell_name = row.get('Cell Name', '')
            desc = '<br/>'.join(f"{f}: {row.get(f, '')}" for f in desc_fields)
            parts.append(_kml_placemark_point(cell_name, lon, lat, desc, color))
            az = row.get('Azimuth')
            if pd.notna(az):
                parts.append(_kml_azimuth_line(cell_name, lon, lat, az, color))
        parts.append('</Folder>')

    parts.append('</Document></kml>')
    return '\n'.join(parts).encode('utf-8')
