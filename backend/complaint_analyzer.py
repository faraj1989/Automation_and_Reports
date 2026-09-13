#!/usr/bin/env python3
"""
Libyana NPM - Coverage/User Complaint Analyzer

Turns a complainant's coordinates into an RF engineering answer: which
cell(s) most plausibly serve that location, and - reusing the same
KPI-threshold/alarm machinery the rest of this dashboard already trusts -
whether those cells are actually healthy right now.

Two modes:
  - Specific user: one point. Ranked candidate serving cells per
    technology (2G/3G/4G), using distance + azimuth-wedge geometry against
    config/Libyana MS EPT_*-Whole Network.xlsx (backend/ept_manager.py).
  - Whole area: a center point + radius. Every active cell whose site
    falls inside the circle, regardless of azimuth (the complaint is about
    the area, not one directional link).

EPT gives real per-cell coordinates/azimuth (RF-engineer-maintained, so
occasionally imperfect - see config/Note for EPT Engineering parameters
file.txt) and, for LTE only, a real per-cell "Cell Radius (m)". GSM/UMTS
have no such column, so config/cell_radius_defaults.csv supplies an
industry-rule-of-thumb default by band + urban-area code - a judgment
call, not a measurement, and editable without touching code.

Serving-cell determination from coordinates alone is inherently
approximate (no terrain, indoor/outdoor, or live signal-strength data) -
this returns ranked candidates for an RF engineer to confirm, not a single
guaranteed answer.
"""
import math
import os
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import pandas as pd

from backend import ept_manager as ept

RADIUS_DEFAULTS_FILE = "config/cell_radius_defaults.csv"

# EPT's own per-technology "is this cell in service" values aren't spelled
# the same way twice (GSM/UMTS use ACTIVATED/DEACTIVATED, LTE uses
# CELL_ACTIVE/CELL_DEACTIVE) - matched case-insensitively by substring so
# one predicate covers all three sheets.
_ACTIVE_TOKEN = "activ"
_INACTIVE_TOKEN = "deactiv"

# Default sector antenna half-beamwidth for the coverage-wedge geometry and
# for drawing the wedge polygons on the map - real sector antennas are
# commonly ~65 degrees nominal, but coverage doesn't stop at a hard edge,
# so this is deliberately wider than half of 65 degrees to avoid excluding
# the true serving cell just because a complaint point sits near a sector
# boundary. Judgment call, exposed as a parameter so it can be tuned per
# investigation rather than baked in silently.
DEFAULT_WEDGE_HALF_DEG = 55

EARTH_RADIUS_KM = 6371.0088


# ------------------------------------------------------------------
# Geometry
# ------------------------------------------------------------------

def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlambda / 2) ** 2
    return 2 * EARTH_RADIUS_KM * math.asin(min(1.0, math.sqrt(a)))


def bearing_deg(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Compass bearing (0-360, 0=North) from point 1 to point 2."""
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dlambda = math.radians(lon2 - lon1)
    x = math.sin(dlambda) * math.cos(p2)
    y = math.cos(p1) * math.sin(p2) - math.sin(p1) * math.cos(p2) * math.cos(dlambda)
    return (math.degrees(math.atan2(x, y)) + 360) % 360


def angular_diff_deg(a: float, b: float) -> float:
    """Smallest difference between two compass angles, 0-180."""
    d = abs(a - b) % 360
    return min(d, 360 - d)


def destination_point(lat: float, lon: float, bearing: float, distance_km: float):
    """Point `distance_km` from (lat, lon) along `bearing` - used to build
    the wedge/circle polygons drawn on the map."""
    p1 = math.radians(lat)
    br = math.radians(bearing)
    ang = distance_km / EARTH_RADIUS_KM
    p2 = math.asin(math.sin(p1) * math.cos(ang) + math.cos(p1) * math.sin(ang) * math.cos(br))
    l2 = math.radians(lon) + math.atan2(
        math.sin(br) * math.sin(ang) * math.cos(p1), math.cos(ang) - math.sin(p1) * math.sin(p2)
    )
    return math.degrees(p2), (math.degrees(l2) + 540) % 360 - 180


def circle_polygon(lat: float, lon: float, radius_km: float, n_points: int = 72) -> List[tuple]:
    return [destination_point(lat, lon, 360 * i / n_points, radius_km) for i in range(n_points + 1)]


def sector_wedge_polygon(lat: float, lon: float, azimuth: float, radius_km: float,
                          half_beamwidth_deg: float = 32.5, n_arc_points: int = 10) -> List[tuple]:
    """Site -> arc from (azimuth - half_beamwidth) to (azimuth + half_beamwidth)
    at radius_km -> back to site: the pie-slice sector shape seen in the RF
    team's own coverage-cone map views (config/*.txt reference screenshot)."""
    start = azimuth - half_beamwidth_deg
    points = [(lat, lon)]
    for i in range(n_arc_points + 1):
        brg = start + (2 * half_beamwidth_deg) * i / n_arc_points
        points.append(destination_point(lat, lon, brg, radius_km))
    points.append((lat, lon))
    return points


# ------------------------------------------------------------------
# EPT loading
# ------------------------------------------------------------------

_TECH_TYPE_COL = {'GSM': 'Type', 'UMTS': 'Type'}
_TECH_BAND_LABEL = {
    'GSM': lambda row: row.get('Type'),  # already 'GSM900'/'DCS1800'
    'UMTS': lambda row: {1: 'UMTS2100', 8: 'UMTS900'}.get(row.get('Frequency Band'), None),
}


def _load_radius_defaults() -> pd.DataFrame:
    if not os.path.exists(RADIUS_DEFAULTS_FILE):
        return pd.DataFrame()
    return pd.read_csv(RADIUS_DEFAULTS_FILE)


def _is_active(value) -> bool:
    v = str(value).lower()
    return _ACTIVE_TOKEN in v and _INACTIVE_TOKEN not in v


def load_all_ept_cells() -> pd.DataFrame:
    """GSM + UMTS + LTE EPT sheets combined into one DataFrame with a common
    shape: Technology, Cell Name, Site Name, Sector Name, Latitude,
    Longitude, Azimuth, Radius (km), Band, Urban Area, Active. LTE cells
    use their own real 'Cell Radius (m)'; GSM/UMTS fall back to
    config/cell_radius_defaults.csv by band + urban-area code. Rows with no
    usable coordinates or azimuth are dropped - they can't be geometrically
    matched regardless of how healthy the cell's KPIs are."""
    defaults = _load_radius_defaults()

    def default_radius(tech, band, urban):
        if defaults.empty or band is None:
            return None
        match = defaults[(defaults['Technology'] == tech) & (defaults['Band'] == band)
                          & (defaults['Urban Area'] == urban)]
        if match.empty:
            match = defaults[(defaults['Technology'] == tech) & (defaults['Band'] == band)]
        return float(match.iloc[0]['Radius (km)']) if not match.empty else None

    frames = []
    for tech in ['GSM', 'UMTS', 'LTE']:
        df = ept.load_ept(tech)
        if df is None or df.empty:
            continue
        df = df.copy()
        df['Technology'] = tech
        df['Active'] = df.get('Active Status', '').apply(_is_active)

        if tech == 'LTE':
            df['Radius (km)'] = pd.to_numeric(df.get('Cell Radius (m)'), errors='coerce') / 1000.0
            df['Band'] = df.get('Frequency Band').apply(lambda b: f"LTE Band {b}" if pd.notna(b) else None)
        else:
            band_fn = _TECH_BAND_LABEL[tech]
            df['Band'] = df.apply(band_fn, axis=1)
            df['Radius (km)'] = df.apply(
                lambda r: default_radius(tech, r['Band'], r.get('Urban area')), axis=1)

        keep_cols = {
            'Technology': 'Technology', 'Cell Name': 'Cell Name', 'Site Name': 'Site Name',
            'Sector Name': 'Sector Name', 'Latitude': 'Latitude', 'Longitude': 'Longitude',
            'Azimuth': 'Azimuth', 'Radius (km)': 'Radius (km)', 'Band': 'Band',
            'Urban area': 'Urban Area', 'Active': 'Active',
        }
        available = {k: v for k, v in keep_cols.items() if k in df.columns}
        frames.append(df[list(available.keys())].rename(columns=available))

    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    for col in ('Latitude', 'Longitude', 'Azimuth', 'Radius (km)'):
        combined[col] = pd.to_numeric(combined[col], errors='coerce')
    return combined.dropna(subset=['Latitude', 'Longitude', 'Azimuth'])


# ------------------------------------------------------------------
# Matching
# ------------------------------------------------------------------

def find_serving_cells(lat: float, lon: float, ept_cells: Optional[pd.DataFrame] = None,
                        top_n: int = 3, wedge_half_deg: float = DEFAULT_WEDGE_HALF_DEG,
                        active_only: bool = True) -> pd.DataFrame:
    """Ranked candidate serving cell(s) per technology for one point.

    A cell is a candidate if the complaint point falls within its radius
    AND within wedge_half_deg of its azimuth boresight - both distance and
    angular deviation feed the ranking (closer + more on-boresight ranks
    higher), so the top row per technology is the single best guess and
    the rest are the alternates an RF engineer should also glance at."""
    cells = ept_cells if ept_cells is not None else load_all_ept_cells()
    if cells.empty:
        return pd.DataFrame()
    if active_only:
        cells = cells[cells['Active']]

    df = cells.copy()
    df['Distance (km)'] = df.apply(lambda r: haversine_km(lat, lon, r['Latitude'], r['Longitude']), axis=1)
    df['Bearing'] = df.apply(lambda r: bearing_deg(r['Latitude'], r['Longitude'], lat, lon), axis=1)
    df['Azimuth Diff'] = df.apply(lambda r: angular_diff_deg(r['Bearing'], r['Azimuth']), axis=1)
    df['Effective Radius (km)'] = df['Radius (km)'].fillna(5.0)  # conservative fallback if no default matched

    candidates = df[(df['Distance (km)'] <= df['Effective Radius (km)']) & (df['Azimuth Diff'] <= wedge_half_deg)].copy()
    if candidates.empty:
        return candidates

    # Lower is better: normalize distance-as-fraction-of-radius and angular
    # deviation-as-fraction-of-wedge onto the same 0-1-ish scale before
    # combining, so a cell with a huge radius doesn't win purely for being
    # "close in absolute km" when it's really at the edge of its own cell.
    candidates['Score'] = (
        candidates['Distance (km)'] / candidates['Effective Radius (km)']
        + candidates['Azimuth Diff'] / wedge_half_deg
    )
    candidates = candidates.sort_values(['Technology', 'Score'])
    candidates['Rank'] = candidates.groupby('Technology').cumcount() + 1
    return candidates[candidates['Rank'] <= top_n].reset_index(drop=True)


def find_cells_in_area(lat: float, lon: float, radius_km: float,
                        ept_cells: Optional[pd.DataFrame] = None,
                        active_only: bool = True) -> pd.DataFrame:
    """Every cell (any azimuth) whose site falls within radius_km of
    (lat, lon) - for "the whole area is bad," not one directional link."""
    cells = ept_cells if ept_cells is not None else load_all_ept_cells()
    if cells.empty:
        return pd.DataFrame()
    if active_only:
        cells = cells[cells['Active']]
    df = cells.copy()
    df['Distance (km)'] = df.apply(lambda r: haversine_km(lat, lon, r['Latitude'], r['Longitude']), axis=1)
    return df[df['Distance (km)'] <= radius_km].sort_values(['Technology', 'Distance (km)']).reset_index(drop=True)


TECH_COLOR = {'GSM': '#C0392B', 'UMTS': '#B8860B', 'LTE': '#1E7B4D'}
HIGHLIGHT_COLOR = '#1F6FEB'


def build_complaint_map(cells_to_plot: pd.DataFrame, center: Optional[tuple] = None,
                         radius_km: Optional[float] = None,
                         highlight_index: Optional[list] = None):
    """One Scattermapbox figure: a coverage-wedge polygon per cell (colored
    by technology, matching the RF team's own coverage-cone map views),
    the complaint point, and the search circle if this is an area lookup.
    Ranked serving-cell candidates (highlight_index, into cells_to_plot's
    own index) are redrawn with a bold outline so the best-guess answer is
    visually obvious against the rest of the neighbourhood."""
    import plotly.graph_objects as go

    fig = go.Figure()
    highlight_index = set(highlight_index or [])
    seen_legend = set()

    for idx, row in cells_to_plot.iterrows():
        tech = row['Technology']
        is_hl = idx in highlight_index
        # NaN (a missing radius that fell through every default) is
        # truthy in Python, so a plain "or" would pass it straight through
        # instead of falling back - confirmed live: that produced NaN
        # polygon points, which Plotly's SVG renderer can't draw at all
        # ("Expected number" path error, silently blanking the whole map).
        radius = row.get('Radius (km)')
        wedge = sector_wedge_polygon(row['Latitude'], row['Longitude'], row['Azimuth'],
                                      radius if pd.notna(radius) else 3.0)
        lats, lons = zip(*wedge)
        color = HIGHLIGHT_COLOR if is_hl else TECH_COLOR.get(tech, '#888888')
        show_legend = tech not in seen_legend and not is_hl
        seen_legend.add(tech)
        fig.add_trace(go.Scattermapbox(
            lat=list(lats), lon=list(lons), mode='lines', fill='toself',
            fillcolor=color, opacity=0.85 if is_hl else 0.35,
            line=dict(color=color, width=3 if is_hl else 1),
            name=f"★ {row['Cell Name']}" if is_hl else tech,
            legendgroup=tech, showlegend=show_legend,
            text=f"{row['Cell Name']} ({tech}) · az {row['Azimuth']:.0f}° · {row.get('Radius (km)', 0):.1f} km radius",
            hoverinfo='text',
        ))

    if center:
        fig.add_trace(go.Scattermapbox(
            lat=[center[0]], lon=[center[1]], mode='markers',
            marker=dict(size=16, color='black', symbol='star'),
            name='Complaint location', hoverinfo='name',
        ))
        if radius_km:
            circle = circle_polygon(center[0], center[1], radius_km)
            clats, clons = zip(*circle)
            fig.add_trace(go.Scattermapbox(
                lat=list(clats), lon=list(clons), mode='lines',
                line=dict(color='black', width=2),
                name=f'{radius_km:.2f} km radius', hoverinfo='name',
            ))

    center_lat = center[0] if center else (cells_to_plot['Latitude'].mean() if not cells_to_plot.empty else 32.0)
    center_lon = center[1] if center else (cells_to_plot['Longitude'].mean() if not cells_to_plot.empty else 20.0)

    # A GSM/UMTS macro wedge can be 10-15km deep - a fixed zoom (confirmed
    # live: zoom=14, city-block scale) left the wedges many times larger
    # than the whole viewport, filling it edge-to-edge with solid color and
    # hiding the base map entirely. Pick zoom from the actual extent being
    # drawn instead, so the view always frames what's plotted.
    radii = cells_to_plot['Radius (km)'].dropna().tolist() if not cells_to_plot.empty else []
    if radius_km:
        radii.append(radius_km)
    max_extent_km = max(radii) if radii else 2.0
    zoom = 15
    for span, z in ((0.6, 15), (1.2, 14), (2.5, 13), (5, 12), (10, 11), (20, 10), (40, 9)):
        if max_extent_km <= span:
            zoom = z
            break
    else:
        zoom = 8

    fig.update_layout(
        mapbox=dict(style='open-street-map', center=dict(lat=center_lat, lon=center_lon), zoom=zoom),
        margin=dict(l=0, r=0, t=0, b=0), height=520,
        legend=dict(orientation='h', y=-0.02),
    )
    return fig


# ------------------------------------------------------------------
# KPI / alarm / interference context (reuses existing, already-trusted logic)
# ------------------------------------------------------------------

_INTERFERENCE_FILES = {
    'GSM': ('output/csv/2G_Interference.csv', 'Interference Band Proportion (4~5)(%)', 'Cell Name'),
}


def get_alarm_history_for_sites(site_names: List[str], center_date: str, lookback_days: int = 7) -> pd.DataFrame:
    """Every NE Is Disconnected event recorded for the given sites in the
    lookback_days ending at center_date, reusing build_daily_noc_alarm_report
    (same hub-aware Power/Transmission logic as the Alarms tab) one day at
    a time and filtering down to just these sites."""
    from backend import noc_alarm_processor as noc

    if not site_names:
        return pd.DataFrame()
    center = datetime.strptime(center_date, '%Y-%m-%d').date()
    rows = []
    for offset in range(lookback_days):
        day = (center - timedelta(days=offset)).isoformat()
        report = noc.build_daily_noc_alarm_report(day)
        summary = report.get('down_sites_summary')
        if summary is None or summary.empty:
            continue
        matched = summary[summary['Site Name'].isin(site_names)]
        if not matched.empty:
            rows.append(matched)
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()


def get_interference_for_cells(cell_names: List[str], tech: str, center_date: str,
                                lookback_days: int = 7) -> pd.DataFrame:
    """2G per-cell interference-band-proportion rows for the window - the
    3G/4G rollups are per (Cell Name, Band, Date) BadHours, a different
    shape best left to the dedicated Interference view rather than
    force-fit here; this covers the one case with a direct daily %."""
    if tech != 'GSM' or not cell_names:
        return pd.DataFrame()
    path, metric_col, cell_col = _INTERFERENCE_FILES['GSM']
    if not os.path.exists(path):
        return pd.DataFrame()
    df = pd.read_csv(path)
    if cell_col not in df.columns or 'Date' not in df.columns:
        return pd.DataFrame()
    center = datetime.strptime(center_date, '%Y-%m-%d').date()
    start = (center - timedelta(days=lookback_days - 1)).isoformat()
    df['_d'] = pd.to_datetime(df['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    sub = df[(df[cell_col].isin(cell_names)) & (df['_d'] >= start) & (df['_d'] <= center_date)]
    return sub.drop(columns=['_d']) if not sub.empty else pd.DataFrame()


def build_complaint_analysis(report_gen, matched_cells: pd.DataFrame, complaint_date: str,
                              lookback_days: int = 7) -> Dict:
    """For every matched cell (from find_serving_cells or find_cells_in_area),
    pull today's failing KPIs (with suggested fix), a recent trend, and its
    parent site's recent alarm history - then write one narrative line per
    cell so the final report reads as an answer, not just a pile of tables.

    Reuses ReportGenerator.get_cell_failing_kpis/build_cell_trend (the
    exact same threshold logic and Suggested Fix rules as the Worst Cells
    section) and noc_alarm_processor's hub-aware alarm correlation - a
    complaint investigation should never disagree with what the rest of
    this dashboard already says about the same cell."""
    from backend.report_generator import CELL_SHEETS, SITE_COL_BY_TECH

    if matched_cells.empty:
        return {'cells': pd.DataFrame(), 'failing_kpis': pd.DataFrame(), 'alarms': pd.DataFrame(),
                'interference': pd.DataFrame(), 'narratives': []}

    all_failing = []
    all_interference = []
    narratives = []
    site_names_by_tech = {}

    for tech, group in matched_cells.groupby('Technology'):
        cell_names = group['Cell Name'].dropna().unique().tolist()
        if not cell_names or tech not in CELL_SHEETS:
            continue
        site_names_by_tech[tech] = group['Site Name'].dropna().unique().tolist()

        failing = report_gen.get_cell_failing_kpis(tech, cell_names, complaint_date)
        if not failing.empty:
            failing.insert(0, 'Technology', tech)
            all_failing.append(failing)

        interference = get_interference_for_cells(cell_names, tech, complaint_date, lookback_days)
        if not interference.empty:
            all_interference.append(interference)

        for cell_name in cell_names:
            cell_failing = failing[failing['Cell'] == cell_name] if not failing.empty else pd.DataFrame()
            if cell_failing.empty:
                narratives.append(f"{tech} · {cell_name}: no KPI thresholds breached on {complaint_date}.")
            else:
                kpi_list = ", ".join(cell_failing['Failing KPI'].tolist())
                fixes = " | ".join(sorted(set(cell_failing['Suggested Fix'].tolist())))
                narratives.append(
                    f"{tech} · {cell_name}: failing {kpi_list} on {complaint_date}. Suggested: {fixes}."
                )

    all_site_names = sorted({s for names in site_names_by_tech.values() for s in names})
    alarms = get_alarm_history_for_sites(all_site_names, complaint_date, lookback_days)
    if not alarms.empty:
        for site in alarms['Site Name'].unique():
            site_alarms = alarms[alarms['Site Name'] == site]
            hours = site_alarms['Down Hours'].sum() if 'Down Hours' in site_alarms.columns else 0
            reasons = sorted(set(site_alarms['Root Cause Bucket'].dropna())) if 'Root Cause Bucket' in site_alarms.columns else []
            narratives.append(
                f"Site {site}: {len(site_alarms)} NE Is Disconnected event(s) in the last {lookback_days} day(s), "
                f"{hours:.1f} total hours down. Root cause: {', '.join(reasons) or 'Investigating'}."
            )

    return {
        'cells': matched_cells,
        'failing_kpis': pd.concat(all_failing, ignore_index=True) if all_failing else pd.DataFrame(),
        'alarms': alarms,
        'interference': pd.concat(all_interference, ignore_index=True) if all_interference else pd.DataFrame(),
        'narratives': narratives,
    }


# ------------------------------------------------------------------
# Word report
# ------------------------------------------------------------------

def generate_complaint_word_report(report_gen, complainant: Dict, mode: str, analysis: Dict) -> bytes:
    """One Word document: complainant details, the narrative findings, then
    supporting tables - reuses ReportGenerator's own Word-building helpers
    (_docx_section_heading/_docx_add_table) so this looks like every other
    export from this dashboard, not a one-off format."""
    import io
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)

    h = doc.add_heading('Network Complaint Investigation', level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
    gen = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Mode: {mode}")
    gen.alignment = WD_ALIGN_PARAGRAPH.CENTER
    gen.runs[0].font.size = Pt(8)
    gen.runs[0].font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    doc.add_paragraph()

    report_gen._docx_section_heading(doc, "Complaint Details")
    details = pd.DataFrame([complainant])
    report_gen._docx_add_table(doc, details)

    report_gen._docx_section_heading(doc, "Findings")
    if analysis['narratives']:
        for line in analysis['narratives']:
            doc.add_paragraph(line, style='List Bullet')
    else:
        doc.add_paragraph("No cells could be matched to the given location/radius.")

    if not analysis['cells'].empty:
        report_gen._docx_section_heading(doc, "Matched Cells")
        report_gen._docx_add_table(doc, analysis['cells'])
    if not analysis['failing_kpis'].empty:
        report_gen._docx_section_heading(doc, "Failing KPIs")
        report_gen._docx_add_table(doc, analysis['failing_kpis'])
    if not analysis['alarms'].empty:
        report_gen._docx_section_heading(doc, "Alarm History")
        report_gen._docx_add_table(doc, analysis['alarms'])
    if not analysis['interference'].empty:
        report_gen._docx_section_heading(doc, "Interference (2G)")
        report_gen._docx_add_table(doc, analysis['interference'])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ------------------------------------------------------------------
# Excel report
# ------------------------------------------------------------------

def generate_complaint_excel_report(report_gen, complainant: Dict, mode: str, analysis: Dict) -> bytes:
    """Same content as generate_complaint_word_report, as .xlsx (one sheet
    per table) - the Matched Cells / Alarm History / Interference tables can
    run to hundreds of rows, which reads far better in Excel than in Word.
    Reuses ReportGenerator.generate_tables_excel_report, the same generic
    per-tab Excel export every other dashboard tab uses."""
    findings = (pd.DataFrame({'Finding': analysis['narratives']}) if analysis['narratives']
                else pd.DataFrame({'Finding': ["No cells could be matched to the given location/radius."]}))
    tables = [
        ("Complaint Details", pd.DataFrame([complainant])),
        ("Findings", findings),
        ("Matched Cells", analysis['cells']),
        ("Failing KPIs", analysis['failing_kpis']),
        ("Alarm History", analysis['alarms']),
        ("Interference (2G)", analysis['interference']),
    ]
    return report_gen.generate_tables_excel_report(
        "Network Complaint Investigation", tables, subtitle=f"Mode: {mode}")
