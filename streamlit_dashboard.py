#!/usr/bin/env python3
"""
Libyana NPM - Network Performance Dashboard (Streamlit)

Reads the exact same config-driven ReportGenerator used to build the daily
email/Excel/Word report (backend/report_generator.py), so the RF and NOC
teams see identical numbers here and in the emailed report, and can browse
and filter the underlying data themselves instead of waiting for a static
file. Nothing about a KPI's name/threshold/weight is hardcoded here - it all
comes from config/kpi_thresholds.csv via HealthChecker, same as the report.
"""

import glob
import os
import sys
from typing import Dict, List

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from backend.report_generator import (
    ReportGenerator, TECH_LABELS, CELL_SHEETS, SCORECARD_SHEETS, SITE_COL_BY_TECH, autofit_excel_columns, base_site_name,
)
from backend import ept_manager as ept
from backend.special_reports_processor import (
    build_nq_template_report, load_hq_traffic_history, compute_hq_traffic_month,
    save_hq_traffic_month, HQ_TRAFFIC_TEMPLATE_FILE, HQ_TRAFFIC_METRICS,
)
from backend.topology_processor import build_site_topology_csv, find_topology_xlsx
from backend import smartcare_cem_processor as smartcare_cem
from backend import device_penetration_processor as device_penetration
from backend import complaint_analyzer
from backend import noc_alarm_processor as noc_alarms
from project_config import env_path_str

# reports/cell_info_report.py (monthly FTPS cell inventory pull) and
# reports/PS Traffic per site v3.py (per-site PS traffic puller), ported
# in from the sibling NOC Automation Suite project (2026-09-14) - the
# dashboard reads their finished output files, same DATA_ROOT convention
# as every scraper in this project.
CELL_INFO_OUTPUT_DIR = env_path_str(
    "CELL_INFO_OUTPUT_DIR",
    os.path.join(env_path_str("DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data"), "Output", "Cell_Info"),
)


def _load_ps_traffic_v3_module():
    """PS Traffic per site v3.py's filename has spaces, so it can't be
    `import`ed normally - load it by path instead, so its combine/summary
    logic (combine_traffic_data, generate_summary_reports) can be reused
    against a differently-sourced input without duplicating that logic."""
    import importlib.util
    path = os.path.join(os.path.dirname(__file__), "reports", "PS Traffic per site v3.py")
    spec = importlib.util.spec_from_file_location("ps_traffic_per_site_v3", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


ps_traffic_v3 = _load_ps_traffic_v3_module()
PS_TRAFFIC_SCRIPT_PATH = os.path.join(os.path.dirname(__file__), "reports", "PS Traffic per site v3.py")


def launch_ps_traffic_refresh():
    """Runs reports/"PS Traffic per site v3.py" detached in the background,
    same on-demand trigger as scheduler.py's own --ps-traffic-weekly (which
    also runs it every Sunday) - see scheduler.py's _run_report_script for
    why it's a subprocess rather than an in-process call, and why the child
    needs PYTHONIOENCODING/PYTHONUTF8 forced (its emoji prints otherwise
    crash under Windows' default cp1252 console codepage)."""
    import subprocess
    child_env = os.environ.copy()
    child_env["PYTHONUTF8"] = "1"
    child_env["PYTHONIOENCODING"] = "utf-8"
    subprocess.Popen(
        [sys.executable, PS_TRAFFIC_SCRIPT_PATH],
        cwd=os.path.dirname(PS_TRAFFIC_SCRIPT_PATH),
        env=child_env,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        start_new_session=True,
    )

st.set_page_config(page_title="Libyana Network Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .header-container {
        background: linear-gradient(90deg, #1f4e78, #17becf);
        padding: 18px 24px;
        border-radius: 10px;
        color: white;
        margin-bottom: 18px;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# DATA ACCESS (cached — same ReportGenerator the report uses)
# ============================================================
# Keyed on the class object itself: when Streamlit hot-reloads an edited
# backend module, ReportGenerator becomes a new class, so this builds a
# fresh instance instead of handing back the old one - which lacked any
# newly added method (AttributeError on build_daily_noc_alarm_summary,
# 2026-10-08) until the whole dashboard was restarted.
@st.cache_resource(max_entries=1)
def get_rg(cls_id):
    return ReportGenerator()


rg = get_rg(id(ReportGenerator))


@st.cache_data(ttl=600)
def cached_dates():
    dates = set()
    for sheet in SCORECARD_SHEETS.values():
        dates.update(rg.get_available_dates(sheet))
    return sorted(dates, reverse=True)


@st.cache_data(ttl=600)
def cached_bundle(target_date, previous_date):
    scorecards = rg.build_all_scorecards(target_date, previous_date)
    health = rg.compute_health_from_scorecards(scorecards)
    return dict(
        scorecards=scorecards,
        health=health,
        worst_cells=rg.build_worst_cells(target_date),
        site_health=rg.build_site_health(target_date),
        topology=rg.build_topology_summary(),
        traffic=rg.build_traffic_section(target_date, previous_date),
        site_inventory=rg.build_site_inventory(target_date),
        site_cards=rg.build_site_summary_cards(target_date),
        freshness=rg.build_data_freshness(target_date),
        trend=rg.build_trend(target_date, days=14),
        alarm_report=rg.build_daily_noc_alarm_report(target_date),
    )


@st.cache_data(ttl=600)
def cached_trend_thresholds(tech):
    return rg.get_trend_kpi_thresholds(tech)


@st.cache_data(ttl=600)
def cached_sheet(name):
    return rg.get_sheet(name)


@st.cache_data(ttl=600)
def cached_cell_trend(tech, cell_names, target_date):
    return rg.build_cell_trend(tech, list(cell_names), target_date, days=14)


@st.cache_data(ttl=600)
def cached_cell_trend_range(tech, cell_names, start_date, end_date):
    return rg.build_cell_trend(tech, list(cell_names), start_date=start_date, end_date=end_date)


@st.cache_data(ttl=600)
def cached_resolve_group_to_cells(tech, group_names, group_col):
    return rg.resolve_group_to_cells(tech, list(group_names), group_col=group_col)


@st.cache_data(ttl=600)
def cached_network_summary(target_date):
    return rg.build_network_summary_block(target_date)


@st.cache_data(ttl=600)
def cached_worst_cells(target_date, n_top):
    return rg.build_worst_cells(target_date, n_top=n_top)


@st.cache_data(ttl=600)
def cached_cell_failing(tech, cell_names, target_date):
    return rg.get_cell_failing_kpis(tech, list(cell_names), target_date)


@st.cache_data(ttl=600)
def cached_cell_thresholds(tech, sheet):
    return rg.get_trend_kpi_thresholds(tech, sheet=sheet)


@st.cache_data(ttl=600)
def cached_cell_dimensions(tech, sheet):
    return rg.get_trend_kpi_dimensions(tech, sheet=sheet)


@st.cache_data(ttl=600)
def cached_site_capacity_advice(site_names, target_date):
    return rg.build_site_capacity_advice(list(site_names), target_date)


@st.cache_data(ttl=600)
def cached_site_topology(site_names):
    return rg.get_site_topology(list(site_names))


@st.cache_data(ttl=600)
def cached_site_detail_summary():
    return rg.build_site_detail_summary()


# Monthly CEM report for Tripoli HQ - keyed on the history workbook's mtime so
# a fresh SmartCare run is picked up immediately, not after the TTL.
@st.cache_data(ttl=3600)
def cached_cem_months(mtime):
    return smartcare_cem.available_months()


@st.cache_data(ttl=3600)
def cached_cem_monthly(month, mtime):
    return smartcare_cem.build_monthly_cem_workbook(month)


# ---- Packet Loss tab (logic: backend/packet_loss_engine.py) ----
# Site lists are passed as tuples so st.cache_data can hash them.

@st.cache_data(ttl=600)
def cached_pl_dates():
    return rg.packet_loss_dates()


@st.cache_data(ttl=600)
def cached_pl_site_list():
    return rg.packet_loss_sites()


@st.cache_data(ttl=600)
def cached_pl_sites(start, end, sites=None):
    return rg.build_packet_loss_sites(start, end, list(sites) if sites else None)


@st.cache_data(ttl=600)
def cached_pl_hubs(start, end, sites=None):
    return rg.build_packet_loss_hubs(start, end, list(sites) if sites else None)


@st.cache_data(ttl=600)
def cached_pl_trend(start, end):
    return rg.build_packet_loss_trend(start, end)


@st.cache_data(ttl=600)
def cached_pl_matrix(start, end, sites):
    return rg.build_packet_loss_matrix(start, end, list(sites))


@st.cache_data(ttl=600)
def cached_pl_hourly(sites, start, end):
    return rg.get_packet_loss_hourly(list(sites), start, end)


@st.cache_data(ttl=600)
def cached_pl_core(start, end):
    return rg.build_packet_loss_core_links(start, end)


# Sequential blue (light = near zero, dark = worst). Loss is capped at 5% on
# the colour scale so the 0.1% / 1% steps stay visible; no-response hours are
# drawn in neutral gray on their own layer - never as "0% loss".
PL_LOSS_SCALE = [[0.0, '#f4f7fb'], [0.02, '#cde2fb'], [0.2, '#6da7ec'],
                 [0.5, '#256abf'], [1.0, '#0d366b']]
PL_NO_RESPONSE = '#8a8984'
PL_SERIES = ['#2a78d6', '#eb6834', '#1baf7a']  # categorical slots 1-3


def _pl_heatmap(z, x, y, zmax, colorbar_title, hover_value, noresp=None, height=None, x_title=None):
    fig = go.Figure(go.Heatmap(
        z=z, x=x, y=y, zmin=0, zmax=zmax, colorscale=PL_LOSS_SCALE, xgap=2, ygap=2,
        colorbar=dict(title=colorbar_title, thickness=12),
        hovertemplate=f"%{{y}} · %{{x}}<br>{hover_value}<extra></extra>",
    ))
    if noresp is not None:
        fig.add_trace(go.Heatmap(
            z=noresp, x=x, y=y, zmin=0, zmax=1, showscale=False, xgap=2, ygap=2,
            colorscale=[[0, PL_NO_RESPONSE], [1, PL_NO_RESPONSE]],
            hovertemplate="%{y} · %{x}<br>No response (no ping replies)<extra></extra>",
        ))
    fig.update_layout(
        height=height or max(220, 26 * len(y) + 90), margin=dict(l=10, r=10, t=10, b=30),
        xaxis=dict(title=x_title, side='top', showgrid=False, type='category'),
        yaxis=dict(autorange='reversed', showgrid=False, type='category'),
    )
    return fig


def render_pl_site_drilldown(site, start, end, key_prefix):
    """Day x hour heatmap for one site: which hours it lost packets, on which
    RAT, and when it stopped answering pings."""
    hourly = cached_pl_hourly((site,), start, end)
    if hourly is None or hourly.empty:
        st.info(f"No hourly detail for {site} in {start} → {end}. Hourly detail is kept for the last "
                f"{int(pl_rules()['hourly_retention_days'])} days; older periods show daily counts only.")
        return
    first_day = hourly['Time'].str[:10].min()
    if first_day > start:
        st.caption(f"Hourly detail available from {first_day} (rolling retention).")
    rat = st.radio("Link", ["Worst of 2G/3G", "2G (ABIS)", "3G (IUB)"], horizontal=True,
                   key=f"{key_prefix}_rat")
    col, nr_col = {"Worst of 2G/3G": ('Loss (%)', None),
                   "2G (ABIS)": ('Loss 2G (%)', 'No Response 2G'),
                   "3G (IUB)": ('Loss 3G (%)', 'No Response 3G')}[rat]
    h = hourly.assign(Date=hourly['Time'].str[:10], Hour=hourly['Time'].str[11:13].astype(int))
    if h[col].isna().all() and (nr_col is None or h[nr_col].isna().all()):
        st.info(f"{site} has no {rat} link.")
        return
    nr = h[col].isna() if nr_col is None else h[nr_col].eq(1)
    z = h.pivot_table(index='Date', columns='Hour', values=col, aggfunc='max').reindex(columns=range(24))
    zn = h.assign(_nr=nr.astype(float).where(nr)).pivot_table(
        index='Date', columns='Hour', values='_nr', aggfunc='max').reindex(index=z.index, columns=range(24))
    hours = [f"{x:02d}:00" for x in range(24)]
    fig = _pl_heatmap(z.values, hours, list(z.index), 5, "Loss %", "Loss %{z:.2f}%",
                      noresp=zn.values, x_title=None)
    st.plotly_chart(fig, width='stretch', key=f"{key_prefix}_heat")
    st.caption("Darker blue = more loss (scale capped at 5%). Gray = no ping response. "
               "Loss that darkens every evening and clears at night points to congestion; "
               "loss spread across all hours points to a link-quality fault.")
    with st.expander("Hourly values (table)"):
        st.dataframe(hourly, width='stretch', hide_index=True, height=300)


@st.cache_data(ttl=600)
def pl_rules():
    from backend.packet_loss_engine import load_rules
    return load_rules()


@st.cache_data(ttl=600)
def cached_topology_table():
    path = os.path.join('config', 'site_topology.csv')
    return pd.read_csv(path) if os.path.exists(path) else None


@st.cache_data(ttl=600)
def cached_traffic_regions():
    return rg.get_regions()


@st.cache_data(ttl=600)
def cached_traffic_fn_hub_nodes():
    return rg.get_fn_hub_nodes()


@st.cache_data(ttl=600)
def cached_traffic_sites_for_region(region):
    return rg.get_sites_for_region(region)


@st.cache_data(ttl=600)
def cached_traffic_sites_for_fn_hub(node_name):
    return rg.get_sites_for_fn_hub(node_name)


@st.cache_data(ttl=600)
def cached_all_traffic_sites():
    return rg.get_all_traffic_sites()


@st.cache_data(ttl=600)
def cached_site_traffic_detail(site_names, start_date, end_date):
    return rg.build_site_traffic_detail(list(site_names), start_date, end_date)


@st.cache_data(ttl=600)
def cached_site_traffic_trend(site_names, start_date, end_date):
    return rg.build_site_traffic_aggregate_trend(list(site_names), start_date, end_date)


@st.cache_data(ttl=600)
def cached_cell_info_months():
    """Month subfolders (YYYY-MM) produced by the sibling cell_info_report.py puller."""
    if not os.path.isdir(CELL_INFO_OUTPUT_DIR):
        return []
    months = [
        d for d in os.listdir(CELL_INFO_OUTPUT_DIR)
        if os.path.isdir(os.path.join(CELL_INFO_OUTPUT_DIR, d)) and len(d) == 7 and d[4] == '-'
    ]
    return sorted(months, reverse=True)


@st.cache_data(ttl=600)
def cached_cell_info_report(month):
    """The merged 2G/3G/4G cell-info workbook for one month, plus its raw bytes for download."""
    month_dir = os.path.join(CELL_INFO_OUTPUT_DIR, month)
    candidates = sorted(
        glob.glob(os.path.join(month_dir, f"merged_df_{month}.xlsx"))
        or glob.glob(os.path.join(month_dir, "merged_df_*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )
    if not candidates:
        return None, None, None
    path = candidates[0]
    with open(path, 'rb') as f:
        raw_bytes = f.read()
    return pd.read_excel(path), os.path.basename(path), raw_bytes


# Column in each per-tech pipeline CSV that holds site-level PS/data traffic,
# aligned to PS Traffic per site v3.py's own Date/Site_Name/Traffic_GB/
# Technology shape so its combine_traffic_data()/generate_summary_reports()
# can be reused unchanged. Traffic_4G.csv only carries downlink (no separate
# uplink column in this pipeline's output), unlike the original script's raw
# source which summed DL+UL - so this tab's 4G figures run slightly under
# the FTPS-sourced tab's for the same reason.
_PS_TRAFFIC_PIPELINE_COLUMNS = {
    '2G': ('output/csv/Traffic_2G.csv', '2G PS Traffic (GB)'),
    '3G': ('output/csv/Traffic_3G.csv', '3G PS Traffic (GB)'),
    '4G': ('output/csv/Traffic_4G.csv', '4G DL Traffic (GB)'),
}


@st.cache_data(ttl=600)
def cached_ps_traffic_from_pipeline():
    """Same per-site PS traffic summaries PS Traffic per site v3.py builds,
    but sourced from output/csv/Traffic_2G/3G/4G.csv - the same daily SFTP
    pipeline that already feeds every other KPI in this dashboard - instead
    of that script's own separate FTPS raw-data puller, whose source folder
    (Input/FTP_RawData) stalled 2026-08-19. Returns None if the pipeline
    CSVs aren't present."""
    data_frames = {}
    for tech, (path, col) in _PS_TRAFFIC_PIPELINE_COLUMNS.items():
        if not os.path.exists(path):
            continue
        df = pd.read_csv(path, usecols=['Date', 'Site', col])
        df = df.rename(columns={'Site': 'Site_Name', col: 'Traffic_GB'})
        df['Technology'] = tech
        data_frames[tech] = df[['Date', 'Site_Name', 'Traffic_GB', 'Technology']]

    if not data_frames:
        return None

    combined_df = ps_traffic_v3.combine_traffic_data(data_frames)
    if combined_df.empty:
        return None
    reports = ps_traffic_v3.generate_summary_reports(combined_df)

    export_df = combined_df.copy()
    export_df['Date'] = export_df['Date'].dt.strftime('%d-%m-%Y')
    export_df = export_df.drop(columns=['Date_Formatted'])
    sheets = {'All_Traffic_Data': export_df}
    for sheet_name, df in reports.items():
        if df.empty:
            continue
        df = df.copy()
        if 'Date' in df.columns and 'Date_Formatted' in df.columns:
            df['Date'] = df['Date_Formatted']
            df = df.drop(columns=['Date_Formatted'])
        sheets[sheet_name] = df
    return sheets


@st.cache_data(ttl=600)
def cached_ps_traffic_from_pipeline_bytes():
    import io
    sheets = cached_ps_traffic_from_pipeline()
    if not sheets:
        return None
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        autofit_excel_columns(writer)
    return buf.getvalue()


# Shorter TTL than the rest (600s) since the underlying live alarm feed
# refreshes every ~5 min upstream - a 10-min-stale cache would mask a
# genuinely fresh update for no benefit.
@st.cache_data(ttl=180)
def cached_alarm_overview():
    return rg.build_alarm_overview()


@st.cache_data(ttl=180)
def cached_site_alarm_status(site_names):
    return rg.get_site_alarm_status(list(site_names))


# Raw historical exports this reads are retained only a few days upstream
# and don't change once a day has fully elapsed, but today's/yesterday's
# figures can still firm up as later cycles capture more Cleared-On times
# - a short TTL, same as the rest of the Alarms section, keeps that timely
# without re-parsing the raw exports on every rerun.
@st.cache_data(ttl=180)
def cached_daily_noc_alarm_report(target_date):
    return rg.build_daily_noc_alarm_report(target_date)


# A multi-day range walks every day in it (~15s cold for all ~19 covered
# days, measured 2026-10-08) - a slightly longer TTL than the single-day
# view, still short enough that Not Cleared flips to Cleared promptly.
@st.cache_data(ttl=300)
def cached_daily_noc_alarm_range(start_date, end_date):
    return rg.build_daily_noc_alarm_range(start_date, end_date)


@st.cache_data(ttl=900)
def cached_noc_alarm_dates():
    return noc_alarms.available_daily_noc_alarm_dates()


def _style_alarm_status(df):
    """Status / Outage Class colouring for the Daily NOC Alarm tables -
    same colours as the Excel export (noc_alarm_processor._STATUS_FILLS)."""
    colors = {
        noc_alarms.STATUS_CLEARED: 'color: #2ca02c',
        noc_alarms.STATUS_NOT_CLEARED: 'color: #d62728; font-weight: bold',
        noc_alarms.STATUS_UNKNOWN: 'color: #b8860b; font-weight: bold',
        noc_alarms.CLASS_LONG_TERM: 'color: #7b4fa0; font-weight: bold',
    }
    cols = [c for c in ('Status', 'Outage Class') if c in df.columns]
    return df.style.map(lambda v: colors.get(v, ''), subset=cols) if cols else df


# SmartCare CEM refreshes weekly upstream - the long TTL just avoids
# re-reading the workbook from disk on every rerun within a session.
@st.cache_data(ttl=3600)
def cached_cem_overview():
    return rg.build_cem_overview()


@st.cache_data(ttl=3600)
def cached_device_penetration_overview():
    return rg.build_device_penetration_overview()


DP_TECH_ORDER = device_penetration.TECH_ORDER
DP_TECH_COLORS = {'5G NR': '#6f42c1', 'LTE (no 5G)': '#1f77b4', '3G (no LTE)': '#ff7f0e',
                  '2G only': '#d62728', 'Other': '#7f7f7f', 'Unidentified': '#bbbbbb'}
_dp_tech_bucket = device_penetration.tech_bucket


@st.cache_data(ttl=3600)
def cached_dp_weekly_summary(week_start, handsets_only):
    history = cached_device_penetration_overview()['history']
    return device_penetration.weekly_summary(history, week_start, handsets_only=handsets_only)


@st.cache_data(ttl=3600)
def cached_dp_weekly_report_bytes(week_start, handsets_only):
    return device_penetration.build_weekly_report_xlsx(cached_dp_weekly_summary(week_start, handsets_only))


# EPT (coordinates/azimuth) only changes when the RF team hand-edits the
# workbook - an hour's cache avoids re-parsing all 3 sheets (~15k rows) on
# every complaint lookup within a session.
@st.cache_data(ttl=3600)
def cached_ept_cells():
    return complaint_analyzer.load_all_ept_cells()


def render_site_detail_and_advice(site_names, target_date, key_prefix):
    """Shared block for Cell Explorer/Special Reports: SiteDetail config
    (bands/scenario/RAT) for the selected site(s), rule-based capacity
    recommendations (config-aware: names a free band slot to activate, or
    flags a sector/cabinet upgrade when none remain), and FN/HUB topology
    (which transmission node this site depends on, and which other sites
    share it - useful for spotting "shared infrastructure" root causes
    instead of investigating each site as an isolated RF problem)."""
    site_detail = cached_sheet('SiteDetail')
    st.subheader("🏗️ Site Detail")
    if site_detail is None or 'Site Name' not in site_detail.columns:
        st.info("SiteDetail.csv not found.")
    else:
        rows = site_detail[site_detail['Site Name'].isin(site_names)]
        if rows.empty:
            st.info("No SiteDetail entry for the selected site(s).")
        else:
            st.dataframe(rows, width='stretch', hide_index=True)
            refreshed = cached_site_detail_summary().get('file_refreshed')
            if refreshed:
                st.caption(f"SiteDetail last refreshed {refreshed.strftime('%Y-%m-%d %H:%M')} - "
                           "'Last Updated' column = that site's last real config change.")

    st.subheader("🔗 FN/HUB Topology")
    topo_rows = cached_site_topology(tuple(site_names))
    if topo_rows is None or topo_rows.empty:
        st.info("No FN/HUB topology entry for the selected site(s) - either not in the reference file, or not a fiber-node-dependent site.")
    else:
        st.dataframe(topo_rows, width='stretch', hide_index=True)
        st.caption("If several sites you're investigating together share the same FN/HUB Node, a transmission "
                   "issue at that node - not independent RF problems - is worth ruling out first.")

    st.subheader("💡 Capacity Recommendations")
    advice = cached_site_capacity_advice(tuple(site_names), target_date)
    if advice is None or advice.empty:
        st.info("No capacity data available for the selected site(s) on this date.")
    else:
        st.dataframe(advice, width='stretch', hide_index=True)

    st.subheader("🚨 Alarm Status (NOC Feed)")
    alarm_status = cached_site_alarm_status(tuple(site_names))
    current_alarm, chronic, downtime = alarm_status['current'], alarm_status['chronic'], alarm_status['downtime']
    if not current_alarm.empty:
        st.error(f"⚠️ {len(current_alarm)} of the selected site(s) currently show an active disconnect/power alarm:")
        st.dataframe(current_alarm, width='stretch', hide_index=True)
    if not chronic.empty:
        st.caption("Chronic alarm history for the selected site(s):")
        st.dataframe(chronic, width='stretch', hide_index=True)
    if not downtime.empty:
        st.caption("Historical downtime for the selected site(s):")
        st.dataframe(downtime, width='stretch', hide_index=True)
    if current_alarm.empty and chronic.empty and downtime.empty:
        st.caption("No current or historical alarm data for the selected site(s) "
                   "(or the NOC alarm feed isn't available on this machine).")


DIMENSION_ORDER = ["Accessibility", "Retainability", "Mobility", "Resource Utilization", "Quality"]


def render_grouped_cell_trend_charts(cell_trend_df, kpi_cols, kpi_thresholds, kpi_dimensions, key_prefix):
    """Render one multi-cell trend chart per KPI, grouped into collapsible
    per-Dimension sections (fixed order, 'Other' catch-all last) instead of
    one flat grid - keeps a large KPI set (golden + supplementary) scannable.
    First non-empty group starts expanded, the rest collapsed."""
    grouped: Dict[str, List[str]] = {}
    for kpi in kpi_cols:
        dim = kpi_dimensions.get(kpi, "Other")
        grouped.setdefault(dim, []).append(kpi)

    ordered_dims = [d for d in DIMENSION_ORDER if d in grouped] + \
        [d for d in grouped if d not in DIMENSION_ORDER]

    for i, dim in enumerate(ordered_dims):
        dim_kpis = grouped[dim]
        with st.expander(f"{dim} ({len(dim_kpis)})", expanded=(i == 0)):
            cols = st.columns(2)
            for j, kpi in enumerate(dim_kpis):
                with cols[j % 2]:
                    render_multi_cell_trend_chart(
                        cell_trend_df, kpi, kpi_thresholds.get(kpi), key=f"{key_prefix}_{kpi}",
                    )


@st.cache_data(ttl=600)
def cached_word_bytes(target_date, previous_date):
    b = cached_bundle(target_date, previous_date)
    path = rg.generate_word_report(
        target_date, previous_date, b['health'], b['scorecards'], b['worst_cells'],
        b['site_health'], b['topology'], b['traffic'], b['site_inventory'],
        b['freshness'], b['trend'], b['alarm_report'],
    )
    with open(path, 'rb') as f:
        return f.read()


@st.cache_data(ttl=600)
def cached_excel_bytes(target_date, previous_date):
    b = cached_bundle(target_date, previous_date)
    path = rg.generate_excel_report(
        target_date, b['health'], b['scorecards'], b['worst_cells'], b['site_health'],
        b['topology'], b['traffic'], b['site_inventory'], b['freshness'], b['trend'],
        b['alarm_report'],
    )
    with open(path, 'rb') as f:
        return f.read()


@st.cache_data(ttl=600)
def cached_nq_template(interference_period='month'):
    return build_nq_template_report(interference_period=interference_period)


@st.cache_data(ttl=600)
def cached_nq_template_bytes(interference_period='month'):
    import io
    sheets = cached_nq_template(interference_period)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        autofit_excel_columns(writer)
    return buf.getvalue()


@st.cache_data(ttl=600)
def cached_hq_traffic_history():
    return load_hq_traffic_history()


@st.cache_data(ttl=600)
def cached_ept(tech):
    return ept.load_ept(tech)


@st.cache_data(ttl=600)
def cached_ept_validate():
    return ept.validate_ept()


@st.cache_data(ttl=600)
def cached_ept_duplicates(tech):
    return ept.get_ept_duplicates(tech)


@st.cache_data(ttl=600)
def cached_ept_kml_bytes():
    return ept.generate_ept_kml()


@st.cache_data(ttl=600)
def cached_ept_review_list_bytes():
    return ept.generate_review_list_excel()


# ============================================================
# STYLING HELPERS
# ============================================================
def _status_color(val):
    val = str(val)
    if 'FAIL' in val or '🔴' in val:
        return 'background-color:#f8cbcc'
    if 'PASS' in val or '🟢' in val:
        return 'background-color:#c6efce'
    if '🟡' in val:
        return 'background-color:#ffeb9c'
    return ''


def _gap_color(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return ''
    return 'background-color:#f8cbcc' if v < 0 else 'background-color:#c6efce'


def style_status(df, status_col='Status'):
    if df is None or df.empty or status_col not in df.columns:
        return df
    return df.style.map(_status_color, subset=[status_col])


def style_scorecard(df):
    """Colors Status (pass/fail) and Gap (margin sign) - positive margin =
    healthy, negative = failing by that much."""
    if df is None or df.empty:
        return df
    styled = df.style
    if 'Status' in df.columns:
        styled = styled.map(_status_color, subset=['Status'])
    if 'Gap' in df.columns:
        styled = styled.map(_gap_color, subset=['Gap'])
    return styled


def score_icon(score):
    if score >= 95:
        return "🟢", "Good"
    if score >= 90:
        return "🟡", "Fair"
    if score >= 80:
        return "🟠", "Poor"
    return "🔴", "Critical"


def render_trend_chart(tdf, kpi, threshold_info, key):
    fig = go.Figure()
    line_color = '#1f77b4'
    y = tdf[kpi]
    if threshold_info is not None:
        threshold, operator = threshold_info
        try:
            latest = y.dropna().iloc[-1]
            ok, _ = rg.health_checker.check_kpi(latest, threshold, operator)
            line_color = '#2ca02c' if ok else '#d62728'
        except (IndexError, TypeError):
            pass
        fig.add_hline(y=threshold, line_dash='dash', line_color='gray',
                       annotation_text=f'Threshold {operator} {threshold}', annotation_position='top left')
    fig.add_trace(go.Scatter(x=tdf['Date'], y=y, mode='lines+markers', name=kpi,
                              line=dict(color=line_color, width=2)))
    fig.update_layout(title=kpi, height=280, margin=dict(l=30, r=20, t=40, b=30),
                       showlegend=False, hovermode='x unified')
    st.plotly_chart(fig, width='stretch', key=key)


def render_multi_cell_trend_chart(cell_trend_df, kpi, threshold_info, key):
    """Same chart style as render_trend_chart, but one line per cell so
    multiple selected cells can be compared on the same KPI."""
    fig = go.Figure()
    if threshold_info is not None:
        threshold, operator = threshold_info
        fig.add_hline(y=threshold, line_dash='dash', line_color='gray',
                       annotation_text=f'Threshold {operator} {threshold}', annotation_position='top left')
    for cell, cdf in cell_trend_df.groupby('Cell'):
        fig.add_trace(go.Scatter(x=cdf['Date'], y=cdf[kpi], mode='lines+markers', name=cell))
    fig.update_layout(title=kpi, height=300, margin=dict(l=30, r=20, t=40, b=30),
                       hovermode='x unified', legend=dict(font=dict(size=8), orientation='h', y=-0.3))
    st.plotly_chart(fig, width='stretch', key=key)


def render_word_export_button(title, tables, key_prefix, subtitle=""):
    """Small 'export this tab as Word' + download button, for any tab that's
    just one or more tables (no charts) - Worst Cells, Scorecards, Site
    Inventory, Data Freshness, Site Health & Topology, Traffic & Capacity."""
    if st.button(f"📄 Export as Word", key=f"{key_prefix}_word_btn"):
        with st.spinner("Building Word report..."):
            st.session_state[f"{key_prefix}_word_bytes"] = rg.generate_tables_word_report(
                title, tables, subtitle=subtitle)
    if f"{key_prefix}_word_bytes" in st.session_state:
        st.download_button(
            "⬇️ Download", data=st.session_state[f"{key_prefix}_word_bytes"],
            file_name=f"{key_prefix}_{target_date}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"{key_prefix}_word_dl",
        )


def render_excel_export_button(title, tables, key_prefix, subtitle=""):
    """Same idea as render_word_export_button, but for a one-click .xlsx
    export of a tab's underlying tables - one sheet per table - for teams
    that want to filter/pivot the numbers rather than read a document."""
    if st.button("📊 Export as Excel", key=f"{key_prefix}_excel_btn"):
        with st.spinner("Building Excel report..."):
            st.session_state[f"{key_prefix}_excel_bytes"] = rg.generate_tables_excel_report(
                title, tables, subtitle=subtitle)
    if f"{key_prefix}_excel_bytes" in st.session_state:
        st.download_button(
            "⬇️ Download", data=st.session_state[f"{key_prefix}_excel_bytes"],
            file_name=f"{key_prefix}_{target_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_excel_dl",
        )


def render_dp_weekly_report(dp_overview):
    """📱 Device Penetration → Weekly Report: one Monday-Sunday week vs the previous
    archived week, plus the same .xlsx the scraper saves under Reports/."""
    weeks = device_penetration.available_weeks(dp_overview['history'])
    wc1, wc2 = st.columns([2, 3])
    with wc1:
        week_idx = st.selectbox("Week", range(len(weeks)), format_func=lambda i: weeks['label'].iloc[i],
                                key="dp_week")
    with wc2:
        handsets_only = st.toggle("Handsets only for Brand / Model / OS", value=True, key="dp_week_handsets",
                                  help="Excludes IoT/POS modules, routers and datacards. Device Type and "
                                       "Technology always cover all devices.")
    week_start = weeks['week_start'].iloc[week_idx]
    rep = cached_dp_weekly_summary(week_start, handsets_only)

    if rep['days'] < 7:
        st.warning(f"⚠️ Partial week — only {rep['days']}/7 days archived. Averages are over the days present.")
    if rep['prev_label'] is None:
        st.info("No earlier week in the archive — week-over-week changes aren't available.")
    elif not rep['prev_is_adjacent']:
        st.warning(f"⚠️ The adjacent previous week is missing from the archive, so changes are vs "
                   f"**{rep['prev_label']}**.")
    else:
        st.caption(f"Changes are vs {rep['prev_label']}.")

    k, pk = rep['kpis'], rep['prev_kpis'] or {}
    delta = lambda key, fmt: (fmt.format(k[key] - pk[key]) if key in pk else None)
    m = st.columns(5)
    m[0].metric("Avg Daily Devices", f"{k['Avg Daily Devices']:,.0f}", delta("Avg Daily Devices", "{:+,.0f}"))
    m[1].metric("LTE/NR-Capable", f"{k['LTE/NR-Capable %']:.1f}%", delta("LTE/NR-Capable %", "{:+.2f} pp"))
    m[2].metric("5G (NR)-Capable", f"{k['5G (NR)-Capable %']:.1f}%", delta("5G (NR)-Capable %", "{:+.2f} pp"))
    m[3].metric("2G-Only", f"{k['2G-Only %']:.1f}%", delta("2G-Only %", "{:+.2f} pp"), delta_color="inverse")
    m[4].metric("Distinct Models", f"{k['Distinct Models (week)']:,}", delta("Distinct Models (week)", "{:+,}"))
    st.caption("Users are daily unique counts, so weekly figures are the average day, not a sum.")

    def share_bar(df, name, n, key, height):
        top = df.head(n)
        has_delta = 'Δ Share (pp)' in top.columns
        fig = go.Figure(go.Bar(
            x=top['Share %'], y=top[name], orientation='h', text=top['Share %'], texttemplate='%{text:.1f}%',
            customdata=top[['Avg Daily Users'] + (['Δ Share (pp)'] if has_delta else [])],
            hovertemplate='%{y}<br>Share: %{x:.2f}%<br>Avg daily users: %{customdata[0]:,.0f}' +
                          ('<br>Δ vs prev: %{customdata[1]:+.2f} pp' if has_delta else '') + '<extra></extra>'))
        fig.update_layout(height=height, margin=dict(l=200, r=20, t=10, b=30),
                          yaxis=dict(autorange='reversed'), xaxis=dict(ticksuffix='%'))
        st.plotly_chart(fig, width='stretch', key=key)

    b1, b2 = st.columns(2)
    with b1:
        st.subheader("🏷️ Top 10 Brands")
        share_bar(rep['brands'], 'Device Brand', 10, 'dp_wk_brands', 380)
    with b2:
        st.subheader("📡 Highest Supported RAT")
        tech = rep['tech']
        fig = go.Figure(go.Bar(x=tech['Highest Supported RAT'], y=tech['Share %'], text=tech['Share %'],
                               texttemplate='%{text:.1f}%',
                               marker_color=[DP_TECH_COLORS.get(b, '#7f7f7f') for b in tech['Highest Supported RAT']]))
        fig.update_layout(height=380, margin=dict(l=30, r=20, t=10, b=30), yaxis=dict(ticksuffix='%'))
        st.plotly_chart(fig, width='stretch', key='dp_wk_tech')

    st.subheader("📲 Top 15 Device Models")
    share_bar(rep['models'], 'Device Model', 15, 'dp_wk_models', 480)

    if not rep['movers'].empty:
        st.subheader("↕️ Biggest Movers vs Previous Week")
        g1, g2 = st.columns(2)
        mv = rep['movers']
        with g1:
            st.caption("▲ Gainers (avg daily users)")
            st.dataframe(mv[mv['Movement'] == '▲ Gainer'].drop(columns='Movement').head(10),
                         width='stretch', hide_index=True)
        with g2:
            st.caption("▼ Losers (avg daily users)")
            st.dataframe(mv[mv['Movement'] == '▼ Loser'].drop(columns='Movement').head(10),
                         width='stretch', hide_index=True)

    with st.expander("📋 Full tables (brands, models, OS, device type)"):
        t = st.tabs(["Brands", "Models", "OS", "Device Type"])
        for tab, key in zip(t, ['brands', 'models', 'os', 'types']):
            with tab:
                st.dataframe(rep[key], width='stretch', hide_index=True, height=400)

    st.download_button(
        "⬇️ Download Weekly Report (.xlsx)",
        data=cached_dp_weekly_report_bytes(week_start, handsets_only),
        file_name=device_penetration.report_filename(week_start),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dp_wk_download",
    )
    st.caption("Same workbook the weekly scraper saves automatically to "
               "Weekly_Device_Penetration_Exports\\Reports\\ — Summary, Brands, Models, Movers, OS, "
               "Device Type, Technology (with charts) and the full model list.")


# ============================================================
# HEADER + SIDEBAR
# ============================================================
st.markdown("""
<div class="header-container">
    <h1 style="margin:0;">📊 Libyana Network Performance Dashboard</h1>
    <p style="margin:0;">2G / 3G / 4G — live from the daily KPI pipeline</p>
</div>
""", unsafe_allow_html=True)

all_dates = cached_dates()
if not all_dates:
    st.error("❌ No processed data found in output/csv/. Run the scheduler at least once.")
    st.stop()

with st.sidebar:
    st.header("📅 Controls")
    target_date = st.selectbox("Report date", all_dates, index=0)
    prev_candidates = [d for d in all_dates if d < target_date]
    previous_date = prev_candidates[0] if prev_candidates else target_date
    st.caption(f"Compared against: {previous_date}")
    st.caption(f"📊 {len(all_dates)} day(s) of history available")
    st.divider()
    if st.button("🔄 Refresh data", width='stretch'):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.header("🧭 Navigation")
    section = st.radio(
        "Section",
        ["📊 Overview", "📡 KPIs & Performance", "🏗️ Sites & Infrastructure",
         "📶 Packet Loss", "🚨 Alarms", "📱 CEM", "📞 User Complaint", "🔎 Investigate",
         "📋 HQ Reports", "📧 Reports"],
        key="nav_section", label_visibility="collapsed",
    )

bundle = cached_bundle(target_date, previous_date)
health = bundle['health']

# ============================================================
# 📊 OVERVIEW — Site Summary, Executive Summary
# ============================================================
if section == "📊 Overview":
    sec_tabs = st.tabs(["🏢 Site Summary", "📋 Executive Summary"])

    with sec_tabs[0]:
        cards = bundle['site_cards']

        def fmt_num(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return "N/A"
            return f"{v:,.0f}"

        st.subheader("Sites On Air")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Sites (2G+3G+4G)", fmt_num(cards.get('Total Sites (2G+3G+4G)')))
        c2.metric("2G Sites", fmt_num(cards.get('2G Sites')))
        c3.metric("3G Sites", fmt_num(cards.get('3G Sites')))
        c4.metric("4G Sites", fmt_num(cards.get('4G Sites')))

        st.subheader("Users & Traffic")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total PS Users", fmt_num(cards.get('Total PS Users')))
        c2.metric("Total CS Users", fmt_num(cards.get('Total CS Users')))
        c3.metric("VoLTE Users", fmt_num(cards.get('VoLTE Users')))
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Subscribers", fmt_num(cards.get('Total Subscribers')))
        c2.metric("Total PS Traffic (GB)", fmt_num(cards.get('Total PS Traffic (GB)')))
        c3.metric("Total CS Traffic (Erl)", fmt_num(cards.get('Total CS Traffic (Erl)')))

        st.subheader("Multi-RAT Site Composition")
        st.dataframe(bundle['site_inventory'], width='stretch', hide_index=True)
        st.caption(f"As of {target_date}")

    with sec_tabs[1]:
        icon, status = score_icon(health.get('overall_score', 0))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Overall Health", f"{health.get('overall_score', 0):.1f}%", delta=f"{icon} {status}")
        c2.metric("KPIs Passed", health.get('passed_kpis', 0))
        c3.metric("KPIs Failed", health.get('failed_kpis', 0))
        freshness = bundle['freshness']
        stale = freshness[freshness['Status'].str.contains('🔴|🟡')] \
            if freshness is not None and not freshness.empty else pd.DataFrame()
        c4.metric("Data Freshness", "🟢 Current" if stale.empty else f"⚠️ {len(stale)} behind")

        st.subheader("Score by Technology")
        tcols = st.columns(len(TECH_LABELS))
        for col, (tech, label) in zip(tcols, TECH_LABELS.items()):
            s = health.get('by_technology', {}).get(tech, {}).get('score', 0)
            ic, st_ = score_icon(s)
            col.metric(label, f"{s:.1f}%", delta=f"{ic} {st_}")

        st.subheader("Top 3 Worst Cells")
        combined = rg.combine_worst_cells(bundle['worst_cells'], n_top=3)
        if not combined.empty:
            st.dataframe(combined, width='stretch', hide_index=True)
        else:
            st.info("No cells flagged.")

        st.subheader("Failing Network KPIs (Busy Hour)")
        alerts = health.get('alerts', [])
        if alerts:
            alerts_df = pd.DataFrame(alerts).rename(columns={
                'kpi': 'KPI Name', 'tech': 'Technology', 'value': "Today's Value",
                'threshold': 'Threshold', 'possible_cause': 'Possible Cause', 'severity': 'Severity',
            })
            st.dataframe(alerts_df, width='stretch', hide_index=True)
        else:
            st.success("No failing whole-network KPIs for this date.")

# ============================================================
# 📡 KPIs & PERFORMANCE — Scorecards, Worst Cells, Traffic & Capacity,
# 14-Day Trend, Data Freshness
# ============================================================
elif section == "📡 KPIs & Performance":
    sec_tabs = st.tabs(["📡 Scorecards", "⚠️ Worst Cells", "📶 Traffic & Capacity",
                         "📈 14-Day Trend", "🕒 Data Freshness"])

    with sec_tabs[0]:
        sub = st.tabs(list(TECH_LABELS.values()))
        for sub_tab, (tech, label) in zip(sub, TECH_LABELS.items()):
            with sub_tab:
                s = health.get('by_technology', {}).get(tech, {}).get('score', 0)
                ic, st_ = score_icon(s)
                st.markdown(f"**Score: {s:.1f}% ({ic} {st_})**")
                df = bundle['scorecards'].get(tech)
                if df is not None and not df.empty:
                    st.dataframe(style_scorecard(df), width='stretch', hide_index=True)
                else:
                    st.warning("No scorecard data for this date.")

        render_word_export_button(
            "Technology Scorecards", [(TECH_LABELS[t], bundle['scorecards'].get(t)) for t in TECH_LABELS],
            key_prefix="scorecards", subtitle=f"Busy Hour — {target_date}",
        )

    with sec_tabs[1]:
        wc_col1, wc_col2 = st.columns([1, 3])
        with wc_col1:
            wc_n_top = st.number_input("Top N worst cells", min_value=1, max_value=100, value=10, step=1,
                                        key="wc_n_top")
        worst_cells_n = cached_worst_cells(target_date, int(wc_n_top))

        sub = st.tabs(list(TECH_LABELS.values()))
        for sub_tab, (tech, label) in zip(sub, TECH_LABELS.items()):
            with sub_tab:
                df = worst_cells_n.get(tech)
                if df is not None and not df.empty:
                    st.dataframe(df, width='stretch', hide_index=True)
                else:
                    st.info("No cells flagged for this technology/date.")

        render_word_export_button(
            "Worst Cells", [(TECH_LABELS[t], worst_cells_n.get(t)) for t in CELL_SHEETS],
            key_prefix="worst_cells", subtitle=f"Top {int(wc_n_top)} per technology — {target_date}",
        )

    with sec_tabs[2]:
        st.dataframe(bundle['traffic'], width='stretch', hide_index=True)
        render_word_export_button(
            "Traffic & Capacity", [("Traffic & Capacity", bundle['traffic'])],
            key_prefix="traffic", subtitle=target_date,
        )

        gi_df = cached_sheet('Gi_Interface_Traffic')
        if gi_df is not None and not gi_df.empty:
            with st.expander("🌐 Core Network — Gi/Sgi Traffic Trend (14-day)", expanded=False):
                gi_trend = gi_df.copy()
                gi_trend['Date'] = pd.to_datetime(gi_trend['Date'], errors='coerce')
                gi_trend = gi_trend.sort_values('Date').tail(14)
                gi_cols = [
                    'PS core_Gi _23G traffic (TB)', 'PS 4G Sgi_traffic (TB)',
                    'PS 234G traffic (TB)', 'PGW-C current VoLTE IMS subscribers(number)',
                ]
                gi_cols = [c for c in gi_cols if c in gi_trend.columns]
                gi_table = gi_trend[['Date'] + gi_cols].copy()
                gi_table['Date'] = gi_table['Date'].dt.strftime('%Y-%m-%d')

                st.dataframe(gi_table, width='stretch', hide_index=True)

                cols = st.columns(2)
                for i, kpi in enumerate(gi_cols):
                    with cols[i % 2]:
                        render_trend_chart(gi_trend, kpi, None, key=f"trend_gi_{kpi}")

                render_excel_export_button(
                    "Core Network — Gi/Sgi Traffic Trend",
                    [("Gi_Sgi_Traffic_Trend", gi_table)],
                    key_prefix="gi_traffic_trend", subtitle=f"Last {len(gi_table)} day(s) — {target_date}",
                )

    with sec_tabs[3]:
        sub = st.tabs(list(TECH_LABELS.values()))
        for sub_tab, (tech, label) in zip(sub, TECH_LABELS.items()):
            with sub_tab:
                tdf = bundle['trend'].get(tech)
                if tdf is None or tdf.empty:
                    st.info("No trend data available.")
                    continue
                kpi_thresholds = cached_trend_thresholds(tech)
                kpi_cols = [c for c in tdf.columns if c != 'Date']
                threshold_kpis = [c for c in kpi_cols if c in kpi_thresholds]
                other_kpis = [c for c in kpi_cols if c not in kpi_thresholds]

                st.caption(f"{len(tdf)} day(s) of history, ending {target_date}. "
                           "Dashed line = threshold; the trend line turns red when the latest value fails it.")

                with st.expander(f"Threshold KPIs ({len(threshold_kpis)})", expanded=True):
                    cols = st.columns(2)
                    for i, kpi in enumerate(threshold_kpis):
                        with cols[i % 2]:
                            render_trend_chart(tdf, kpi, kpi_thresholds.get(kpi), key=f"trend_{tech}_{kpi}")

                if other_kpis:
                    with st.expander(f"Traffic & Capacity ({len(other_kpis)})", expanded=False):
                        cols = st.columns(2)
                        for i, kpi in enumerate(other_kpis):
                            with cols[i % 2]:
                                render_trend_chart(tdf, kpi, None, key=f"trend_{tech}_{kpi}")

    with sec_tabs[4]:
        st.dataframe(style_status(bundle['freshness']), width='stretch', hide_index=True)
        render_word_export_button(
            "Data Freshness", [("Data Freshness", bundle['freshness'])],
            key_prefix="freshness", subtitle=target_date,
        )

# ============================================================
# 🏗️ SITES & INFRASTRUCTURE — Site Health & Topology, Site Inventory,
# Site Detail, EPT
# ============================================================
elif section == "🏗️ Sites & Infrastructure":
    sec_tabs = st.tabs(["🏗️ Site Health & Topology", "🏢 Site Inventory", "📋 Site Detail", "🗺️ EPT"])

    with sec_tabs[0]:
        st.dataframe(bundle['site_health'], width='stretch', hide_index=True)
        topo = bundle['topology']
        topo_col, refresh_col = st.columns([4, 1])
        with topo_col:
            if topo.get('loaded'):
                st.success(
                    f"Topology reference loaded: {topo['nodes']} nodes "
                    f"({topo['fn_count']} FN, {topo['hub_count']} HUB), "
                    f"{topo['site_relationships']} site relationships, "
                    f"regions: {', '.join(topo['regions'])}."
                )
            else:
                st.warning("Topology reference file not found.")
        with refresh_col:
            xlsx_path = find_topology_xlsx()
            if st.button("🔄 Rebuild from Excel", width='stretch',
                         help=f"Re-parse {os.path.basename(xlsx_path) if xlsx_path else 'the FN-HUB Excel file'} "
                              "in config/ - use after updating it.",
                         disabled=not xlsx_path):
                with st.spinner("Re-parsing topology Excel reference..."):
                    build_site_topology_csv()
                cached_bundle.clear()
                cached_topology_table.clear()
                cached_site_topology.clear()
                st.rerun()
        st.caption("Alarm-to-topology impact correlation is Phase 3/4, pending NetEco integration.")
        render_word_export_button(
            "Site Health & Topology", [("Availability by Technology", bundle['site_health'])],
            key_prefix="site_health", subtitle=target_date,
        )

        if topo.get('loaded'):
            st.divider()
            st.subheader("🔗 FN/HUB Site Lookup")
            st.caption("Which FN/HUB node a site depends on, and which other sites share it - useful before "
                       "investigating several problem sites as independent RF issues.")
            topo_search = st.text_input("Site Name contains...", "", key="topo_search")
            all_topo = cached_topology_table()
            if all_topo is not None and not all_topo.empty:
                if topo_search:
                    site_matches = sorted(all_topo[all_topo['Connected_Site'].str.contains(
                        topo_search, case=False, na=False)]['Connected_Site'].unique().tolist())
                    lookup_result = cached_site_topology(tuple(site_matches)) if site_matches else pd.DataFrame()
                    if lookup_result.empty:
                        st.info("No matching site found in the topology reference.")
                    else:
                        st.dataframe(lookup_result, width='stretch', hide_index=True)
                with st.expander(f"📋 Browse all {len(all_topo):,} relationships"):
                    st.dataframe(all_topo, width='stretch', hide_index=True, height=400)

    with sec_tabs[1]:
        st.dataframe(bundle['site_inventory'], width='stretch', hide_index=True)
        st.caption("Site outages: pending NetEco alarm integration (Phase 3).")
        render_word_export_button(
            "Site Inventory", [("Multi-RAT Site Composition", bundle['site_inventory'])],
            key_prefix="site_inventory", subtitle=target_date,
        )

    with sec_tabs[2]:
        sd_summary = cached_site_detail_summary()
        if not sd_summary.get('loaded'):
            st.warning("SiteDetail.csv not found.")
        else:
            refreshed = sd_summary.get('file_refreshed')
            latest_change = sd_summary.get('latest_change')
            st.caption(
                f"Last refreshed by daily pipeline: **{refreshed.strftime('%Y-%m-%d %H:%M') if refreshed else 'unknown'}**"
                f"  ·  Latest site config change: **{latest_change.strftime('%Y-%m-%d') if latest_change is not None else 'unknown'}**",
                help="The pipeline re-checks every site daily, but a site's 'Last Updated' only moves "
                     "when its bands/sectors/RAT actually change (or it is new).",
            )
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Sites", f"{sd_summary['total_sites']:,}")
            m2.metric("Multi-Carrier LTE Sites", f"{sd_summary['multi_carrier_lte_sites']:,}")
            m3.metric("Single-Carrier LTE Sites", f"{sd_summary['single_carrier_lte_sites']:,}",
                      help="Only one 4G band slot active - the first place to look for a free-capacity "
                           "upgrade if a site here is also PRB-congested (see Investigate tab).")
            m4.metric("No LTE Yet", f"{sd_summary['no_lte_sites']:,}")

            bc1, bc2, bc3 = st.columns(3)
            with bc1:
                st.caption("RAT combination")
                st.dataframe(pd.Series(sd_summary['rat_counts'], name='Sites').sort_values(ascending=False),
                             width='stretch')
            with bc2:
                st.caption("Scenario")
                st.dataframe(pd.Series(sd_summary['scenario_counts'], name='Sites').sort_values(ascending=False),
                             width='stretch')
            with bc3:
                st.caption("4G band adoption (sites with slot active)")
                st.dataframe(pd.Series(sd_summary['band_adoption'], name='Sites').sort_values(ascending=False),
                             width='stretch')

            recent = sd_summary.get('recent_changes')
            if recent is not None and not recent.empty:
                with st.expander(f"🆕 Recently changed / new sites ({len(recent)} in the 30 days up to latest change)"):
                    st.dataframe(recent, width='stretch', hide_index=True)

            st.divider()
            sd_search = st.text_input("Filter by Site Name contains...", "", key="site_detail_search")
            sd_table = cached_sheet('SiteDetail')
            if sd_table is not None and not sd_table.empty:
                if sd_search:
                    sd_table = sd_table[sd_table['Site Name'].str.contains(sd_search, case=False, na=False)]
                st.caption(f"{len(sd_table):,} site(s)")
                st.dataframe(sd_table, width='stretch', hide_index=True, height=450)
                render_word_export_button(
                    "Site Detail", [("Site Detail", sd_table)],
                    key_prefix="site_detail_all", subtitle=target_date,
                )

    with sec_tabs[3]:
        ept_last_updated = ept.get_ept_last_updated()
        ept_file_bytes = ept.get_ept_file_bytes()

        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            if ept_last_updated:
                st.caption(f"EPT file last updated: **{ept_last_updated.strftime('%Y-%m-%d %H:%M')}**")
            else:
                st.error("No EPT file found in config/ (expected 'Libyana MS EPT_*-Whole Network.xlsx').")
        with c2:
            if ept_file_bytes:
                st.download_button(
                    "⬇️ Download EPT (Excel)", data=ept_file_bytes,
                    file_name=os.path.basename(ept.find_ept_file()),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch', key="ept_excel_dl",
                )
        with c3:
            if st.button("🗺️ Prepare KML (Google Earth)", width='stretch'):
                with st.spinner("Building KML (all technologies)..."):
                    st.session_state['ept_kml_bytes'] = cached_ept_kml_bytes()
        if 'ept_kml_bytes' in st.session_state:
            st.download_button(
                "⬇️ Download KML", data=st.session_state['ept_kml_bytes'],
                file_name="Libyana_EPT_Whole_Network.kml",
                mime="application/vnd.google-earth.kml+xml",
                width='stretch', key="ept_kml_dl",
            )

        st.divider()
        dq_col, dq_btn_col = st.columns([3, 1])
        with dq_col:
            st.subheader("🔍 Data Quality vs. System (Ground Truth)")
            st.caption("EPT is manually maintained by RF engineers and can contain typos/stale rows. "
                       "This compares it against the cells actually reporting KPIs in the pipeline.")
        with dq_btn_col:
            st.download_button(
                "⬇️ Download Review List", data=cached_ept_review_list_bytes(),
                file_name="EPT_Review_List.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch', key="ept_review_dl",
                help="Duplicates, EPT-only, and missing-from-EPT rows per technology, ready to work through.",
            )
        st.dataframe(cached_ept_validate(), width='stretch', hide_index=True)

        ept_tech = st.selectbox("Technology", list(CELL_SHEETS.keys()),
                                 format_func=lambda t: TECH_LABELS[t], key="ept_tech")
        ept_dupes = cached_ept_duplicates(ept_tech)
        if ept_dupes is not None and not ept_dupes.empty:
            with st.expander(f"⚠️ {ept_dupes['Cell Name'].nunique()} Cell Name(s) with conflicting duplicate rows "
                              f"({len(ept_dupes)} rows)"):
                st.caption("Same Cell Name appears more than once with different values — needs an RF engineer "
                           "to confirm which row is correct.")
                st.dataframe(ept_dupes, width='stretch', hide_index=True)

        st.divider()
        ept_df = cached_ept(ept_tech)
        if ept_df is None or ept_df.empty:
            st.warning(f"No EPT data found for {TECH_LABELS[ept_tech]}.")
        else:
            ept_search = st.text_input("Filter by Site Name / Cell Name / City contains...", key="ept_search")
            ept_filtered = ept_df
            if ept_search:
                name_cols = [c for c in ['Site Name', 'Cell Name', 'City'] if c in ept_df.columns]
                mask = False
                for c in name_cols:
                    mask = mask | ept_filtered[c].astype(str).str.contains(ept_search, case=False, na=False)
                ept_filtered = ept_filtered[mask]
            st.caption(f"{len(ept_filtered):,} row(s)")

            map_col, table_col = st.columns([1, 2])
            with map_col:
                if 'Latitude' in ept_filtered.columns and 'Longitude' in ept_filtered.columns:
                    map_df = ept_filtered[['Latitude', 'Longitude']].dropna().rename(
                        columns={'Latitude': 'lat', 'Longitude': 'lon'})
                    if not map_df.empty:
                        st.map(map_df, size=15, zoom=6)
            with table_col:
                st.dataframe(ept_filtered, width='stretch', hide_index=True, height=400)

# ============================================================
# 📶 PACKET LOSS — IUB/ABIS backhaul ping quality. Archives built by
# backend/transmission_kpi_processor.py, logic in
# backend/packet_loss_engine.py, thresholds in config/packet_loss_rules.csv.
# ============================================================
elif section == "📶 Packet Loss":
    from backend.packet_loss_engine import CLASS_ORDER, FLAGGED_CLASSES

    st.caption("Which sites suffer from backhaul (ABIS 2G / IUB 3G) packet loss, and why. Sites are ranked by "
               "**how many hours** they lost packets — not by the daily average, which hides intermittent "
               "problems. No-response (NIL) hours are counted separately, never as 0%.")

    pl_dates = cached_pl_dates()
    if not pl_dates:
        st.info("No packet loss history yet — run the pipeline (Step 5) or "
                "`python -m backend.transmission_kpi_processor <FTP root> --backfill`.")
    else:
        rules = pl_rules()
        pl_latest = max([d for d in pl_dates if d <= str(target_date)] or [pl_dates[-1]])
        pl_view = st.radio("View", ["🌐 Network view", "📝 Special report (selected sites)"],
                           horizontal=True, key="pl_view")

        CLASS_HELP = (f"🔴 **Chronic** — lost ≥{rules['loss_hour_pct']:g}% in ≥{rules['chronic_share']:.0%} of hours · "
                      f"🟠 **Recurrent** — loss hours on ≥{int(rules['recurrent_min_days'])} different days · "
                      "🟡 **Sporadic** — isolated loss hours · ⚫ **Outage only** — no-response hours, no loss · "
                      "⚪ **Not reporting** — no ping reply all period · 🟢 **Clean**")

        def _pl_export_buttons(start, end, sites, label, cls, hubs, key_prefix):
            c_x, c_w = st.columns(2)
            with c_x:
                if st.button("📊 Prepare Excel report", key=f"{key_prefix}_xl_btn", width='stretch'):
                    with st.spinner("Building Excel report..."):
                        st.session_state[f"{key_prefix}_xl"] = rg.build_packet_loss_excel(
                            start, end, list(sites) if sites else None, label)
                if f"{key_prefix}_xl" in st.session_state:
                    st.download_button(
                        "⬇️ Download Excel", data=st.session_state[f"{key_prefix}_xl"],
                        file_name=f"Packet_Loss_{(label or 'Network').replace(' ', '_')[:40]}_{start}_{end}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"{key_prefix}_xl_dl", width='stretch')
            with c_w:
                summary = cls['Class'].value_counts().reindex(CLASS_ORDER, fill_value=0) \
                    .rename_axis('Class').reset_index(name='Sites')
                word_cols = ['Class', 'Site', 'Region', 'Suspected Cause', 'Pattern', 'Affected RAT',
                             'Loss Hours', 'Loss Days', 'Avg Loss (%)', 'No Response Hours']
                shown = cls if sites else cls[cls['Class'].isin(FLAGGED_CLASSES)]
                render_word_export_button(
                    f"Packet Loss — {label or 'Whole network'}",
                    [("Summary", summary), ("Sites", shown[word_cols]),
                     ("Hub / shared-path events", hubs.head(30) if hubs is not None else None)],
                    key_prefix=f"{key_prefix}_word", subtitle=f"{start} to {end}")

        # ---------------------------------------------------------- network view
        if pl_view.startswith("🌐"):
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                pl_period = st.radio("Period", ["Day", "Last 7 Days", "Last 30 Days", "Custom range"],
                                     index=1, horizontal=True, key="pl_period")
            if pl_period == "Custom range":
                with c2:
                    pl_start = st.selectbox("Start", pl_dates, index=max(0, len(pl_dates) - 7), key="pl_c_start")
                with c3:
                    pl_end = st.selectbox("End", pl_dates, index=pl_dates.index(pl_latest), key="pl_c_end")
            else:
                span = {"Day": 1, "Last 7 Days": 7, "Last 30 Days": 30}[pl_period]
                pl_end = pl_latest
                pl_start = (pd.Timestamp(pl_end) - pd.Timedelta(days=span - 1)).strftime('%Y-%m-%d')

            if pl_start > pl_end:
                st.error("Start date is after end date.")
            else:
                cls_all = cached_pl_sites(pl_start, pl_end)
                hubs = cached_pl_hubs(pl_start, pl_end)
                if cls_all is None or cls_all.empty:
                    st.info("No packet loss data for this period.")
                else:
                    counts = cls_all['Class'].value_counts()
                    tiles = st.columns(6)
                    for tile, c in zip(tiles, CLASS_ORDER):
                        tile.metric(c, f"{int(counts.get(c, 0)):,}")
                    st.caption(f"{pl_start} → {pl_end} · {len(cls_all):,} sites · "
                               f"{0 if hubs is None else len(hubs)} FN/HUB node(s) with shared-path events. "
                               + CLASS_HELP)

                    t_sites, t_hubs, t_trend, t_other = st.tabs(
                        ["🏚️ Affected sites", "🔗 Hub / shared-path events", "📈 Network trend",
                         "🔌 Core links & not reporting"])

                    with t_sites:
                        f1, f2, f3 = st.columns([2, 2, 2])
                        with f1:
                            cls_pick = st.multiselect("Class", CLASS_ORDER, default=CLASS_ORDER[:4], key="pl_cls")
                        with f2:
                            regions = sorted(cls_all['Region'].dropna().unique().tolist())
                            reg_pick = st.multiselect("Region", regions, key="pl_region")
                        with f3:
                            pl_search = st.text_input("Site / hub contains...", "", key="pl_search")
                        view = cls_all[cls_all['Class'].isin(cls_pick)] if cls_pick else cls_all
                        if reg_pick:
                            view = view[view['Region'].isin(reg_pick)]
                        if pl_search:
                            q = pl_search.strip()
                            view = view[view['Site'].str.contains(q, case=False, na=False)
                                        | view['FN/HUB Chain'].fillna('').str.contains(q, case=False)
                                        | view['Top Hub Node'].fillna('').str.contains(q, case=False)]
                        st.dataframe(
                            view, width='stretch', hide_index=True, height=430,
                            column_config={
                                'Persistence (%)': st.column_config.ProgressColumn(
                                    'Persistence (%)', min_value=0, max_value=100, format="%.0f%%",
                                    help="Share of measured hours with loss ≥ the loss-hour threshold"),
                                'Suspected Cause': st.column_config.TextColumn(width='large'),
                            })
                        st.caption(f"{len(view):,} site(s) shown")

                        if not view.empty:
                            st.markdown("**🔍 Site drill-down (hour by hour)**")
                            dd_site = st.selectbox("Site", view['Site'].tolist(), key="pl_dd_site")
                            render_pl_site_drilldown(dd_site, pl_start, pl_end, "pl_dd")

                    with t_hubs:
                        st.caption(f"One row per FN/HUB node where ≥{int(rules['hub_min_sites'])} of its sites "
                                   f"(and ≥{rules['hub_min_share']:.0%} of them) were hit in the same hour. Each "
                                   "affected hour is credited to the highest failing node in the chain, so a hub "
                                   "outage is one ticket here, not one per downstream site.")
                        if hubs is None or hubs.empty:
                            st.success("No shared-path events in this period.")
                        else:
                            st.dataframe(hubs, width='stretch', hide_index=True, height=400,
                                         column_config={'Affected Sites': st.column_config.TextColumn(width='large')})

                    with t_trend:
                        tr_start = min(pl_start, (pd.Timestamp(pl_end) - pd.Timedelta(days=29)).strftime('%Y-%m-%d'))
                        trend = cached_pl_trend(tr_start, pl_end)
                        if trend is None or trend.empty:
                            st.info("No trend data.")
                        else:
                            series = ['Sites with 3+ loss hours', 'Sites with no-response hours', 'Hub events']
                            fig = go.Figure()
                            for color, s in zip(PL_SERIES, series):
                                if s in trend.columns:
                                    fig.add_trace(go.Scatter(x=trend['Date'], y=trend[s], name=s, mode='lines+markers',
                                                             line=dict(color=color, width=2), marker=dict(size=8)))
                            fig.update_layout(height=340, hovermode='x unified', margin=dict(l=10, r=10, t=10, b=10),
                                              legend=dict(orientation='h', y=1.12), yaxis_title='Count')
                            st.plotly_chart(fig, width='stretch', key="pl_trend_chart")
                            st.caption(f"Daily counts, {tr_start} → {pl_end}. A jump in no-response sites on one day "
                                       "across many hubs usually means a power/regional event — check the Alarms tab.")
                            with st.expander("Trend table"):
                                st.dataframe(trend, width='stretch', hide_index=True)

                    with t_other:
                        st.markdown("**⚪ Not reporting** — no ping reply for the whole period "
                                    "(link down, site off-air, test/decommissioned node or ping not configured)")
                        nr = cls_all[cls_all['Class'] == '⚪ Not reporting'][
                            ['Site', 'Region', 'GBSC', 'RATs', 'No Response Hours', 'FN/HUB Chain']]
                        st.dataframe(nr, width='stretch', hide_index=True)
                        st.markdown("**🔌 Core-network links** (IUR / A / IUCS / IUPS) — not site backhaul, "
                                    "kept out of the site ranking")
                        st.dataframe(cached_pl_core(pl_start, pl_end), width='stretch', hide_index=True)

                    st.divider()
                    _pl_export_buttons(pl_start, pl_end, None, "", cls_all, hubs, "pl_net")

        # ---------------------------------------------------------- special report
        else:
            st.caption("Build a packet loss report for a specific set of sites over any date range — e.g. a "
                       "transmission ticket, a hub/FN, or a region you're verifying after a fix.")
            site_list = cached_pl_site_list()

            def _pl_add_sites(new_sites):
                cur = list(st.session_state.get('pl_sr_sites', []))
                st.session_state['pl_sr_sites'] = cur + [s for s in new_sites if s not in cur]

            a1, a2 = st.columns(2)
            with a1:
                hub_nodes = sorted({n.split(' (')[0] for v in site_list['FN/HUB Chain'].dropna()
                                    for n in v.split(', ')})
                hub_pick = st.selectbox("Add all sites of an FN/HUB node", [""] + hub_nodes, key="pl_sr_hub")
                if hub_pick:
                    hub_sites = site_list.loc[site_list['FN/HUB Chain'].fillna('').str.contains(
                        rf"(?:^|, ){hub_pick} \(", regex=True), 'Site'].tolist()
                    if hub_pick in set(site_list['Site']):
                        hub_sites = [hub_pick] + hub_sites
                    st.button(f"➕ Add {len(hub_sites)} site(s) of {hub_pick}", key="pl_sr_hub_add",
                              on_click=_pl_add_sites, args=(hub_sites,))
            with a2:
                reg_opts = sorted(site_list['Region'].dropna().unique().tolist())
                reg_pick = st.selectbox("Add all sites of a region", [""] + reg_opts, key="pl_sr_region")
                if reg_pick:
                    reg_sites = site_list.loc[site_list['Region'] == reg_pick, 'Site'].tolist()
                    st.button(f"➕ Add {len(reg_sites)} site(s) of {reg_pick}", key="pl_sr_reg_add",
                              on_click=_pl_add_sites, args=(reg_sites,))

            sr_search = st.text_input("Filter site list contains...", "", key="pl_sr_search")
            options = site_list['Site'].tolist()
            if sr_search:
                options = [o for o in options if sr_search.lower() in o.lower()]
            options = sorted(set(options) | set(st.session_state.get('pl_sr_sites', [])))
            sr_sites = st.multiselect("Sites (e.g. MRDH001, KUFR004...)", options, key="pl_sr_sites")

            d1, d2, d3 = st.columns([1, 1, 2])
            with d1:
                sr_start = st.selectbox("Start date", pl_dates, index=max(0, len(pl_dates) - 7), key="pl_sr_start")
            with d2:
                sr_end = st.selectbox("End date", pl_dates, index=pl_dates.index(pl_latest), key="pl_sr_end")
            with d3:
                sr_label = st.text_input('Report label (e.g. "DARN009 MW link - TT#1234")', key="pl_sr_label")

            if not sr_sites:
                st.caption("Pick sites (or add a whole FN/HUB node / region) and a date range to preview and export.")
            elif sr_start > sr_end:
                st.error("Start date is after end date.")
            else:
                sel = tuple(sr_sites)
                cls_sel = cached_pl_sites(sr_start, sr_end, sel)
                hubs_sel = cached_pl_hubs(sr_start, sr_end, sel)
                if cls_sel is None or cls_sel.empty:
                    st.info("No packet loss data for these sites in this date range.")
                else:
                    counts = cls_sel['Class'].value_counts()
                    tiles = st.columns(6)
                    for tile, c in zip(tiles, CLASS_ORDER):
                        tile.metric(c, f"{int(counts.get(c, 0)):,}")
                    st.caption(f"{sr_label + ' · ' if sr_label else ''}{len(cls_sel)} site(s) · "
                               f"{sr_start} → {sr_end}. " + CLASS_HELP)

                    st.subheader("📋 Site classification")
                    st.dataframe(cls_sel, width='stretch', hide_index=True,
                                 column_config={'Persistence (%)': st.column_config.ProgressColumn(
                                     'Persistence (%)', min_value=0, max_value=100, format="%.0f%%")})

                    st.subheader("🗓️ Loss hours per day")
                    mat = cached_pl_matrix(sr_start, sr_end, sel)
                    if mat is not None and not mat.empty:
                        days = [c for c in mat.columns if c not in ('Site', 'Total')]
                        fig = _pl_heatmap(mat[days].values, days, mat['Site'].tolist(), 24,
                                          "Loss hours", "%{z} loss hour(s)")
                        st.plotly_chart(fig, width='stretch', key="pl_sr_matrix")
                        st.caption("Hours per day with loss ≥ "
                                   f"{rules['loss_hour_pct']:g}% (0–24). Rows sorted by total.")

                    if hubs_sel is not None and not hubs_sel.empty:
                        st.subheader("🔗 Shared-path events touching these sites")
                        st.dataframe(hubs_sel, width='stretch', hide_index=True)

                    st.subheader("🔍 Hour-by-hour drill-down")
                    sr_dd = st.selectbox("Site", cls_sel['Site'].tolist(), key="pl_sr_dd")
                    render_pl_site_drilldown(sr_dd, sr_start, sr_end, "pl_sr_dd")

                    st.divider()
                    _pl_export_buttons(sr_start, sr_end, sel, sr_label.strip() or ', '.join(sr_sites[:3]) +
                                       (f" +{len(sr_sites) - 3} more" if len(sr_sites) > 3 else ''),
                                       cls_sel, hubs_sel, "pl_sr")

# ============================================================
# 🚨 ALARMS — live site alarm status + historical NOC rollups
# (read-only feed from a sibling NOC Automation Suite; see
# backend/noc_alarm_processor.py)
# ============================================================
elif section == "🚨 Alarms":
    st.caption("Live site alarm status and historical chronic-offender/downtime trends from the NOC "
               "alarm pipeline (MAE/NetEco/NCE). This reads a separate automation suite's output "
               "read-only - if that suite isn't installed or running on this machine, this section "
               "stays empty instead of erroring.")

    alarm_overview = cached_alarm_overview()
    live, hist = alarm_overview['live'], alarm_overview['historical']

    if not live.get('loaded') and not hist.get('loaded'):
        st.info("No NOC alarm feed found. This section reads the output of a separate NOC automation "
                "suite (MAE/NetEco/NCE alarm scrapers) when it's installed and running on this machine.")
    else:
        st.subheader("🔌 Currently Disconnected Sites")
        if live.get('loaded'):
            age = live['age_minutes']
            if live.get('is_stale'):
                st.warning(f"⚠️ Live alarm feed last updated {age:.0f} min ago — the alarm scraper "
                           f"pipeline may be stalled; treat current-alarm counts below as possibly stale.")
            else:
                st.caption(f"✅ Live feed updated {age:.0f} min ago" if age is not None else "✅ Live feed found")

            alarms_df = live['alarms']
            live_metrics = live.get('metrics', {})
            m1, m2, m3 = st.columns(3)
            m1.metric("Count of NE Is Disconnected", f"{live_metrics.get('ne_disconnected', len(alarms_df)):,}")
            m2.metric("Count of Mains Failure", f"{live_metrics.get('mains_failure', 0):,}")
            m3.metric("Count of NCE Transmission Alarm Sites", f"{live_metrics.get('nce_transmission', 0):,}")

            alarm_search = st.text_input("Filter by Site Name contains...", "", key="alarm_search")
            alarms_filtered = alarms_df
            if alarm_search and 'Site Name' in alarms_filtered.columns:
                alarms_filtered = alarms_filtered[
                    alarms_filtered['Site Name'].astype(str).str.contains(alarm_search, case=False, na=False)]
            st.dataframe(alarms_filtered, width='stretch', hide_index=True, height=350)
            render_word_export_button(
                "Currently Disconnected Sites", [("Disconnected Sites", alarms_filtered)],
                key_prefix="live_disconnected", subtitle=pd.Timestamp.now().strftime('%Y-%m-%d %H:%M'),
            )
        else:
            st.info("Live current-alarm feed not found.")

        st.divider()
        st.subheader("📊 Historical Alarm Insights")
        if hist.get('loaded'):
            age_txt = f" ({hist['age_minutes']:.0f} min ago)" if hist['age_minutes'] is not None else ""
            if hist.get('is_stale'):
                st.warning(f"⚠️ Historical rollup generated {hist['generated_at']}{age_txt} — more than 2 "
                           f"days old, the historical alarm scraper pipeline may be stalled.")
            else:
                st.caption(f"Generated {hist['generated_at']}{age_txt}")

            hist_tabs = st.tabs(["🔁 Chronic Offenders", "⏱️ Site Downtime", "📈 Daily Trend", "📂 Category Rollup"])
            with hist_tabs[0]:
                co = hist['chronic_offenders']
                if co.empty:
                    st.info("No chronic-offender data.")
                else:
                    co_search = st.text_input("Filter by Site contains...", "", key="co_search")
                    co_filtered = co[co['Site'].astype(str).str.contains(co_search, case=False, na=False)] \
                        if co_search and 'Site' in co.columns else co
                    sort_col = 'Occurrences' if 'Occurrences' in co_filtered.columns else co_filtered.columns[0]
                    st.dataframe(co_filtered.sort_values(sort_col, ascending=False),
                                 width='stretch', hide_index=True, height=400)
            with hist_tabs[1]:
                sd = hist['site_downtime']
                if sd.empty:
                    st.info("No site downtime data.")
                else:
                    sort_col = 'Total Outage Minutes' if 'Total Outage Minutes' in sd.columns else sd.columns[0]
                    st.dataframe(sd.sort_values(sort_col, ascending=False),
                                 width='stretch', hide_index=True, height=400)
            with hist_tabs[2]:
                dt = hist['daily_trend']
                if dt.empty:
                    st.info("No daily trend data.")
                else:
                    st.dataframe(dt, width='stretch', hide_index=True)
            with hist_tabs[3]:
                cr = hist['category_rollup']
                if cr.empty:
                    st.info("No category rollup data.")
                else:
                    st.dataframe(cr, width='stretch', hide_index=True)
        else:
            st.info("Historical alarm insights snapshot not found.")

        st.divider()
        st.subheader("📅 Daily NOC Alarm Analysis")
        st.caption("Pick one day or a date range to see which sites went down, total summed downtime "
                   "hours, whether each outage has cleared, and each down site's NetEco (power) and NCE "
                   "(transmission) alarm evidence merged into one row. Computed directly from this "
                   "project's own raw MAE/NetEco/NCE exports plus the live current-alarm feed, as far "
                   "back as their retention window covers.")

        noc_dates = cached_noc_alarm_dates()
        yesterday = (pd.Timestamp.now().normalize() - pd.Timedelta(days=1)).date()
        today_d = pd.Timestamp.now().date()
        min_day = pd.Timestamp(noc_dates[0]).date() if noc_dates else yesterday
        rc1, rc2 = st.columns([2, 1])
        noc_range = rc1.date_input(
            "Date range (pick one day, or a start and end day)", value=(yesterday, yesterday),
            min_value=min_day, max_value=today_d, key="daily_noc_alarm_range",
            help=f"Raw exports currently reach back to {min_day}.")
        exclude_long_term = rc2.toggle(
            "Exclude Long-Term Outages from KPIs", value=True, key="daily_noc_exclude_lt",
            help=f"Sites with an outage lasting {noc_alarms.LONG_TERM_OUTAGE_DAYS}+ days (e.g. ECV001, "
                 f"down since 2026-05) are left out of the tiles and daily summary, but always stay "
                 f"in the tables and the Excel export, marked 'Long-Term Outage'.")
        if not isinstance(noc_range, (tuple, list)):
            noc_range = (noc_range,)
        if len(noc_range) == 0:
            noc_range = (yesterday,)
        range_start = noc_range[0]
        range_end = noc_range[1] if len(noc_range) > 1 else noc_range[0]

        with st.spinner("Loading alarm history for the selected days..."):
            noc = cached_daily_noc_alarm_range(range_start.isoformat(), range_end.isoformat())
        noc_sites, noc_events, noc_dq = noc['sites'], noc['events'], noc['data_quality']

        if noc_dq.empty:
            st.info(f"No raw historical alarm export covers {range_start} – {range_end} — either the NOC "
                    f"alarm feed isn't available on this machine, or those dates have aged out of the "
                    f"retention window.")
        else:
            partial_days = noc_dq[noc_dq['Data Coverage'] != 'Full']
            if partial_days.empty:
                st.caption("✅ Full data coverage for every selected day (MAE, NetEco and NCE).")
            else:
                st.warning(f"⚠️ {len(partial_days)} of {len(noc_dq)} selected day(s) have partial data "
                           f"coverage ({', '.join(partial_days['Date'])}) — counts from the listed sources "
                           f"are undercounted on those days, not zero. See the Data Quality tab.")

            with st.expander("🔎 Filters", expanded=False):
                f1, f2, f3 = st.columns(3)
                site_text = f1.text_input("Site Code (comma-separated, partial match)", "",
                                          key="daily_noc_f_site", placeholder="e.g. BGZ, COAST073")
                f_status = f2.multiselect("Alarm Status", [noc_alarms.STATUS_CLEARED,
                                                           noc_alarms.STATUS_NOT_CLEARED,
                                                           noc_alarms.STATUS_UNKNOWN],
                                          key="daily_noc_f_status")
                f_class = f3.multiselect("Outage Class", [noc_alarms.CLASS_NORMAL, noc_alarms.CLASS_LONG_TERM],
                                         key="daily_noc_f_class")
                f4, f5, f6 = st.columns(3)
                f_power = f4.multiselect("Power Reason", list(noc_alarms.POWER_REASON_FILTERS),
                                         key="daily_noc_f_power")
                f_trans = f5.multiselect("Transmission Reason", list(noc_alarms.TRANSMISSION_FILTERS),
                                         key="daily_noc_f_trans")
                f_cov = f6.multiselect("Data Coverage", ["Full", "Partial"], key="daily_noc_f_cov")
                f_hours = st.slider("Down Hours (per site, per day)", 0.0, 24.0, (0.0, 24.0), step=0.25,
                                    key="daily_noc_f_hours")

            site_codes = [c for c in site_text.split(',') if c.strip()]
            hours_filtered = f_hours != (0.0, 24.0)
            active_filters = {
                'Site Code': ', '.join(c.strip() for c in site_codes), 'Alarm Status': ', '.join(f_status),
                'Outage Class': ', '.join(f_class), 'Power Reason': ', '.join(f_power),
                'Transmission Reason': ', '.join(f_trans), 'Data Coverage': ', '.join(f_cov),
                'Down Hours': f"{f_hours[0]:g}–{f_hours[1]:g}" if hours_filtered else '',
            }
            active_filters = {k: v for k, v in active_filters.items() if v}
            sites_view = noc_alarms.apply_noc_alarm_filters(
                noc_sites, noc_dq, site_codes=site_codes, statuses=f_status, outage_classes=f_class,
                power_reasons=f_power, transmission=f_trans,
                min_hours=f_hours[0] if hours_filtered else None,
                max_hours=f_hours[1] if hours_filtered else None, coverage=f_cov)
            if not noc_events.empty and not sites_view.empty:
                kept = pd.MultiIndex.from_frame(sites_view[['Date', 'Site Name']])
                events_view = noc_events[pd.MultiIndex.from_frame(noc_events[['Date', 'Site Name']]).isin(kept)]
            else:
                events_view = noc_events.iloc[0:0]
            kpi = noc_alarms.noc_alarm_kpis(sites_view, exclude_long_term=exclude_long_term)

            n_days = len(noc_dq)
            scope_bits = [f"{n_days} day(s): {range_start}" + (f" → {range_end}" if n_days > 1 else "")]
            scope_bits.append(f"🔎 Filtered results ({len(active_filters)} filter(s) active)" if active_filters
                              else "All sites (unfiltered totals)")
            scope_bits.append(f"Long-Term Outages {'excluded from' if exclude_long_term else 'included in'} KPIs")
            st.caption(" · ".join(scope_bits) + (
                " · Site counts are distinct sites across the range." if n_days > 1 else ""))

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Count of NE Is Disconnected", f"{kpi['ne_disconnected']:,}")
            k2.metric("Total Downtime (Hours)", f"{kpi['total_down_hours']:,.1f}")
            k3.metric("Count of Mains Failure", f"{kpi['mains_failure']:,}")
            k4.metric("Count of NCE Transmission Alarm Sites", f"{kpi['nce_transmission']:,}")
            k5, k6, k7, k8 = st.columns(4)
            k5.metric("Cleared", f"{kpi['cleared']:,}", help="Every outage event has a clearance timestamp.")
            k6.metric("Not Cleared", f"{kpi['not_cleared']:,}",
                      help="Still active in the latest live current-alarm snapshot (fresh, under "
                           f"{noc_alarms.LIVE_STALE_MINUTES} min old).")
            k7.metric("Unknown", f"{kpi['unknown']:,}",
                      help="No clearance timestamp, and the only live snapshot showing it active is stale - "
                           "can't tell whether the site is still down.")
            k8.metric("Long-Term Outage Sites", f"{kpi['long_term']:,}",
                      help=f"Sites with an outage lasting {noc_alarms.LONG_TERM_OUTAGE_DAYS}+ days. Counted here "
                           f"even when excluded from the other KPIs.")

            summary_view = noc_alarms.summarize_noc_alarm_days(sites_view, noc_dq, exclude_long_term)
            active_view = noc_alarms.active_outages(sites_view)

            tab_names = ["🏢 Site Details", "🔁 Outage Events", "📆 Daily Summary", "🚨 Active Outages",
                         "🧪 Data Quality"]
            single_day = range_start == range_end
            if single_day:
                tab_names += ["🔋 NetEco Alarms (this day)", "📡 NCE Alarms (this day)"]
            noc_tabs = st.tabs(tab_names)
            with noc_tabs[0]:
                st.caption("One row per site per day. Status: Cleared / Not Cleared / Unknown; "
                           "Outage Events = how many separate outages the site had that day "
                           "(each listed in the Outage Events tab).")
                if sites_view.empty:
                    st.success("No down sites match this selection.")
                else:
                    st.dataframe(_style_alarm_status(sites_view), width='stretch', hide_index=True, height=420)
            with noc_tabs[1]:
                st.caption("Every individual NE Is Disconnected event - a site that went down three times in "
                           "a day has three rows. Hours This Day is the part of the outage inside that day.")
                if events_view.empty:
                    st.info("No outage events match this selection.")
                else:
                    st.dataframe(_style_alarm_status(events_view), width='stretch', hide_index=True, height=420)
            with noc_tabs[2]:
                st.caption("The KPI tiles, one row per day" + (" (filtered)" if active_filters else "") +
                           (", Long-Term Outages excluded from the counts" if exclude_long_term else "") + ".")
                st.dataframe(summary_view, width='stretch', hide_index=True)
            with noc_tabs[3]:
                st.caption("Sites whose latest outage in the selected range is still Not Cleared or Unknown, "
                           "longest-down first.")
                if active_view.empty:
                    st.success("No active outages in this selection.")
                else:
                    st.dataframe(_style_alarm_status(active_view), width='stretch', hide_index=True)
            with noc_tabs[4]:
                st.caption("Which raw export each day was computed from, and whether it spans the whole day.")
                st.dataframe(noc_dq, width='stretch', hide_index=True)
            if single_day:
                day_report = cached_daily_noc_alarm_report(range_start.isoformat())
                with noc_tabs[5]:
                    ne = day_report['neteco_alarms']
                    if ne.empty:
                        st.info("No NetEco alarm data for this date.")
                    else:
                        st.dataframe(ne, width='stretch', hide_index=True, height=350)
                with noc_tabs[6]:
                    nc = day_report['nce_alarms']
                    if nc.empty:
                        st.info("No NCE alarm data for this date.")
                    else:
                        st.dataframe(nc, width='stretch', hide_index=True, height=350)

            # Export reflects exactly what's on screen (range + filters +
            # long-term toggle); bytes are tied to that selection so a stale
            # file can't be downloaded after the filters change.
            export_sig = (str(range_start), str(range_end), exclude_long_term, tuple(sorted(active_filters.items())))
            ex1, ex2 = st.columns(2)
            if ex1.button("📊 Build Excel workbook", key="daily_noc_xlsx_btn"):
                info_rows = [
                    ('Report', 'Daily NOC Alarm Analysis'),
                    ('Date Range', f"{range_start} to {range_end}"),
                    ('Generated', pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')),
                    ('Long-Term Outages', f"{'Excluded from' if exclude_long_term else 'Included in'} "
                                          f"Daily Summary KPIs (threshold {noc_alarms.LONG_TERM_OUTAGE_DAYS} days); "
                                          f"always listed in detail sheets"),
                    ('Filters', '; '.join(f"{k}: {v}" for k, v in active_filters.items()) or 'None (all sites)'),
                ]
                with st.spinner("Building Excel workbook..."):
                    st.session_state["daily_noc_xlsx"] = (export_sig, noc_alarms.build_noc_alarm_workbook(
                        [('Daily Summary', summary_view), ('Site Details', sites_view),
                         ('Outage Events', events_view), ('Active Outages', active_view),
                         ('Data Quality', noc_dq)], info_rows))
            built = st.session_state.get("daily_noc_xlsx")
            if built and built[0] == export_sig:
                ex1.download_button(
                    "⬇️ Download Excel", data=built[1],
                    file_name=f"Daily_NOC_Alarm_Analysis_{range_start}_to_{range_end}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="daily_noc_xlsx_dl")
            with ex2:
                render_word_export_button(
                    "Daily NOC Alarm Analysis", [("Down Sites", sites_view)],
                    key_prefix="daily_noc_alarm", subtitle=f"{range_start} to {range_end}",
                )

# ============================================================
# 📱 CEM — SmartCare Customer Experience Management: app traffic/TCP
# quality + weekly Device Penetration mix. The CEM half now runs from this
# project's own scrapers/smartcare_cem_scraper.py + reports/
# run_smartcare_analysis_task.py (weekly, on-demand - see run_smartcare_
# pipeline.bat); Device Penetration still reads a sibling suite's weekly export -
# see backend/smartcare_cem_processor.py and backend/device_penetration_processor.py
# ============================================================
elif section == "📱 CEM":
    st.caption("Network-wide subscriber experience data from the SmartCare portal (DPI probe traffic mix, "
               "TCP-level connection quality, and device model mix) - a different data source than the "
               "counter-based KPIs elsewhere in this dashboard, refreshed weekly. No per-site breakdown is "
               "available.")

    sec_tabs = st.tabs(["📶 App Traffic & Quality", "📱 Device Penetration"])

    with sec_tabs[0]:
        cem_overview = cached_cem_overview()
        if not cem_overview.get('loaded'):
            st.info("No SmartCare CEM data found. Run run_smartcare_pipeline.bat (or wait for its weekly "
                    "scheduled run) to produce Comprehensive_Analysis_Historical.xlsx.")
        else:
            if cem_overview.get('is_stale'):
                st.warning(f"⚠️ SmartCare CEM data is {cem_overview['age_days']:.1f} days old — this refreshes "
                           f"weekly upstream, so a gap this large means a scheduled run was likely missed.")
            else:
                st.caption(f"Updated {cem_overview['age_days']:.1f} day(s) ago")

            metrics_df = cem_overview['metrics']
            top100 = cem_overview['top100_per_day']
            top10_week = cem_overview['top10_per_week']

            if not metrics_df.empty:
                latest = metrics_df.iloc[-1]
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("TCP Connection Success Rate", f"{latest['tcp_connection_success_rate']:.3f}%"
                          if 'tcp_connection_success_rate' in latest else "N/A")
                m2.metric("Success Rate (incl. RST)", f"{latest['tcp_connection_success_rate_included_rst']:.3f}%"
                          if 'tcp_connection_success_rate_included_rst' in latest else "N/A")
                m3.metric("DL TCP Retransmission Rate", f"{latest['downlink_tcp_retransmission_rate']:.2f}%"
                          if 'downlink_tcp_retransmission_rate' in latest else "N/A")
                m4.metric("Avg TCP Packet Loss Rate", f"{latest['average_tcp_packet_loss_rate']:.2f}%"
                          if 'average_tcp_packet_loss_rate' in latest else "N/A")
                m5.metric("Total Traffic (GB)", f"{latest['total_traffic_gb']:,.0f}"
                          if 'total_traffic_gb' in latest else "N/A")
                st.caption(f"As of {latest['date']} — {len(metrics_df)} day(s) of history available" if 'date' in latest else "")

                qc1, qc2 = st.columns(2)
                with qc1:
                    st.subheader("📈 Connection Success Rate Trend")
                    st.caption("Zoomed to the observed range - a flat 0-100% axis would hide real day-to-day movement this close to 100%.")
                    if 'tcp_connection_success_rate' in metrics_df.columns:
                        succ_fig = go.Figure()
                        succ_fig.add_trace(go.Scatter(x=metrics_df['date'], y=metrics_df['tcp_connection_success_rate'],
                                                       mode='lines+markers', name='Success Rate', line=dict(color='#2ca02c')))
                        if 'tcp_connection_success_rate_included_rst' in metrics_df.columns:
                            succ_fig.add_trace(go.Scatter(x=metrics_df['date'], y=metrics_df['tcp_connection_success_rate_included_rst'],
                                                           mode='lines', name='Incl. RST', line=dict(color='#98df8a', dash='dot')))
                        lo = metrics_df['tcp_connection_success_rate'].min()
                        hi = metrics_df['tcp_connection_success_rate'].max()
                        pad = max((hi - lo) * 0.15, 0.01)
                        succ_fig.update_layout(height=320, margin=dict(l=30, r=20, t=10, b=30), hovermode='x unified',
                                                yaxis=dict(range=[lo - pad, hi + pad], ticksuffix='%'),
                                                legend=dict(orientation='h', y=-0.3))
                        st.plotly_chart(succ_fig, width='stretch', key='cem_success_trend')
                    else:
                        st.info("No connection success rate data available.")

                with qc2:
                    st.subheader("📉 Retransmission & Packet Loss Trend")
                    st.caption("Same 0-4%-ish scale, so these three are comparable on one chart.")
                    loss_cols = {
                        'downlink_tcp_retransmission_rate': 'DL Retransmission Rate',
                        'average_tcp_packet_loss_rate': 'Avg Packet Loss Rate',
                        'downlink_tcp_packet_loss_rate': 'DL Packet Loss Rate',
                    }
                    present = {c: label for c, label in loss_cols.items() if c in metrics_df.columns}
                    if present:
                        loss_fig = go.Figure()
                        for col, label in present.items():
                            loss_fig.add_trace(go.Scatter(x=metrics_df['date'], y=metrics_df[col], mode='lines+markers', name=label))
                        loss_fig.update_layout(height=320, margin=dict(l=30, r=20, t=10, b=30), hovermode='x unified',
                                                yaxis=dict(ticksuffix='%'), legend=dict(orientation='h', y=-0.3))
                        st.plotly_chart(loss_fig, width='stretch', key='cem_loss_trend')
                    else:
                        st.info("No retransmission/packet loss data available.")

                st.subheader("📦 Total Traffic Volume Trend")
                if 'total_traffic_gb' in metrics_df.columns:
                    vol_fig = go.Figure(go.Scatter(x=metrics_df['date'], y=metrics_df['total_traffic_gb'],
                                                    mode='lines', fill='tozeroy', line=dict(color='#1f77b4')))
                    vol_fig.update_layout(height=280, margin=dict(l=30, r=20, t=10, b=30),
                                           yaxis_title='Traffic (GB)', hovermode='x unified')
                    st.plotly_chart(vol_fig, width='stretch', key='cem_volume_trend')
                else:
                    st.info("No traffic volume data available.")
            else:
                st.info("No daily metrics data available.")

            st.subheader("📶 Top Applications by Traffic (Daily)")
            if not top100.empty and 'date' in top100.columns:
                cem_dates = sorted(top100['date'].astype(str).unique(), reverse=True)
                cem_date_choice = st.selectbox("Date", cem_dates, key="cem_app_date")
                top_n = smartcare_cem.top_apps_for_date(top100, cem_date_choice, n=15)
                if top_n.empty:
                    st.info("No application traffic data for this date.")
                else:
                    day_total = top100.loc[top100['date'].astype(str) == cem_date_choice, 'total_traffic_gb'].sum()
                    top_n = top_n.assign(**{'% of Day Total': (top_n['total_traffic_gb'] / day_total * 100).round(1)}) if day_total else top_n
                    bar_fig = go.Figure(go.Bar(x=top_n['total_traffic_gb'], y=top_n['application'], orientation='h',
                                                text=top_n.get('% of Day Total'),
                                                texttemplate='%{text}%' if '% of Day Total' in top_n.columns else None))
                    bar_fig.update_layout(height=420, margin=dict(l=120, r=20, t=20, b=30),
                                           yaxis=dict(autorange='reversed'), xaxis_title='Traffic (GB)')
                    st.plotly_chart(bar_fig, width='stretch', key='cem_top_apps_chart')
                    st.dataframe(top_n, width='stretch', hide_index=True)
            else:
                st.info("No application traffic data available.")

            st.subheader("🗓️ Top Applications Trend (Weekly)")
            if not top10_week.empty and 'week' in top10_week.columns:
                latest_week = top10_week['week'].max()
                top_apps_latest = (top10_week[top10_week['week'] == latest_week]
                                    .nlargest(6, 'total_traffic_gb')['application'].tolist())
                week_fig = go.Figure()
                for app in top_apps_latest:
                    app_trend = top10_week[top10_week['application'] == app].sort_values('week')
                    week_fig.add_trace(go.Scatter(x=app_trend['week'], y=app_trend['total_traffic_gb'],
                                                   mode='lines+markers', name=app))
                week_fig.update_layout(height=350, margin=dict(l=30, r=20, t=10, b=30), hovermode='x unified',
                                        yaxis_title='Traffic (GB)', legend=dict(orientation='h', y=-0.3),
                                        xaxis_title='Week starting')
                st.plotly_chart(week_fig, width='stretch', key='cem_weekly_apps_trend')
                st.caption(f"Top 6 applications from the week of {latest_week}, tracked across all available weeks.")
            else:
                st.info("No weekly application traffic data available.")

    with sec_tabs[1]:
        dp_overview = cached_device_penetration_overview()
        if not dp_overview.get('loaded'):
            st.info("No Device Penetration data found. This tab reads the output of a separate NOC "
                    "automation suite's weekly Device Penetration export when it's installed and running "
                    "on this machine.")
        else:
            if dp_overview.get('is_stale'):
                st.warning(f"⚠️ Device Penetration data is {dp_overview['age_days']:.1f} days old — this "
                           f"refreshes weekly upstream, so a gap this large means a scheduled run was likely missed.")
            else:
                st.caption(f"Updated {dp_overview['age_days']:.1f} day(s) ago")

            dp_view = st.radio("View", ["📸 Latest Snapshot", "📅 Weekly Report"], horizontal=True,
                               key="dp_view", label_visibility="collapsed")
            if dp_view == "📅 Weekly Report":
                render_dp_weekly_report(dp_overview)
            else:
                snap = dp_overview['latest_snapshot']
                st.caption(f"Snapshot day: {str(dp_overview['latest_time'])[:10]} (latest day in the archive)")

                users_col = 'Number of Users(count)'
                if 'Broadband Capable' in snap.columns and users_col in snap.columns:
                    total_users = snap[users_col].sum()
                    broadband_users = snap.loc[snap['Broadband Capable'], users_col].sum()
                    nr_users = (snap.loc[snap['Device Technology'].astype(str).str.contains('NR', na=False), users_col].sum()
                                if 'Device Technology' in snap.columns else 0)
                    m0, m1, m2, m3, m4 = st.columns(5)
                    m0.metric("Distinct Device Models", f"{snap['Device Model'].nunique():,}",
                              help="How many different device models were seen that day (Samsung A06, "
                                   "iPhone 13 Pro Max, ...) - one table row per model.")
                    m1.metric("Total Devices Seen", f"{total_users:,.0f}",
                              help="Sum of 'Number of Users' across all models - the number of devices "
                                   "(subscribers) active on the network that day.")
                    m2.metric("LTE/NR-Capable", f"{100 * broadband_users / total_users:.1f}%" if total_users else "N/A")
                    m3.metric("5G (NR)-Capable", f"{100 * nr_users / total_users:.1f}%" if total_users else "N/A")
                    m4.metric("Legacy-Only (2G/3G)", f"{100 * (1 - broadband_users / total_users):.1f}%" if total_users else "N/A")
                    st.caption("A high legacy-only share is context for capacity planning: even an idle new LTE "
                               "band gets little uptake at a site until subscriber devices there catch up.")

                # IoT/POS modules (e.g. QUECTEL EC200U - one model, ~6% of all devices)
                # and routers swamp the brand/model rankings, so by default those
                # charts show subscriber handsets only.
                handsets_only = st.toggle("Handsets only (SmartPhone / FeaturePhone / Tablet) for Brand, Model & OS charts",
                                          value=True, key="dp_handsets_only")
                rank_snap = snap
                if handsets_only and 'Device Type' in snap.columns:
                    rank_snap = snap[snap['Device Type'].isin(['SmartPhone', 'FeaturePhone', 'Tablet'])]
                rank_total = rank_snap[users_col].sum() if users_col in rank_snap.columns else 0

                c1, c2 = st.columns(2)
                if 'Device Type' in snap.columns and users_col in snap.columns:
                    with c1:
                        st.subheader("📱 By Device Type")
                        type_totals = snap.groupby('Device Type')[users_col].sum().sort_values(ascending=False)
                        fig_type = go.Figure(go.Bar(x=type_totals.values, y=type_totals.index, orientation='h'))
                        fig_type.update_layout(height=350, margin=dict(l=120, r=20, t=20, b=30),
                                                yaxis=dict(autorange='reversed'), xaxis_title='Users')
                        st.plotly_chart(fig_type, width='stretch', key='dp_type_chart')
                if 'Device Brand' in rank_snap.columns and users_col in rank_snap.columns:
                    with c2:
                        st.subheader("🏷️ Top Brands")
                        brand_totals = rank_snap.groupby('Device Brand')[users_col].sum().sort_values(ascending=False).head(10)
                        brand_share = (100 * brand_totals / rank_total).round(1) if rank_total else None
                        fig_brand = go.Figure(go.Bar(x=brand_totals.values, y=brand_totals.index, orientation='h',
                                                     text=brand_share, texttemplate='%{text}%' if rank_total else None))
                        fig_brand.update_layout(height=350, margin=dict(l=120, r=20, t=20, b=30),
                                                 yaxis=dict(autorange='reversed'), xaxis_title='Users')
                        st.plotly_chart(fig_brand, width='stretch', key='dp_brand_chart')

                if 'Device Model' in rank_snap.columns and users_col in rank_snap.columns:
                    st.subheader("📲 Top 15 Device Models")
                    top_models = rank_snap.nlargest(15, users_col)
                    model_share = (100 * top_models[users_col] / rank_total).round(1) if rank_total else None
                    fig_model = go.Figure(go.Bar(
                        x=top_models[users_col], y=top_models['Device Model'], orientation='h',
                        text=model_share, texttemplate='%{text}%' if rank_total else None,
                        customdata=top_models[['Device Brand', 'Device Technology']].fillna('') if
                        {'Device Brand', 'Device Technology'} <= set(top_models.columns) else None,
                        hovertemplate='%{y}<br>Brand: %{customdata[0]}<br>Tech: %{customdata[1]}'
                                      '<br>Users: %{x:,}<extra></extra>',
                    ))
                    fig_model.update_layout(height=480, margin=dict(l=200, r=20, t=20, b=30),
                                             yaxis=dict(autorange='reversed'), xaxis_title='Users')
                    st.plotly_chart(fig_model, width='stretch', key='dp_model_chart')

                c3, c4 = st.columns(2)
                if 'Device OS' in rank_snap.columns and users_col in rank_snap.columns:
                    with c3:
                        st.subheader("🧩 Device OS")
                        os_totals = rank_snap.groupby('Device OS')[users_col].sum().sort_values(ascending=False)
                        # Fold the long tail (Symbian, Windows, BlackBerry...) into "Other" so the donut stays legible
                        if len(os_totals) > 5:
                            os_totals = pd.concat([os_totals.head(4), pd.Series({'Other': os_totals.iloc[4:].sum()})])
                        fig_os = go.Figure(go.Pie(labels=os_totals.index, values=os_totals.values, hole=0.5,
                                                  sort=False, textinfo='label+percent'))
                        fig_os.update_layout(height=350, margin=dict(l=20, r=20, t=20, b=20), showlegend=False)
                        st.plotly_chart(fig_os, width='stretch', key='dp_os_chart')
                if 'Device Technology' in snap.columns and users_col in snap.columns:
                    with c4:
                        st.subheader("📡 Highest Supported RAT")
                        tech_bucket = _dp_tech_bucket(snap['Device Technology'])
                        tech_totals = snap.groupby(tech_bucket)[users_col].sum().reindex(DP_TECH_ORDER).dropna()
                        tech_share = (100 * tech_totals / tech_totals.sum()).round(1)
                        fig_tech = go.Figure(go.Bar(x=tech_totals.index, y=tech_totals.values,
                                                    text=tech_share, texttemplate='%{text}%',
                                                    marker_color=[DP_TECH_COLORS[b] for b in tech_totals.index]))
                        fig_tech.update_layout(height=350, margin=dict(l=30, r=20, t=20, b=30), yaxis_title='Users')
                        st.plotly_chart(fig_tech, width='stretch', key='dp_tech_chart')
                        st.caption("All device types, grouped by the newest RAT each supports.")

                hist = dp_overview.get('history')
                if hist is not None and not hist.empty and hist['Time'].nunique() > 1 and users_col in hist.columns:
                    st.subheader("📈 Trends Across Snapshots")
                    t1, t2 = st.columns(2)
                    hist_rank = hist
                    if handsets_only and 'Device Type' in hist.columns:
                        hist_rank = hist[hist['Device Type'].isin(['SmartPhone', 'FeaturePhone', 'Tablet'])]
                    if 'Device Brand' in hist_rank.columns:
                        with t1:
                            brand_by_time = hist_rank.groupby(['Time', 'Device Brand'])[users_col].sum().unstack(fill_value=0)
                            brand_pct = brand_by_time.div(brand_by_time.sum(axis=1), axis=0) * 100
                            top_brands = brand_by_time.loc[brand_by_time.index.max()].nlargest(6).index
                            fig_bt = go.Figure()
                            for b in top_brands:
                                fig_bt.add_trace(go.Scatter(x=brand_pct.index, y=brand_pct[b], mode='lines+markers', name=b))
                            fig_bt.update_layout(height=340, margin=dict(l=30, r=20, t=30, b=30), hovermode='x unified',
                                                  title=dict(text='Top 6 brands — share of devices', font=dict(size=14)),
                                                  yaxis=dict(ticksuffix='%'), legend=dict(orientation='h', y=-0.25))
                            st.plotly_chart(fig_bt, width='stretch', key='dp_brand_trend')
                    if 'Device Technology' in hist.columns:
                        with t2:
                            tech_by_time = (hist.groupby(['Time', _dp_tech_bucket(hist['Device Technology'])])[users_col]
                                            .sum().unstack(fill_value=0))
                            tech_pct = tech_by_time.div(tech_by_time.sum(axis=1), axis=0) * 100
                            fig_tt = go.Figure()
                            for b in [b for b in DP_TECH_ORDER if b in tech_pct.columns and b != 'Unidentified']:
                                fig_tt.add_trace(go.Scatter(x=tech_pct.index, y=tech_pct[b], mode='lines+markers',
                                                            name=b, line=dict(color=DP_TECH_COLORS[b])))
                            fig_tt.update_layout(height=340, margin=dict(l=30, r=20, t=30, b=30), hovermode='x unified',
                                                  title=dict(text='Highest supported RAT — share of devices', font=dict(size=14)),
                                                  yaxis=dict(ticksuffix='%'), legend=dict(orientation='h', y=-0.25))
                            st.plotly_chart(fig_tt, width='stretch', key='dp_tech_trend')
                    st.caption("Snapshot dates come from the upstream export; gaps (e.g. Sep 10–27) mean no export ran.")

                dp_search = st.text_input("Filter by Device Model contains...", "", key="dp_search")
                dp_filtered = snap
                if dp_search and 'Device Model' in dp_filtered.columns:
                    dp_filtered = dp_filtered[dp_filtered['Device Model'].astype(str).str.contains(dp_search, case=False, na=False)]
                st.dataframe(
                    dp_filtered.sort_values('Number of Users(count)', ascending=False) if 'Number of Users(count)' in dp_filtered.columns else dp_filtered,
                    width='stretch', hide_index=True, height=400,
                )

# ============================================================
# 📞 USER COMPLAINT — coordinate-based coverage complaint analysis
# (backend/complaint_analyzer.py: EPT-based serving-cell matching, reusing
# the same KPI-threshold/alarm machinery as the rest of this dashboard)
# ============================================================
elif section == "📞 User Complaint":
    st.caption("Match a complainant's coordinates to their likely serving cell(s) - or every cell in an "
               "area - then check those cells' KPIs and alarm history against the same thresholds the rest "
               "of this dashboard uses. Serving-cell matching from coordinates alone is approximate (no "
               "terrain, indoor/outdoor, or live signal data), so treat ranked candidates as an "
               "investigation starting point, not a guaranteed answer.")

    complaint_mode = st.radio("Complaint type", ["📍 Specific user", "⭕ Whole area"],
                               horizontal=True, key="cpl_mode")

    with st.form("cpl_form"):
        cc1, cc2 = st.columns(2)
        with cc1:
            cpl_name = st.text_input("Complainant name", key="cpl_name")
        with cc2:
            cpl_phone = st.text_input("Phone number", key="cpl_phone")
        cc3, cc4 = st.columns(2)
        with cc3:
            cpl_date = st.date_input("Complaint date", value=pd.Timestamp.now().normalize() - pd.Timedelta(days=1),
                                      key="cpl_date")
        with cc4:
            cpl_reason = st.selectbox("Complaint reason", ["No Signal / No Coverage", "Call Drop",
                                                             "Weak Signal", "Slow Data / Poor Throughput",
                                                             "Poor Voice Quality", "Other"], key="cpl_reason")

        paste = st.text_input("Paste coordinates as \"lat, lon\" (from Google Maps, optional shortcut)",
                               key="cpl_paste", placeholder="e.g. 31.86520, 23.92030")
        paste_lat, paste_lon = None, None
        if paste and ',' in paste:
            try:
                paste_lat, paste_lon = (float(x.strip()) for x in paste.split(',')[:2])
            except ValueError:
                st.warning("Couldn't parse that as \"lat, lon\" - using the fields below instead.")

        loc1, loc2, loc3 = st.columns(3)
        with loc1:
            cpl_lat = st.number_input("Latitude", value=paste_lat if paste_lat is not None else 32.0,
                                       format="%.6f", key="cpl_lat")
        with loc2:
            cpl_lon = st.number_input("Longitude", value=paste_lon if paste_lon is not None else 20.0,
                                       format="%.6f", key="cpl_lon")
        with loc3:
            if complaint_mode == "⭕ Whole area":
                cpl_radius_m = st.number_input("Search radius (meters)", min_value=50, max_value=10000,
                                                value=600, step=50, key="cpl_radius_m")
            else:
                cpl_wedge = st.number_input("Azimuth match tolerance (± degrees)", min_value=15, max_value=90,
                                             value=complaint_analyzer.DEFAULT_WEDGE_HALF_DEG, key="cpl_wedge")

        cpl_submitted = st.form_submit_button("🔍 Find cells")

    if cpl_submitted:
        ept_cells = cached_ept_cells()
        st.session_state['cpl_complainant'] = {
            'Name': cpl_name, 'Phone': cpl_phone, 'Date': str(cpl_date), 'Reason': cpl_reason,
            'Latitude': cpl_lat, 'Longitude': cpl_lon,
        }
        st.session_state['cpl_active_mode'] = complaint_mode
        st.session_state['cpl_center'] = (cpl_lat, cpl_lon)
        if complaint_mode == "⭕ Whole area":
            radius_km = cpl_radius_m / 1000.0
            found = complaint_analyzer.find_cells_in_area(cpl_lat, cpl_lon, radius_km, ept_cells=ept_cells)
            st.session_state['cpl_radius_km'] = radius_km
            found = found.assign(Include=True) if not found.empty else found
        else:
            found = complaint_analyzer.find_serving_cells(cpl_lat, cpl_lon, ept_cells=ept_cells,
                                                            wedge_half_deg=cpl_wedge)
            st.session_state['cpl_radius_km'] = None
            found = found.assign(Include=found['Rank'] == 1) if not found.empty else found
        st.session_state['cpl_found'] = found
        st.session_state.pop('cpl_analysis', None)

    found = st.session_state.get('cpl_found')
    if found is not None:
        center = st.session_state['cpl_center']
        radius_km = st.session_state.get('cpl_radius_km')
        mode = st.session_state['cpl_active_mode']

        if found.empty:
            st.warning("No active cells matched this location" +
                       (f" within {radius_km*1000:.0f} m." if radius_km else
                        " within any cell's coverage wedge - try widening the azimuth tolerance."))
        else:
            st.subheader("🗺️ Map")
            highlight_idx = found[found.get('Include', False)].index.tolist() if mode == "📍 Specific user" else []
            fig = complaint_analyzer.build_complaint_map(found, center=center, radius_km=radius_km,
                                                          highlight_index=highlight_idx)
            st.plotly_chart(fig, width='stretch')

            st.subheader("📡 Cells found — select which to include in the analysis")
            display_cols = ['Include', 'Technology', 'Cell Name', 'Site Name', 'Distance (km)',
                             'Azimuth', 'Azimuth Diff', 'Radius (km)']
            if mode == "📍 Specific user":
                display_cols.insert(1, 'Rank')
            display_cols = [c for c in display_cols if c in found.columns]
            edited = st.data_editor(
                found[display_cols], width='stretch', hide_index=True, key="cpl_editor",
                disabled=[c for c in display_cols if c != 'Include'],
                column_config={"Include": st.column_config.CheckboxColumn(required=True)},
            )

            if st.button("📊 Analyze selected cells", key="cpl_analyze_btn"):
                selected = found.loc[edited[edited['Include']].index]
                if selected.empty:
                    st.warning("Select at least one cell to analyze.")
                else:
                    with st.spinner("Checking KPIs, alarms, and interference for the selected cells..."):
                        analysis = complaint_analyzer.build_complaint_analysis(
                            rg, selected, str(st.session_state['cpl_complainant']['Date']), lookback_days=7)
                    st.session_state['cpl_analysis'] = analysis

        analysis = st.session_state.get('cpl_analysis')
        if analysis:
            st.divider()
            st.subheader("🧾 Findings")
            for line in analysis['narratives']:
                st.markdown(f"- {line}")
            if not analysis['narratives']:
                st.info("No findings to report for the selected cells.")

            find_tabs = st.tabs(["⚠️ Failing KPIs", "🚨 Alarm History", "📶 Interference (2G)"])
            with find_tabs[0]:
                if analysis['failing_kpis'].empty:
                    st.success("No KPI threshold breaches on the complaint date for the selected cells.")
                else:
                    st.dataframe(analysis['failing_kpis'], width='stretch', hide_index=True)
            with find_tabs[1]:
                if analysis['alarms'].empty:
                    st.success("No NE Is Disconnected events found for the selected sites in the lookback window.")
                else:
                    st.dataframe(analysis['alarms'], width='stretch', hide_index=True)
            with find_tabs[2]:
                if analysis['interference'].empty:
                    st.info("No 2G interference data for the selected cells in this window.")
                else:
                    st.dataframe(analysis['interference'], width='stretch', hide_index=True)

            cpl_file_stub = (f"Complaint_Report_{st.session_state['cpl_complainant']['Name'] or 'unnamed'}_"
                              f"{st.session_state['cpl_complainant']['Date']}")
            dl_cols = st.columns(2)
            with dl_cols[0]:
                report_bytes = complaint_analyzer.generate_complaint_word_report(
                    rg, st.session_state['cpl_complainant'], mode, analysis)
                st.download_button(
                    "📄 Generate complaint report (Word)", data=report_bytes,
                    file_name=f"{cpl_file_stub}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="cpl_word_dl",
                )
            with dl_cols[1]:
                excel_bytes = complaint_analyzer.generate_complaint_excel_report(
                    rg, st.session_state['cpl_complainant'], mode, analysis)
                st.download_button(
                    "📊 Generate complaint report (Excel)", data=excel_bytes,
                    file_name=f"{cpl_file_stub}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="cpl_excel_dl",
                )

# ============================================================
# 🔎 INVESTIGATE — Cell Explorer, Special Reports
# ============================================================
elif section == "🔎 Investigate":
    sec_tabs = st.tabs(["🔎 Cell Explorer", "🔧 Special Reports"])

    with sec_tabs[0]:
        st.caption("Browse raw cell-level KPIs for any technology/date and filter by cell or site name. "
                   "Click a column header in the table to sort, or use the search icon in the table toolbar.")
        c1, c2, c3 = st.columns([1, 1, 2])
        with c1:
            explore_tech = st.selectbox("Technology", list(CELL_SHEETS.keys()),
                                         format_func=lambda t: TECH_LABELS[t])
        with c2:
            explore_date = st.selectbox("Date", all_dates, key="explore_date")
        with c3:
            search = st.text_input("Filter by Cell Name / Site Name contains...", "")

        sheet_name = CELL_SHEETS[explore_tech]
        site_col = SITE_COL_BY_TECH.get(explore_tech)  # GSM='Site Name', UMTS='NodeB Name', LTE='eNodeB Name'
        raw = cached_sheet(sheet_name)
        if raw is None:
            st.error(f"{sheet_name}.csv not found.")
        else:
            name_col = 'Cell Name' if 'Cell Name' in raw.columns else (site_col if site_col in raw.columns else None)
            filtered = raw[raw['Date'].astype(str) == str(explore_date)] if 'Date' in raw.columns else raw
            if search:
                name_cols = [c for c in ['Cell Name', site_col] if c and c in filtered.columns]
                if name_cols:
                    mask = False
                    for c in name_cols:
                        mask = mask | filtered[c].astype(str).str.contains(search, case=False, na=False)
                    filtered = filtered[mask]
            st.caption(f"{len(filtered):,} row(s)")
            st.dataframe(filtered, width='stretch', hide_index=True, height=400)

            if name_col is None:
                st.info(f"This sheet has no Cell Name / {site_col} column to select on.")
            else:
                st.divider()
                has_site_col = bool(site_col) and site_col in filtered.columns and 'Cell Name' in filtered.columns
                if has_site_col:
                    group_mode = st.radio(
                        "Group by", ["Cell", "Site"], horizontal=True, key="cell_explorer_group_mode",
                        help=f"Site groups every cell/sector under the selected {site_col.lower()}(s) together - "
                             "useful for checking a change/operation request that only touched specific sites.",
                    )
                else:
                    group_mode = "Cell"
                pick_col = site_col if group_mode == "Site" else name_col

                options = sorted(filtered[pick_col].dropna().astype(str).unique().tolist())
                picked = st.multiselect(
                    f"Select one or more {pick_col.lower()}s to inspect together (e.g. BGZ001, BGZ002...)",
                    options=options, key="cell_explorer_select",
                )

                if picked:
                    if pick_col == site_col:
                        selected = cached_resolve_group_to_cells(explore_tech, tuple(picked), site_col)
                        st.caption(f"Resolved to {len(selected)} cell(s)/sector(s) under {len(picked)} selected site(s).")
                    else:
                        selected = picked

                if picked and not selected:
                    st.warning("No cells found under the selected site(s) for this technology/date.")
                elif picked:
                    st.subheader(f"📋 Combined KPIs — {len(selected)} cell(s)")
                    combined = filtered[filtered[name_col].isin(selected)]
                    st.dataframe(combined, width='stretch', hide_index=True)

                    if pick_col == site_col:
                        render_site_detail_and_advice(picked, explore_date, key_prefix="ce")

                    sel_tuple = tuple(selected)
                    failing = cached_cell_failing(explore_tech, sel_tuple, explore_date)
                    st.subheader("⚠️ Failing KPIs & Suggested Fixes")
                    if failing is not None and not failing.empty:
                        st.dataframe(failing, width='stretch', hide_index=True)
                    else:
                        st.success("No threshold KPIs are failing for the selected cell(s) on this date.")

                    st.subheader("📈 14-Day Trend (selected cells)")
                    cell_trend = cached_cell_trend(explore_tech, sel_tuple, explore_date)
                    if cell_trend is None or cell_trend.empty:
                        st.info("No historical data available for the selected cell(s).")
                    else:
                        kpi_thresholds = cached_cell_thresholds(explore_tech, sheet_name)
                        kpi_dimensions = cached_cell_dimensions(explore_tech, sheet_name)
                        kpi_cols = [c for c in cell_trend.columns if c not in ('Date', 'Cell')]
                        st.caption("Dashed line = threshold. One line per selected cell; hover to compare. "
                                   "Grouped by KPI dimension - click a section to expand.")
                        render_grouped_cell_trend_charts(
                            cell_trend, kpi_cols, kpi_thresholds, kpi_dimensions,
                            key_prefix=f"cell_trend_{explore_tech}",
                        )

                        st.divider()
                        if st.button("📄 Export Selection as Word", key="ce_word_btn"):
                            with st.spinner("Building Word report (tables + trend charts)..."):
                                ce_label = ', '.join(picked[:3]) + (f" +{len(picked) - 3} more" if len(picked) > 3 else '')
                                ce_start, ce_end = cell_trend['Date'].min(), cell_trend['Date'].max()
                                ce_path = rg.generate_group_word_report(
                                    explore_tech, selected, ce_label, ce_start, ce_end)
                            if ce_path:
                                with open(ce_path, 'rb') as f:
                                    st.session_state['ce_word_bytes'] = f.read()
                                st.session_state['ce_word_filename'] = os.path.basename(ce_path)
                        if 'ce_word_bytes' in st.session_state:
                            st.download_button(
                                "⬇️ Download", data=st.session_state['ce_word_bytes'],
                                file_name=st.session_state.get('ce_word_filename', 'Cell_Explorer_Report.docx'),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="ce_word_dl",
                            )
                else:
                    st.caption("Select cell(s) or site(s) above to see a combined table, failing-KPI summary, "
                               "and 14-day trend charts for them.")

    with sec_tabs[1]:
        st.caption("Build a standalone report for a specific set of sites/cells over any date range - e.g. to "
                   "verify a change/operation request that only touched certain sites.")

        sr_tech = st.selectbox("Technology", list(CELL_SHEETS.keys()),
                                format_func=lambda t: TECH_LABELS[t], key="sr_tech")
        sr_site_col = SITE_COL_BY_TECH.get(sr_tech)
        sr_sheet = CELL_SHEETS[sr_tech]
        sr_raw = cached_sheet(sr_sheet)

        if sr_raw is None:
            st.error(f"{sr_sheet}.csv not found.")
        else:
            sr_has_site_col = bool(sr_site_col) and sr_site_col in sr_raw.columns and 'Cell Name' in sr_raw.columns
            sr_group_mode = st.radio(
                "Group by", ["Cell", "Site"], horizontal=True, key="sr_group_mode",
                index=1 if sr_has_site_col else 0, disabled=not sr_has_site_col,
            )
            sr_pick_col = sr_site_col if (sr_group_mode == "Site" and sr_has_site_col) else 'Cell Name'
            sr_search = st.text_input("Filter by Cell Name / Site Name contains...", "", key="sr_search")
            sr_options = sorted(sr_raw[sr_pick_col].dropna().astype(str).unique().tolist()) \
                if sr_pick_col in sr_raw.columns else []
            if sr_search:
                sr_options = [o for o in sr_options if sr_search.lower() in o.lower()]
            sr_picked = st.multiselect(
                f"Select one or more {sr_pick_col.lower()}s (e.g. BGZ001, BGZ002...)",
                options=sr_options, key="sr_select",
            )

            asc_dates = sorted(all_dates)
            c1, c2, c3 = st.columns([1, 1, 2])
            with c1:
                sr_start = st.selectbox("Start date", asc_dates, index=0, key="sr_start")
            with c2:
                sr_end = st.selectbox("End date", asc_dates, index=len(asc_dates) - 1, key="sr_end")
            with c3:
                sr_label = st.text_input("Report label (e.g. \"BGZ070 antenna swap - CR#1234\")", key="sr_label")

            if not sr_picked:
                st.caption("Select site(s)/cell(s) and a date range above to preview KPIs, failing checks, "
                           "trend charts, and export a report.")
            elif sr_start > sr_end:
                st.error("Start date is after end date.")
            else:
                if sr_pick_col == sr_site_col:
                    sr_cells = cached_resolve_group_to_cells(sr_tech, tuple(sr_picked), sr_site_col)
                    st.caption(f"Resolved to {len(sr_cells)} cell(s)/sector(s) under {len(sr_picked)} selected site(s).")
                else:
                    sr_cells = sr_picked

                if not sr_cells:
                    st.warning("No cells found for the selected group.")
                else:
                    sr_trend = cached_cell_trend_range(sr_tech, tuple(sr_cells), sr_start, sr_end)
                    if sr_trend is None or sr_trend.empty:
                        st.info("No data available for this group/date range.")
                    else:
                        latest_date = sr_trend['Date'].max()
                        st.subheader(f"📋 Combined KPIs — {latest_date}")
                        st.dataframe(sr_trend[sr_trend['Date'] == latest_date].drop(columns=['Date']),
                                     width='stretch', hide_index=True)

                        if sr_pick_col == sr_site_col:
                            render_site_detail_and_advice(sr_picked, latest_date, key_prefix="sr")

                        st.subheader("⚠️ Failing KPIs & Suggested Fixes")
                        sr_failing = cached_cell_failing(sr_tech, tuple(sr_cells), latest_date)
                        if sr_failing is not None and not sr_failing.empty:
                            st.dataframe(sr_failing, width='stretch', hide_index=True)
                        else:
                            st.success("No threshold KPIs are failing for this group on the latest date.")

                        st.subheader(f"📈 Trend ({sr_start} to {sr_end})")
                        sr_kpi_thresholds = cached_cell_thresholds(sr_tech, sr_sheet)
                        sr_kpi_dimensions = cached_cell_dimensions(sr_tech, sr_sheet)
                        sr_kpi_cols = [c for c in sr_trend.columns if c not in ('Date', 'Cell')]
                        render_grouped_cell_trend_charts(
                            sr_trend, sr_kpi_cols, sr_kpi_thresholds, sr_kpi_dimensions,
                            key_prefix=f"sr_trend_{sr_tech}",
                        )

                        st.divider()
                        group_label = sr_label.strip() or ', '.join(sr_picked[:3]) + \
                            (f" +{len(sr_picked) - 3} more" if len(sr_picked) > 3 else '')
                        if st.button("📄 Prepare Special Report (.docx)", width='stretch', key="sr_prepare"):
                            with st.spinner("Building special report..."):
                                sr_path = rg.generate_group_word_report(sr_tech, sr_cells, group_label, sr_start, sr_end)
                            if sr_path:
                                with open(sr_path, 'rb') as f:
                                    st.session_state['sr_bytes'] = f.read()
                                st.session_state['sr_filename'] = os.path.basename(sr_path)
                            else:
                                st.error("Could not generate the report - no data for this group/date range.")
                        if 'sr_bytes' in st.session_state:
                            st.download_button(
                                "⬇️ Download Special Report", data=st.session_state['sr_bytes'],
                                file_name=st.session_state.get('sr_filename', 'Special_Report.docx'),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                width='stretch',
                            )

# ============================================================
# 📋 HQ REPORTS — recurring HQ/Tripoli report templates (e.g. NQ Data
# Collection Template), built straight from output/csv/ instead of the
# old manual per-sheet scripts. One sub-tab per report, more added over
# time as they come up.
# ============================================================
elif section == "📋 HQ Reports":
    sec_tabs = st.tabs([
        "📄 NQ Data Collection Template", "🌐 Traffic & Availability (Tripoli HQ)",
        "📶 Monthly Cell Info Report", "📈 PS Traffic per site", "🚦 Traffic per site",
        "📊 Monthly Comprehensive Analysis (Tripoli)",
    ], on_change="rerun", key="hq_tabs")

    if sec_tabs[0].open:
        with sec_tabs[0]:
            st.caption("EAST branch only, built from output/csv/ history. North/South/Middle/West and "
                       "nationwide \"Libyana\" totals aren't produced here - this pipeline only ever "
                       "has FTP/OSS access to the East branch.")

            period_labels = {'Day': 'day', 'Week': 'week', 'Month': 'month', 'Quarter': 'quarter'}
            period_choice = st.radio(
                "External Interference period", list(period_labels.keys()), index=2,
                horizontal=True,
                help="Only affects the External Interference table below. The bad-day "
                     "persistence bar scales with the chosen window (~1-in-5 days bad, "
                     "same duty cycle at every granularity) rather than staying frozen "
                     "at the monthly '>5 days' count.",
            )
            interference_period = period_labels[period_choice]

            nq_sheets = cached_nq_template(interference_period)

            sheet_tabs = st.tabs(list(nq_sheets.keys()))
            for sheet_tab, (sheet_name, df) in zip(sheet_tabs, nq_sheets.items()):
                with sheet_tab:
                    if df is not None and not df.empty:
                        st.dataframe(df, width='stretch', hide_index=True)
                        st.caption(f"{len(df)} row(s)")
                    else:
                        st.info("No data available for this sheet yet.")

            st.divider()
            st.download_button(
                "⬇️ Download NQ Data Collection Template (.xlsx)",
                data=cached_nq_template_bytes(interference_period),
                file_name=f"NQ_Data_Collection_Template_{target_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.caption("Download only for now — upload to SharePoint manually until auto-upload is set up.")

    if sec_tabs[1].open:
        with sec_tabs[1]:
            import calendar as _calendar

            st.caption(
                "EAST branch only. Mirrors Sheet1 of config/Traffic and network availability "
                "Needed from Tripoli HQ.xlsx — Sum of 4G PS Traffic, Sum of Gi Interface Traffic "
                "Volume, Average DL PRB Utilization, Average 2G Network Availability, and Average "
                "DL Throughput per User, one row per calendar month. Already-filled months are "
                "read straight from that file; pick a month below and click Compute to fill in "
                "the newest one from the pipeline's own output/csv/ history — needed once a "
                "month, after that month ends."
            )

            target_dt = pd.to_datetime(target_date)
            default_month = target_dt.month - 1 or 12
            default_year = target_dt.year if target_dt.month > 1 else target_dt.year - 1

            pc1, pc2, pc3 = st.columns([1, 2, 2])
            with pc1:
                sel_year = int(st.number_input("Year", value=default_year, step=1, format="%d"))
            with pc2:
                sel_month = st.selectbox(
                    "Month", list(range(1, 13)), index=default_month - 1,
                    format_func=lambda m: _calendar.month_name[m],
                )
            with pc3:
                st.write("")
                st.write("")
                compute_clicked = st.button(
                    f"🔄 Compute & Save {_calendar.month_name[sel_month]} {sel_year}", width='stretch',
                )

            if compute_clicked:
                computed = compute_hq_traffic_month('output/csv', sel_year, sel_month)
                missing = [f for f in HQ_TRAFFIC_METRICS if computed.get(f) is None]
                if len(missing) == len(HQ_TRAFFIC_METRICS):
                    st.warning(
                        f"No pipeline data found for {_calendar.month_name[sel_month]} {sel_year} "
                        "in output/csv/ — nothing to save."
                    )
                else:
                    save_hq_traffic_month(computed)
                    st.cache_data.clear()
                    if missing:
                        st.warning(f"Saved, but no source data yet for: {', '.join(missing)}.")
                    for src, (have, total) in computed['coverage'].items():
                        if have < total:
                            st.info(
                                f"⚠️ {src}: only {have}/{total} day(s) of the month were in "
                                "output/csv/ — this month's figure is based on a partial month."
                            )
                    st.success(f"Saved {_calendar.month_name[sel_month]} {sel_year} to the Tripoli HQ template.")
                    st.rerun()

            st.divider()
            history_df = cached_hq_traffic_history()
            if not history_df.empty:
                display_df = history_df.copy()
                display_df['Month'] = display_df['Month'].apply(lambda m: _calendar.month_name[int(m)])
                st.dataframe(display_df, width='stretch', hide_index=True)
                st.caption(f"{len(display_df)} month(s) on file.")
            else:
                st.info("No months saved yet.")

            st.divider()
            if os.path.exists(HQ_TRAFFIC_TEMPLATE_FILE):
                with open(HQ_TRAFFIC_TEMPLATE_FILE, 'rb') as f:
                    st.download_button(
                        "⬇️ Download Traffic & Availability Report (.xlsx)", data=f.read(),
                        file_name="Traffic and network availability Needed from Tripoli HQ.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width='stretch',
                    )
            st.caption("Download and send to Tripoli HQ manually.")

    if sec_tabs[2].open:
        with sec_tabs[2]:
            st.caption(
                "Merged 2G/3G/4G cell inventory (site, cell, CI/LAC, band, traffic, lat/long, "
                "azimuth, GCI/eGCI) pulled monthly via FTPS by the sibling cell_info_report.py "
                f"puller. Read directly from {CELL_INFO_OUTPUT_DIR}."
            )

            months = cached_cell_info_months()
            if not months:
                st.info(
                    "No monthly Cell Info output found yet. Run the sibling project's "
                    "cell_info_report.py to pull and build the first month's merged file."
                )
            else:
                sel_ci_month = st.selectbox("Month", months, index=0)
                ci_df, ci_filename, ci_bytes = cached_cell_info_report(sel_ci_month)
                if ci_df is None:
                    st.warning(f"No merged_df_*.xlsx found under {sel_ci_month}/ yet.")
                else:
                    tech_options = sorted(ci_df['Network Type'].dropna().unique().tolist()) \
                        if 'Network Type' in ci_df.columns else []
                    mc1, mc2, mc3, mc4 = st.columns(4)
                    mc1.metric("Total cells", f"{len(ci_df):,}")
                    for col, tech in zip((mc2, mc3, mc4), tech_options[:3]):
                        col.metric(f"{tech} cells", f"{(ci_df['Network Type'] == tech).sum():,}")

                    sel_techs = st.multiselect("Filter by Network Type", tech_options, default=tech_options)
                    filtered_ci = ci_df[ci_df['Network Type'].isin(sel_techs)] if tech_options else ci_df
                    st.dataframe(filtered_ci, width='stretch', hide_index=True)
                    st.caption(f"{len(filtered_ci)} of {len(ci_df)} row(s) shown — source file: {ci_filename}")

                    st.download_button(
                        "⬇️ Download Monthly Cell Info Report (.xlsx)", data=ci_bytes,
                        file_name=ci_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    if sec_tabs[3].open:
        with sec_tabs[3]:
            st.caption(
                "Daily PS (data) traffic per site across 2G/3G/4G, in the same shape "
                "PS Traffic per site v3.py produces (Date, Site, per-tech + Total GB, Region)."
            )
            if st.button("🔄 Refresh PS Traffic now"):
                launch_ps_traffic_refresh()
                st.session_state['ps_traffic_refresh_started'] = True
            if st.session_state.get('ps_traffic_refresh_started'):
                st.info("Refresh started in the background — reload this tab in a few minutes.")

            ps_sheets = cached_ps_traffic_from_pipeline()
            if ps_sheets is None:
                st.info("No PS traffic data available yet — output/csv/Traffic_2G/3G/4G.csv haven't been built.")
            else:
                ps_bytes = cached_ps_traffic_from_pipeline_bytes()
                date_col_config = {"Date": st.column_config.TextColumn("Date")}

                ps_sheet_tabs = st.tabs(list(ps_sheets.keys()))
                for ps_tab, (ps_sheet_name, ps_df) in zip(ps_sheet_tabs, ps_sheets.items()):
                    with ps_tab:
                        if ps_sheet_name == "All_Traffic_Data" and "Site_Name" in ps_df.columns:
                            site_query = st.text_input(
                                "Filter by site name (contains)", key="ps_traffic_site_filter"
                            )
                            shown = (
                                ps_df[ps_df['Site_Name'].str.contains(site_query, case=False, na=False)]
                                if site_query else ps_df
                            )
                        else:
                            shown = ps_df
                        # Date is already a plain "dd-mm-yyyy" string, but Streamlit's
                        # dataframe widget auto-detects date-like text and re-renders
                        # it per the browser's own locale (eg. "1/7/2026") unless
                        # explicitly pinned to TextColumn - forcing our fixed format
                        # to actually reach the screen unchanged.
                        st.dataframe(
                            shown, width='stretch', hide_index=True,
                            column_config=date_col_config if 'Date' in shown.columns else None,
                        )
                        st.caption(f"{len(shown)} row(s)")

                st.divider()
                st.download_button(
                    "⬇️ Download PS Traffic per site Report (.xlsx)", data=ps_bytes,
                    file_name="PS_Traffic_per_site_pipeline.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    if sec_tabs[4].open:
        with sec_tabs[4]:
            st.caption(
                "PS + CS traffic per site (2G/3G/4G), filterable by date range and by Region / FN-HUB node / "
                "hand-picked sites - e.g. \"CS and PS traffic for the KUFRA region, last 14 days.\""
            )

            tp_col1, tp_col2 = st.columns(2)
            with tp_col1:
                tp_start = st.date_input(
                    "Start date", value=pd.Timestamp(target_date) - pd.Timedelta(days=13), key="tp_start_date"
                )
            with tp_col2:
                tp_end = st.date_input("End date", value=pd.Timestamp(target_date), key="tp_end_date")

            tp_scope_mode = st.radio(
                "Scope", ["All Sites", "By Region", "By FN/HUB Node", "Specific Sites"],
                horizontal=True, key="tp_scope_mode",
            )

            tp_sites = []
            tp_scope_label = "All Sites"
            tp_group_sites, tp_group_label, tp_group_key = [], "", ""
            if tp_scope_mode == "All Sites":
                tp_sites = cached_all_traffic_sites()
            elif tp_scope_mode == "By Region":
                tp_region = st.selectbox("Region", cached_traffic_regions(), key="tp_region")
                if tp_region:
                    tp_group_sites = cached_traffic_sites_for_region(tp_region)
                    tp_group_label, tp_group_key = f"Region: {tp_region}", f"region_{tp_region}"
            elif tp_scope_mode == "By FN/HUB Node":
                tp_node = st.selectbox("FN/HUB Node", cached_traffic_fn_hub_nodes(), key="tp_fn_hub_node")
                if tp_node:
                    tp_group_sites = cached_traffic_sites_for_fn_hub(tp_node)
                    tp_group_label, tp_group_key = f"FN/HUB Node: {tp_node}", f"node_{tp_node}"
            else:
                tp_picked = st.multiselect(
                    "Select site(s)", cached_all_traffic_sites(), key="tp_sites_picked"
                )
                tp_sites = tp_picked
                tp_scope_label = f"{len(tp_picked)} selected site(s)" if tp_picked else "No sites selected"

            if tp_group_label:
                # Editable group: pre-filled with the Region/FN-HUB's sites -
                # untick to drop some, pick any other site to add it. Keyed per
                # group so switching node/region re-seeds the list.
                tp_group_default = sorted({base_site_name(s) for s in tp_group_sites})
                tp_sites = st.multiselect(
                    "Sites in scope (remove or add sites)",
                    sorted(set(cached_all_traffic_sites()) | set(tp_group_default)),
                    default=tp_group_default, key=f"tp_group_sites_{tp_group_key}",
                )
                def _site_list(names, limit=3):
                    names = sorted(names)
                    return ", ".join(names[:limit]) + (f" +{len(names) - limit} more" if len(names) > limit else "")
                tp_added = set(tp_sites) - set(tp_group_default)
                tp_removed = set(tp_group_default) - set(tp_sites)
                tp_scope_label = tp_group_label
                tp_changes = ([f"+ {_site_list(tp_added)}"] if tp_added else []) + \
                             ([f"− {_site_list(tp_removed)}"] if tp_removed else [])
                if tp_changes:
                    tp_scope_label += f" ({' / '.join(tp_changes)})"

            if not tp_sites:
                st.info("Select a region, FN/HUB node, or specific site(s) above to see traffic data.")
            elif tp_start > tp_end:
                st.error("Start date must be on or before end date.")
            else:
                tp_start_str = tp_start.strftime('%Y-%m-%d')
                tp_end_str = tp_end.strftime('%Y-%m-%d')
                st.caption(f"Scope: **{tp_scope_label}** ({len(tp_sites)} site(s)) — {tp_start_str} to {tp_end_str}")

                tp_sites_key = tuple(sorted(tp_sites))
                tp_detail = cached_site_traffic_detail(tp_sites_key, tp_start_str, tp_end_str)
                if tp_detail is None or tp_detail.empty:
                    st.warning("No traffic data for the selected scope/date range.")
                else:
                    _, tp_no_data = rg.split_sites_by_traffic_data(tp_sites, tp_detail)
                    if tp_no_data:
                        st.warning(f"{len(tp_no_data)} site(s) in scope have no traffic data in this range "
                                   f"(transmission-only FN/HUB node, or not in the traffic export): "
                                   + ", ".join(tp_no_data))
                    st.subheader("📋 Per-Site Detail")
                    st.dataframe(
                        tp_detail, width='stretch', hide_index=True,
                        column_config={"Date": st.column_config.TextColumn("Date")},
                    )
                    st.caption(f"{len(tp_detail)} row(s)")

                    st.subheader(f"📈 Combined Trend — {tp_scope_label}")
                    st.caption("Sum across every site in the selected scope, one line per metric - a region/FN-HUB "
                               "with many sites still renders as one readable line instead of one per site.")
                    tp_trend = cached_site_traffic_trend(tp_sites_key, tp_start_str, tp_end_str)
                    if tp_trend is None or tp_trend.empty:
                        st.info("No trend data available.")
                    else:
                        trend_metric_cols = [c for c in tp_trend.columns if c != 'Date']
                        tp_chart_cols = st.columns(2)
                        for i, metric in enumerate(trend_metric_cols):
                            with tp_chart_cols[i % 2]:
                                render_trend_chart(tp_trend, metric, None, key=f"tp_trend_{metric}")

                    st.divider()
                    st.subheader("📥 Export")
                    tp_ec1, tp_ec2 = st.columns(2)
                    with tp_ec1:
                        if st.button("📄 Prepare Word Report (.docx)", key="tp_word_btn", width='stretch'):
                            with st.spinner("Building Word report (table + trend charts)..."):
                                tp_word_path = rg.generate_traffic_group_word_report(
                                    tp_scope_label, tp_sites, tp_start_str, tp_end_str)
                            if tp_word_path:
                                with open(tp_word_path, 'rb') as f:
                                    st.session_state['tp_word_bytes'] = f.read()
                                st.session_state['tp_word_filename'] = os.path.basename(tp_word_path)
                        if 'tp_word_bytes' in st.session_state:
                            st.download_button(
                                "⬇️ Download Word Report", data=st.session_state['tp_word_bytes'],
                                file_name=st.session_state.get('tp_word_filename', 'Traffic_per_Site_Report.docx'),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="tp_word_dl", width='stretch',
                            )
                    with tp_ec2:
                        if st.button("📊 Prepare Excel Report (.xlsx)", key="tp_excel_btn", width='stretch'):
                            with st.spinner("Building Excel report (data + charts)..."):
                                st.session_state['tp_excel_bytes'] = rg.generate_traffic_group_excel_report(
                                    tp_scope_label, tp_sites, tp_start_str, tp_end_str)
                        if st.session_state.get('tp_excel_bytes'):
                            st.download_button(
                                "⬇️ Download Excel Report", data=st.session_state['tp_excel_bytes'],
                                file_name=f"Traffic_per_Site_{tp_start_str}_to_{tp_end_str}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="tp_excel_dl", width='stretch',
                            )

    if sec_tabs[5].open:
        with sec_tabs[5]:
            st.caption(
                "SmartCare CEM Comprehensive Analysis for one calendar month, in the 3-sheet format Tripoli HQ "
                "receives (Top100 per day · Top10 Application per week · Metrics). Built automatically on the 2nd "
                "of each month for the previous month; pick any month here to rebuild and download it."
            )
            if not smartcare_cem.is_available():
                st.info(f"CEM history workbook not found: {smartcare_cem.CEM_WORKBOOK_PATH}")
            else:
                cem_mtime = os.path.getmtime(smartcare_cem.CEM_WORKBOOK_PATH)
                months = cached_cem_months(cem_mtime)
                mc1, mc2 = st.columns([1, 2])
                # Default to last month - the one HQ is waiting for - not the
                # current, still-incomplete month.
                prev_month = (pd.Timestamp.now().normalize().replace(day=1)
                              - pd.Timedelta(days=1)).strftime('%Y-%m')
                with mc1:
                    cem_month = st.selectbox(
                        "Month", months, key="cem_monthly_month",
                        index=months.index(prev_month) if prev_month in months else 0,
                        format_func=lambda m: pd.Timestamp(f"{m}-01").strftime("%B %Y"),
                    )
                data, filename, info = cached_cem_monthly(cem_month, cem_mtime)
                with mc2:
                    st.metric("Days with data", f"{len(info['present'])} / {info['days_in_month']}")
                if info['complete']:
                    st.success(f"✅ {pd.Timestamp(f'{cem_month}-01'):%B %Y} is complete - ready to send.")
                else:
                    st.warning(
                        f"⚠️ {len(info['missing'])} day(s) missing: {', '.join(d[5:] for d in info['missing'])}. "
                        "The daily SmartCare run normally fills them within a few days; for older gaps export "
                        "that date range manually in SmartCare (≤ 25 days per export) into SmartCare_Exports."
                    )
                if info['low_rows']:
                    st.caption(f"ℹ️ Smaller-than-usual day(s) (fewer application rows): "
                               f"{', '.join(d[5:] for d in info['low_rows'])} - usually just quiet days.")

                st.download_button(
                    f"⬇️ Download {filename}", data=data, file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="cem_monthly_dl", width='stretch',
                )
                saved = os.path.join(smartcare_cem.MONTHLY_REPORT_DIR, filename)
                if os.path.exists(saved):
                    st.caption(f"Auto-saved copy: `{saved}` "
                               f"({pd.Timestamp.fromtimestamp(os.path.getmtime(saved)):%Y-%m-%d %H:%M})")

                import io as _io
                prev_tabs = st.tabs(["Top100 per day", "Top10 Application per week", "Metrics"])
                for ptab, sheet in zip(prev_tabs, ["Top100 per day", "Top10 Application per week", "Metrics"]):
                    with ptab:
                        pdf = pd.read_excel(_io.BytesIO(data), sheet_name=sheet)
                        st.dataframe(pdf, width='stretch', hide_index=True, height=350)
                        st.caption(f"{len(pdf):,} row(s)")

# ============================================================
# 📧 REPORTS — network summary + copy-paste text + Word/Excel export
# ============================================================
elif section == "📧 Reports":
    sec_tabs = st.tabs(["📧 Report & Export"])

    with sec_tabs[0]:
        st.subheader("📈 Network Summary")
        summary_rows = cached_network_summary(target_date)
        c1, c2 = st.columns(2)
        for i, (label, value) in enumerate(summary_rows):
            (c1 if i % 2 == 0 else c2).metric(label, value)

        st.divider()
        st.subheader("📥 Export")
        st.caption("The Word report includes the site/network summary, every scorecard table, and "
                   "14-day trend charts with threshold lines (same as the Trend tab). "
                   "First export for a given date takes a few seconds to render the charts; cached after that.")
        ec1, ec2 = st.columns(2)
        with ec1:
            if st.button("📄 Prepare Word Report (.docx)", width='stretch'):
                with st.spinner("Building Word report (tables + trend charts)..."):
                    st.session_state['word_bytes'] = cached_word_bytes(target_date, previous_date)
            if 'word_bytes' in st.session_state:
                st.download_button(
                    "⬇️ Download Word Report", data=st.session_state['word_bytes'],
                    file_name=f"Network_Report_{target_date}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    width='stretch',
                )
        with ec2:
            if st.button("📊 Prepare Excel Report (.xlsx)", width='stretch'):
                with st.spinner("Building Excel report..."):
                    st.session_state['excel_bytes'] = cached_excel_bytes(target_date, previous_date)
            if 'excel_bytes' in st.session_state:
                st.download_button(
                    "⬇️ Download Excel Report", data=st.session_state['excel_bytes'],
                    file_name=f"Network_Report_{target_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch',
                )
