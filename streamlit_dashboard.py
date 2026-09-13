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
    ReportGenerator, TECH_LABELS, CELL_SHEETS, SCORECARD_SHEETS, SITE_COL_BY_TECH, autofit_excel_columns,
)
from backend import ept_manager as ept
from backend.special_reports_processor import (
    build_nq_template_report, load_hq_traffic_history, compute_hq_traffic_month,
    save_hq_traffic_month, HQ_TRAFFIC_TEMPLATE_FILE, HQ_TRAFFIC_METRICS,
)
from backend.topology_processor import build_site_topology_csv, find_topology_xlsx
from backend import smartcare_cem_processor as smartcare_cem
from backend import complaint_analyzer
from project_config import env_path_str

# External FTPS-pulled reports (scripts/cell_info_report.py and the sibling
# "PS Traffic per site" puller live outside this repo, in the sister
# NOC Automation Suite project) - the dashboard only reads their finished
# output files, same DATA_ROOT convention as every scraper in this project.
CELL_INFO_OUTPUT_DIR = env_path_str(
    "CELL_INFO_OUTPUT_DIR",
    os.path.join(env_path_str("DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data"), "Output", "Cell_Info"),
)
PS_TRAFFIC_OUTPUT_DIR = env_path_str(
    "PS_TRAFFIC_OUTPUT_DIR",
    os.path.join(env_path_str("DATA_ROOT", r"C:\Users\user\Desktop\Libyana_Data"), "Output", "PS_Traffic_Output"),
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
@st.cache_resource
def get_rg():
    return ReportGenerator()


rg = get_rg()


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


@st.cache_data(ttl=600)
def cached_packet_loss_report(period, target_date):
    return rg.build_packet_loss_report(period, target_date)


@st.cache_data(ttl=600)
def cached_topology_table():
    path = os.path.join('config', 'site_topology.csv')
    return pd.read_csv(path) if os.path.exists(path) else None


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


@st.cache_data(ttl=600)
def cached_ps_traffic_report():
    """The combined PS-traffic-per-site workbook (all sheets) from the sibling puller."""
    candidates = sorted(
        glob.glob(os.path.join(PS_TRAFFIC_OUTPUT_DIR, "Combined_Traffic_Report.xlsx"))
        or glob.glob(os.path.join(PS_TRAFFIC_OUTPUT_DIR, "PS_Traffic_Combined_Report_*.xlsx")),
        key=os.path.getmtime, reverse=True,
    )
    if not candidates:
        return None, None, None
    path = candidates[0]
    sheets = pd.read_excel(path, sheet_name=None)
    with open(path, 'rb') as f:
        raw_bytes = f.read()
    return sheets, os.path.basename(path), raw_bytes


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


# SmartCare CEM refreshes weekly upstream - the long TTL just avoids
# re-reading the workbook from disk on every rerun within a session.
@st.cache_data(ttl=3600)
def cached_cem_overview():
    return rg.build_cem_overview()


@st.cache_data(ttl=3600)
def cached_device_penetration_overview():
    return rg.build_device_penetration_overview()


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
# 📶 PACKET LOSS — IUB/ABIS backhaul ping quality (Transmission_KPIs.csv,
# already archived by backend/transmission_kpi_processor.py)
# ============================================================
elif section == "📶 Packet Loss":
    st.caption("IUB/ABIS backhaul ping packet loss & delay, ranked by Avg Packet Loss(%) descending. "
               "A site/adjacency appearing here consistently across day/week/month is a transport link "
               "issue worth escalating, not a one-off blip.")

    pl_period_labels = {'Day': 'day', 'Last 7 Days': 'week', 'Last 30 Days': 'month'}
    c1, c2 = st.columns([1, 3])
    with c1:
        pl_period_choice = st.radio("Period", list(pl_period_labels.keys()), key="pl_period")
    pl_period = pl_period_labels[pl_period_choice]

    pl_report = cached_packet_loss_report(pl_period, target_date)
    if pl_report is None or pl_report.empty:
        st.info("No Transmission KPI data available for this period.")
    else:
        elevated_threshold = rg.PACKET_LOSS_ELEVATED_PCT
        elevated = pl_report[pl_report['Avg Packet Loss(%)'] > elevated_threshold]
        m1, m2, m3 = st.columns(3)
        m1.metric("Adjacencies Reporting", f"{len(pl_report):,}")
        m2.metric(f"Elevated (Avg > {elevated_threshold}%)", f"{len(elevated):,}")
        m3.metric("Worst Avg Packet Loss", f"{pl_report['Avg Packet Loss(%)'].max():.2f}%")

        pl_search = st.text_input("Filter by Site Name / Adjacent Node Name contains...", "", key="pl_search")
        pl_filtered = pl_report
        if pl_search:
            name_cols = [c for c in ['Site Name', 'Adjacent Node Name'] if c in pl_filtered.columns]
            mask = False
            for c in name_cols:
                mask = mask | pl_filtered[c].astype(str).str.contains(pl_search, case=False, na=False)
            pl_filtered = pl_filtered[mask]

        st.caption(f"{len(pl_filtered):,} adjacenc(y/ies) — {pl_period_choice.lower()}, ending {target_date}")
        st.dataframe(pl_filtered, width='stretch', hide_index=True, height=450)
        render_word_export_button(
            "Packet Loss", [(f"Packet Loss ({pl_period_choice})", pl_filtered)],
            key_prefix="packet_loss", subtitle=target_date,
        )

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
        st.caption("Pick any day to see how many sites went down, total summed downtime hours, and each "
                   "down site's NetEco (power) and NCE (transmission) alarm evidence merged into one row. "
                   "Computed directly from this project's own raw MAE/NetEco/NCE historical exports, as "
                   "far back as their retention window covers (currently a couple of weeks).")

        daily_date = st.date_input("Date", value=pd.Timestamp.now().normalize() - pd.Timedelta(days=1),
                                    key="daily_noc_alarm_date")
        daily_report = cached_daily_noc_alarm_report(daily_date.strftime('%Y-%m-%d'))

        if not daily_report.get('available'):
            st.info(f"No raw historical alarm export covers {daily_date} — either the NOC alarm feed "
                    f"isn't available on this machine, or that date has aged out of the retention window.")
        else:
            src_status = " | ".join(
                f"{src}: {'✅' if info['available'] else '❌ unavailable'}"
                for src, info in daily_report['sources'].items()
            )
            st.caption(src_status)

            daily_metrics = daily_report.get('metrics', {})
            dm1, dm2, dm3, dm4 = st.columns(4)
            dm1.metric("Count of NE Is Disconnected", f"{daily_report['sites_down']:,}")
            dm2.metric("Total Downtime (Hours)", f"{daily_report['total_down_hours']:,.1f}")
            dm3.metric("Count of Mains Failure", f"{daily_metrics.get('mains_failure', 0):,}")
            dm4.metric("Count of NCE Transmission Alarm Sites", f"{daily_metrics.get('nce_transmission', 0):,}")

            down_summary = daily_report['down_sites_summary']
            if down_summary.empty:
                st.success(f"No sites recorded as down on {daily_date}.")
            else:
                daily_search = st.text_input("Filter by Site contains...", "", key="daily_noc_search")
                down_filtered = down_summary[down_summary['Site Name'].astype(str).str.contains(
                    daily_search, case=False, na=False)] if daily_search else down_summary
                st.dataframe(down_filtered, width='stretch', hide_index=True, height=400)
                render_word_export_button(
                    "Daily NOC Alarm Analysis", [("Down Sites", down_filtered)],
                    key_prefix="daily_noc_alarm", subtitle=str(daily_date),
                )

            daily_tabs = st.tabs(["🔋 NetEco Alarms (this day)", "📡 NCE Alarms (this day)"])
            with daily_tabs[0]:
                ne = daily_report['neteco_alarms']
                if ne.empty:
                    st.info("No NetEco alarm data for this date.")
                else:
                    st.dataframe(ne, width='stretch', hide_index=True, height=350)
            with daily_tabs[1]:
                nc = daily_report['nce_alarms']
                if nc.empty:
                    st.info("No NCE alarm data for this date.")
                else:
                    st.dataframe(nc, width='stretch', hide_index=True, height=350)

# ============================================================
# 📱 CEM — SmartCare Customer Experience Management: app traffic/TCP
# quality + weekly Device Penetration mix. The CEM half now runs from this
# project's own scrapers/smartcare_cem_scraper.py + reports/
# run_smartcare_analysis_task.py (weekly, on-demand - see run_smartcare_
# pipeline.bat); Device Mix still reads a sibling suite's weekly export -
# see backend/smartcare_cem_processor.py and backend/device_penetration_processor.py
# ============================================================
elif section == "📱 CEM":
    st.caption("Network-wide subscriber experience data from the SmartCare portal (DPI probe traffic mix, "
               "TCP-level connection quality, and device model mix) - a different data source than the "
               "counter-based KPIs elsewhere in this dashboard, refreshed weekly. No per-site breakdown is "
               "available.")

    sec_tabs = st.tabs(["📶 App Traffic & Quality", "📱 Device Mix"])

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

            snap = dp_overview['latest_snapshot']
            st.caption(f"Snapshot: {dp_overview['latest_time']} — {len(snap):,} device models")

            if 'Broadband Capable' in snap.columns and 'Number of Users(count)' in snap.columns:
                total_users = snap['Number of Users(count)'].sum()
                broadband_users = snap.loc[snap['Broadband Capable'], 'Number of Users(count)'].sum()
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Devices Seen", f"{total_users:,.0f}")
                m2.metric("LTE/NR-Capable", f"{100 * broadband_users / total_users:.1f}%" if total_users else "N/A")
                m3.metric("Legacy-Only (2G/3G)", f"{100 * (1 - broadband_users / total_users):.1f}%" if total_users else "N/A")
                st.caption("A high legacy-only share is context for capacity planning: even an idle new LTE "
                           "band gets little uptake at a site until subscriber devices there catch up.")

            c1, c2 = st.columns(2)
            if 'Device Type' in snap.columns and 'Number of Users(count)' in snap.columns:
                with c1:
                    st.subheader("📱 By Device Type")
                    type_totals = snap.groupby('Device Type')['Number of Users(count)'].sum().sort_values(ascending=False)
                    fig_type = go.Figure(go.Bar(x=type_totals.values, y=type_totals.index, orientation='h'))
                    fig_type.update_layout(height=350, margin=dict(l=120, r=20, t=20, b=30),
                                            yaxis=dict(autorange='reversed'), xaxis_title='Users')
                    st.plotly_chart(fig_type, width='stretch', key='dp_type_chart')
            if 'Device Brand' in snap.columns and 'Number of Users(count)' in snap.columns:
                with c2:
                    st.subheader("🏷️ Top Brands")
                    brand_totals = snap.groupby('Device Brand')['Number of Users(count)'].sum().sort_values(ascending=False).head(10)
                    fig_brand = go.Figure(go.Bar(x=brand_totals.values, y=brand_totals.index, orientation='h'))
                    fig_brand.update_layout(height=350, margin=dict(l=120, r=20, t=20, b=30),
                                             yaxis=dict(autorange='reversed'), xaxis_title='Users')
                    st.plotly_chart(fig_brand, width='stretch', key='dp_brand_chart')

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
        "📶 Monthly Cell Info Report", "📈 PS Traffic per site",
    ])

    with sec_tabs[0]:
        st.caption("EAST branch only, built from output/csv/ history. Not included yet: "
                   "Network Daily KPI's (a few columns still need source confirmation).")

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

    with sec_tabs[3]:
        st.caption(
            "Daily PS (data) traffic per site across 2G/3G/4G, pulled and combined by the "
            f"sibling \"PS Traffic per site\" puller. Read directly from {PS_TRAFFIC_OUTPUT_DIR}."
        )

        ps_sheets, ps_filename, ps_bytes = cached_ps_traffic_report()
        if ps_sheets is None:
            st.info(
                "No combined PS Traffic report found yet. Run the sibling project's "
                "\"PS Traffic per site v3.py\" puller to build it."
            )
        else:
            st.caption(f"Source file: {ps_filename}")
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
                    st.dataframe(shown, width='stretch', hide_index=True)
                    st.caption(f"{len(shown)} row(s)")

            st.divider()
            st.download_button(
                "⬇️ Download PS Traffic per site Report (.xlsx)", data=ps_bytes,
                file_name=ps_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

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
