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

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from backend.report_generator import (
    ReportGenerator, TECH_LABELS, CELL_SHEETS, SCORECARD_SHEETS, SITE_COL_BY_TECH,
)
from backend import ept_manager as ept

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
def cached_worst_cells(target_date, n_top):
    return rg.build_worst_cells(target_date, n_top=n_top)


@st.cache_data(ttl=600)
def cached_cell_failing(tech, cell_names, target_date):
    return rg.get_cell_failing_kpis(tech, list(cell_names), target_date)


@st.cache_data(ttl=600)
def cached_cell_thresholds(tech, sheet):
    return rg.get_trend_kpi_thresholds(tech, sheet=sheet)


@st.cache_data(ttl=600)
def cached_word_bytes(target_date, previous_date):
    b = cached_bundle(target_date, previous_date)
    path = rg.generate_word_report(
        target_date, previous_date, b['health'], b['scorecards'], b['worst_cells'],
        b['site_health'], b['topology'], b['traffic'], b['site_inventory'],
        b['freshness'], b['trend'],
    )
    with open(path, 'rb') as f:
        return f.read()


@st.cache_data(ttl=600)
def cached_excel_bytes(target_date, previous_date):
    b = cached_bundle(target_date, previous_date)
    path = rg.generate_excel_report(
        target_date, b['health'], b['scorecards'], b['worst_cells'], b['site_health'],
        b['topology'], b['traffic'], b['site_inventory'], b['freshness'], b['trend'],
    )
    with open(path, 'rb') as f:
        return f.read()


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


# ============================================================
# HEADER + SIDEBAR
# ============================================================
st.markdown("""
<div class="header-container">
    <h1 style="margin:0;">📊 Libyana Network Performance Dashboard</h1>
    <p style="margin:0;">EAST Region | 2G / 3G / 4G — live from the daily KPI pipeline</p>
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
         "🔎 Investigate", "📧 Reports"],
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
        st.caption(f"As of {target_date} | Region: EAST")

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
# 🏗️ SITES & INFRASTRUCTURE — Site Health & Topology, Site Inventory, EPT
# ============================================================
elif section == "🏗️ Sites & Infrastructure":
    sec_tabs = st.tabs(["🏗️ Site Health & Topology", "🏢 Site Inventory", "🗺️ EPT"])

    with sec_tabs[0]:
        st.dataframe(bundle['site_health'], width='stretch', hide_index=True)
        topo = bundle['topology']
        if topo.get('loaded'):
            st.success(
                f"Topology reference loaded: {topo['nodes']} nodes "
                f"({topo['fn_count']} FN, {topo['hub_count']} HUB), "
                f"{topo['site_relationships']} site relationships, "
                f"regions: {', '.join(topo['regions'])}."
            )
        else:
            st.warning("Topology reference file not found.")
        st.caption("Alarm-to-topology impact correlation is Phase 3/4, pending NetEco integration.")
        render_word_export_button(
            "Site Health & Topology", [("Availability by Technology", bundle['site_health'])],
            key_prefix="site_health", subtitle=target_date,
        )

    with sec_tabs[1]:
        st.dataframe(bundle['site_inventory'], width='stretch', hide_index=True)
        st.caption("Site outages: pending NetEco alarm integration (Phase 3).")
        render_word_export_button(
            "Site Inventory", [("Multi-RAT Site Composition", bundle['site_inventory'])],
            key_prefix="site_inventory", subtitle=target_date,
        )

    with sec_tabs[2]:
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
                        selected = rg.resolve_group_to_cells(explore_tech, picked, group_col=site_col)
                        st.caption(f"Resolved to {len(selected)} cell(s)/sector(s) under {len(picked)} selected site(s).")
                    else:
                        selected = picked

                if picked and not selected:
                    st.warning("No cells found under the selected site(s) for this technology/date.")
                elif picked:
                    st.subheader(f"📋 Combined KPIs — {len(selected)} cell(s)")
                    combined = filtered[filtered[name_col].isin(selected)]
                    st.dataframe(combined, width='stretch', hide_index=True)

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
                        kpi_cols = [c for c in cell_trend.columns if c not in ('Date', 'Cell')]
                        st.caption("Dashed line = threshold. One line per selected cell; hover to compare.")
                        cols = st.columns(2)
                        for i, kpi in enumerate(kpi_cols):
                            with cols[i % 2]:
                                render_multi_cell_trend_chart(
                                    cell_trend, kpi, kpi_thresholds.get(kpi),
                                    key=f"cell_trend_{explore_tech}_{kpi}",
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
            sr_options = sorted(sr_raw[sr_pick_col].dropna().astype(str).unique().tolist()) \
                if sr_pick_col in sr_raw.columns else []
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
                    sr_cells = rg.resolve_group_to_cells(sr_tech, sr_picked, group_col=sr_site_col)
                    st.caption(f"Resolved to {len(sr_cells)} cell(s)/sector(s) under {len(sr_picked)} selected site(s).")
                else:
                    sr_cells = sr_picked

                if not sr_cells:
                    st.warning("No cells found for the selected group.")
                else:
                    sr_trend = rg.build_cell_trend(sr_tech, sr_cells, start_date=sr_start, end_date=sr_end)
                    if sr_trend is None or sr_trend.empty:
                        st.info("No data available for this group/date range.")
                    else:
                        latest_date = sr_trend['Date'].max()
                        st.subheader(f"📋 Combined KPIs — {latest_date}")
                        st.dataframe(sr_trend[sr_trend['Date'] == latest_date].drop(columns=['Date']),
                                     width='stretch', hide_index=True)

                        st.subheader("⚠️ Failing KPIs & Suggested Fixes")
                        sr_failing = rg.get_cell_failing_kpis(sr_tech, sr_cells, latest_date)
                        if sr_failing is not None and not sr_failing.empty:
                            st.dataframe(sr_failing, width='stretch', hide_index=True)
                        else:
                            st.success("No threshold KPIs are failing for this group on the latest date.")

                        st.subheader(f"📈 Trend ({sr_start} to {sr_end})")
                        sr_kpi_thresholds = rg.get_trend_kpi_thresholds(sr_tech, sheet=sr_sheet)
                        sr_kpi_cols = [c for c in sr_trend.columns if c not in ('Date', 'Cell')]
                        sr_cols = st.columns(2)
                        for i, kpi in enumerate(sr_kpi_cols):
                            with sr_cols[i % 2]:
                                render_multi_cell_trend_chart(
                                    sr_trend, kpi, sr_kpi_thresholds.get(kpi), key=f"sr_trend_{sr_tech}_{kpi}",
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
# 📧 REPORTS — network summary + copy-paste text + Word/Excel export
# ============================================================
elif section == "📧 Reports":
    sec_tabs = st.tabs(["📧 Report & Export"])

    with sec_tabs[0]:
        st.subheader("📈 Network Summary")
        summary_rows = rg.build_network_summary_block(target_date)
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

        st.divider()
        st.subheader("✉️ Copy-Paste Text")
        st.caption("Identical content to the automated daily report — select all and copy into an email.")
        email_text = rg.generate_email_text(
            target_date, previous_date, health, bundle['scorecards'], bundle['worst_cells'],
            bundle['site_health'], bundle['topology'], bundle['traffic'],
            bundle['site_inventory'], bundle['freshness'], bundle['trend'],
        )
        st.text_area("Report text", email_text, height=500)
