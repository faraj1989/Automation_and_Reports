#!/usr/bin/env python3
"""
Libyana NPM - Packet Loss Engine (IUB/ABIS backhaul ping quality)

Pure, vectorized logic behind the Packet Loss pipeline and dashboard tab.
The raw input is the BSC6900 PACKETLOSS export: a rolling 7-day window of
hourly ping loss/delay per backhaul link (ABIS = 2G BTS<->BSC, IUB = 3G
NodeB<->RNC, plus a few core links), re-exported every day.

    RAW 7-DAY FILE
      1. normalize_raw        key = GBSC + Adjacent Node ID (IDs repeat across
                              BSCs), NIL -> NaN = "no ping response", RAT, Site
      2. link-hour table      one row per link per hour (ABIS / IUB preserved)
      3. build_site_hours     worse of ABIS/IUB per site-hour, affected RAT,
                              no-response, delay anomaly vs the link's own baseline
      4. detect_hub_events    same FN/HUB node, same hour, >= N sites, >= X% of node
      5. build_site_daily     per site per day counts (the permanent archive)
      6. classify_sites       window (1/7/30 days or any range) -> class,
                              persistence, pattern, suspected cause

Why hour counts instead of averages: a daily average hides intermittent
problems (a site with five 1-5% loss hours and 19 clean hours averages
~0.08% and looks healthy), so everything downstream is driven by how MANY
hours a site was hurting, with the average kept only as a secondary figure.

All thresholds live in config/packet_loss_rules.csv (see load_rules).
"""

import os
import logging
from datetime import datetime

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RULES_PATH = os.path.join(_BASE_DIR, 'config', 'packet_loss_rules.csv')
TOPOLOGY_PATH = os.path.join(_BASE_DIR, 'config', 'site_topology.csv')

DEFAULT_RULES = {
    'loss_hour_pct': 1.0,
    'minor_hour_pct': 0.1,
    'chronic_share': 0.5,
    'recurrent_min_days': 3,
    'recurrent_day_share': 0.2,
    'recurrent_min_hours_short': 3,
    'hub_min_sites': 3,
    'hub_min_share': 0.6,
    'hub_cause_share': 0.5,
    'delay_ratio': 2.0,
    'delay_margin_ms': 20.0,
    'busy_hours': [20, 21, 22, 23, 0, 1],
    'quiet_hours': [3, 4, 5, 6],
    'pattern_ratio': 2.0,
    'pattern_min_diff': 0.5,
    'hourly_retention_days': 31,
}

RAW_METRICS = {
    'T7816:Average Ping Packet Loss Rate of Adjacent Node(%)': 'Loss',
    'T7817:Maximum Ping Packet Loss Rate of Adjacent Node(%)': 'Max Loss',
    'T7812:Average Ping Delay of Adjacent Node(ms)': 'Delay',
    'T7813:Maximum Ping Delay of Adjacent Node(ms)': 'Max Delay',
}

RAT_BY_TYPE = {'ABIS': '2G', 'IUB': '3G'}

CLASS_ORDER = ['🔴 Chronic', '🟠 Recurrent', '🟡 Sporadic', '⚫ Outage only', '⚪ Not reporting', '🟢 Clean']
FLAGGED_CLASSES = CLASS_ORDER[:2]


# ---------------------------------------------------------------- config

def load_rules(path=RULES_PATH):
    """config/packet_loss_rules.csv (Parameter,Value,Description) over the
    built-in defaults, so a missing/partial file never breaks the pipeline."""
    rules = dict(DEFAULT_RULES)
    try:
        cfg = pd.read_csv(path, dtype=str)
        for _, row in cfg.iterrows():
            key, val = str(row['Parameter']).strip(), str(row['Value']).strip()
            if key not in rules:
                continue
            if isinstance(DEFAULT_RULES[key], list):
                rules[key] = [int(x) for x in val.split(',') if x.strip() != '']
            elif isinstance(DEFAULT_RULES[key], int):
                rules[key] = int(float(val))
            else:
                rules[key] = float(val)
    except Exception as e:
        logger.warning(f"Packet loss rules: using defaults ({e})")
    return rules


def load_topology(path=TOPOLOGY_PATH):
    """FN/HUB membership from config/site_topology.csv (rebuilt from the FN-HUB
    Excel). A site usually belongs to several nodes (its own FN, the upstream
    FN, the HUB), so this returns every (Node, Site) pair plus per-site
    Region and a readable chain string."""
    empty = (pd.DataFrame(columns=['Node', 'Node Type', 'Site']), {}, {})
    if not os.path.exists(path):
        return empty
    try:
        t = pd.read_csv(path, dtype=str)
    except Exception as e:
        logger.warning(f"Could not read topology {path}: {e}")
        return empty
    t = t.dropna(subset=['Node_Name', 'Connected_Site'])
    pairs = pd.DataFrame({
        'Node': t['Node_Name'].str.strip().str.upper(),
        'Node Type': t['Node_Type'].fillna('').str.strip(),
        'Site': t['Connected_Site'].str.strip().str.upper(),
        'Region': t['Region'].fillna('').str.strip(),
    }).drop_duplicates(['Node', 'Site'])

    region = pairs.groupby('Site')['Region'].first().to_dict()
    chain_src = pairs[pairs['Node'] != pairs['Site']]
    chain = (chain_src.assign(lbl=chain_src['Node'] + ' (' + chain_src['Node Type'] + ')')
             .groupby('Site')['lbl'].apply(lambda s: ', '.join(sorted(set(s)))).to_dict())
    return pairs[['Node', 'Node Type', 'Site']], region, chain


def _region_fallback(site):
    """Region from the site-name prefix when the site isn't in the topology
    (e.g. BGZ052 -> BGZ, COAST010 -> COAST)."""
    s = pd.Series(site, dtype='object').astype(str)
    return s.str.extract(r'^([A-Za-z]+)', expand=False).str.upper()


# ---------------------------------------------------------------- step 1-2

def normalize_raw(raw: pd.DataFrame) -> pd.DataFrame:
    """Raw export -> link-hour table.

    - Key is (GBSC, Adjacent Node ID): IDs are only unique per BSC (ID 21 is
      ECV001 on BGZMBSC01 but Test2 on BYDMBSC01).
    - NIL loss (no valid ping samples - almost always with delay 0) becomes
      NaN and is flagged 'No Response'; it is never treated as 0% loss.
    - Site: ABIS rows carry it; IUB rows leave Site Name blank, so it's the
      node name minus its leading 'U' (UKUFR004 -> KUFR004). Upper-cased so
      both RATs (and the topology) join on the same key."""
    df = raw.rename(columns=RAW_METRICS).copy()
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
    df['Adjacent Node ID'] = pd.to_numeric(df['Adjacent Node ID'], errors='coerce')
    df = df.dropna(subset=['Time', 'GBSC', 'Adjacent Node ID'])
    df['Adjacent Node ID'] = df['Adjacent Node ID'].astype('int64')

    for col in RAW_METRICS.values():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        else:
            df[col] = np.nan
    df['Integrity'] = pd.to_numeric(df.get('Integrity', pd.Series(dtype=str)).astype(str).str.rstrip('%'),
                                    errors='coerce')
    df['Backward Bandwidth'] = pd.to_numeric(df.get('Backward Bandwidth'), errors='coerce')

    ntype = df['Adjacent Node Type'].astype(str).str.strip().str.upper()
    name = df['Adjacent Node Name'].astype(str).str.strip()
    site_col = df['Site Name'].astype(str).str.strip() if 'Site Name' in df.columns else pd.Series('', index=df.index)
    site_col = site_col.where(~site_col.isin(['', 'nan', 'None']))
    iub_site = name.where(~(name.str.upper().str.startswith('U') & (name.str.len() > 1)), name.str[1:])
    df['Site'] = np.select(
        [ntype.eq('ABIS'), ntype.eq('IUB')],
        [site_col.fillna(name), iub_site],
        default=None,
    )
    df['Site'] = df['Site'].astype('object').where(df['Site'].notna(), None)
    df['Site'] = df['Site'].map(lambda s: s.upper() if isinstance(s, str) else s)
    df['Adjacent Node Type'] = ntype
    df['RAT'] = ntype.map(RAT_BY_TYPE).fillna('Core')
    df['Date'] = df['Time'].dt.strftime('%Y-%m-%d')
    df['Hour'] = df['Time'].dt.hour
    df['No Response'] = df['Loss'].isna()
    return df


def complete_dates(link_hours: pd.DataFrame, today=None):
    """Dates this export fully covers (all 24 hours present) and that are
    already over - KPI history must only hold complete business days."""
    today = pd.Timestamp(today or datetime.now().date()).strftime('%Y-%m-%d')
    hours = link_hours.groupby('Date')['Hour'].nunique()
    return sorted(d for d, n in hours.items() if n >= 24 and d < today)


def add_delay_flags(link_hours: pd.DataFrame, rules) -> pd.DataFrame:
    """High-delay hour = delay well above the link's OWN normal level (its
    median over this export). A fixed ms limit doesn't work: satellite links
    (e.g. TAZB002) sit at ~557 ms permanently, most fiber/MW links at 1-15 ms."""
    df = link_hours
    base = df.groupby(['GBSC', 'Adjacent Node ID'])['Delay'].transform('median')
    limit = np.maximum(base * rules['delay_ratio'], base + rules['delay_margin_ms'])
    df['Delay Baseline'] = base
    df['High Delay'] = (df['Delay'] > limit).fillna(False)
    return df


# ---------------------------------------------------------------- step 2b: link daily

def build_link_daily(link_hours: pd.DataFrame, rules) -> pd.DataFrame:
    """One row per (Date, GBSC, Adjacent Node ID) - Transmission_KPIs.csv.
    Keeps the original Avg/Max columns (Network Daily KPI's report reads
    them) and adds the hour counts."""
    df = link_hours.assign(
        _loss=link_hours['Loss'] >= rules['loss_hour_pct'],
        _minor=(link_hours['Loss'] >= rules['minor_hour_pct']) & (link_hours['Loss'] < rules['loss_hour_pct']),
    )
    out = df.groupby(['Date', 'GBSC', 'Adjacent Node ID'], as_index=False).agg(**{
        'Site Name': ('Site', 'first'),
        'Adjacent Node Name': ('Adjacent Node Name', 'first'),
        'Adjacent Node Type': ('Adjacent Node Type', 'first'),
        'RAT': ('RAT', 'first'),
        'BTSID': ('BTSID', 'first'),
        'Backward Bandwidth': ('Backward Bandwidth', 'first'),
        'Integrity': ('Integrity', 'min'),
        'Avg Packet Loss(%)': ('Loss', 'mean'),
        'Max Packet Loss(%)': ('Max Loss', 'max'),
        'Avg Delay(ms)': ('Delay', 'mean'),
        'Max Delay(ms)': ('Max Delay', 'max'),
        'Hours Reported': ('Loss', 'count'),
        'Loss Hours': ('_loss', 'sum'),
        'Minor Hours': ('_minor', 'sum'),
        'No Response Hours': ('No Response', 'sum'),
        'High Delay Hours': ('High Delay', 'sum'),
    })
    for c in ['Avg Packet Loss(%)', 'Max Packet Loss(%)', 'Avg Delay(ms)', 'Max Delay(ms)']:
        out[c] = out[c].round(4)
    return out


# ---------------------------------------------------------------- step 3

def build_site_hours(link_hours: pd.DataFrame) -> pd.DataFrame:
    """One row per (Site, Time), ABIS/IUB only. The WORSE of the two RATs
    decides site health: they ride the same physical transmission path
    (hourly ABIS vs IUB loss correlate at ~0.91), so one lossy RAT is enough
    to say the site's backhaul is hurting. Per-RAT values are kept so a
    RAT-specific problem (ABIS-only / IUB-only) is still visible."""
    df = link_hours[link_hours['RAT'].isin(['2G', '3G']) & link_hours['Site'].notna()]
    keys = ['Site', 'Time']
    site = df.groupby(keys, as_index=False).agg(
        GBSC=('GBSC', 'first'),
        Loss=('Loss', 'max'),
        Delay=('Delay', 'max'),
        NoResp=('No Response', 'all'),
        HighDelay=('High Delay', 'any'),
    )
    per_rat = df.pivot_table(index=keys, columns='RAT', values='Loss', aggfunc='max')
    nr_rat = df.assign(nr=df['No Response'].astype(int)).pivot_table(
        index=keys, columns='RAT', values='nr', aggfunc='min')
    per_rat = per_rat.reindex(columns=['2G', '3G']).add_prefix('Loss_')
    nr_rat = nr_rat.reindex(columns=['2G', '3G']).add_prefix('NoResp_')
    site = site.merge(per_rat.reset_index(), on=keys, how='left').merge(nr_rat.reset_index(), on=keys, how='left')
    site['Date'] = site['Time'].dt.strftime('%Y-%m-%d')
    site['Hour'] = site['Time'].dt.hour
    rats = df.groupby('Site')['RAT'].apply(lambda s: '+'.join(sorted(set(s))))
    site['RATs'] = site['Site'].map(rats)
    return site


# ---------------------------------------------------------------- step 4

def detect_hub_events(site_hours: pd.DataFrame, pairs: pd.DataFrame, rules):
    """Same FN/HUB node + same hour + >= hub_min_sites sites affected (loss
    hour or no response) + those are >= hub_min_share of the node's sites.

    Because a site belongs to a chain of nodes (own FN -> upstream FN -> HUB),
    a failing HUB also makes every FN under it qualify. Each affected
    site-hour is attributed to the qualifying node with the MOST affected
    sites - the highest failing point in the chain - and only those nodes are
    reported, so one hub outage is one event, not one per downstream FN.

    Returns (site_hours with 'Hub Event Node', hub-hour events DataFrame)."""
    sh = site_hours.copy()
    sh['Hub Event Node'] = None
    ev_cols = ['Node', 'Time', 'Affected', 'Total', 'Share', 'LossSites', 'NoRespSites', 'AvgLoss']
    if pairs is None or pairs.empty:
        return sh, pd.DataFrame(columns=ev_cols)

    present = set(sh['Site'].unique())
    p = pairs[pairs['Site'].isin(present)]
    node_total = p.groupby('Node')['Site'].nunique()

    aff = sh.loc[(sh['Loss'] >= rules['loss_hour_pct']) | sh['NoResp'], ['Site', 'Time', 'Loss', 'NoResp']]
    if aff.empty:
        return sh, pd.DataFrame(columns=ev_cols)
    m = aff.merge(p[['Node', 'Site']], on='Site')
    m['IsLoss'] = ~m['NoResp']
    cnt = m.groupby(['Node', 'Time'], as_index=False).agg(
        Affected=('Site', 'nunique'), LossSites=('IsLoss', 'sum'),
        NoRespSites=('NoResp', 'sum'), AvgLoss=('Loss', 'mean'))
    cnt['Total'] = cnt['Node'].map(node_total)
    cnt['Share'] = cnt['Affected'] / cnt['Total']
    ev = cnt[(cnt['Affected'] >= rules['hub_min_sites']) & (cnt['Share'] >= rules['hub_min_share'])]
    if ev.empty:
        return sh, pd.DataFrame(columns=ev_cols)

    cand = m[['Site', 'Time', 'Node']].merge(ev[['Node', 'Time', 'Affected', 'Total']], on=['Node', 'Time'])
    best = (cand.sort_values(['Affected', 'Total', 'Node'], ascending=[False, False, True])
                .drop_duplicates(['Site', 'Time']))
    sh = sh.drop(columns='Hub Event Node').merge(
        best[['Site', 'Time', 'Node']].rename(columns={'Node': 'Hub Event Node'}), on=['Site', 'Time'], how='left')

    kept = best[['Node', 'Time']].drop_duplicates()
    ev = ev.merge(kept, on=['Node', 'Time'])
    ev.attrs['members'] = best  # site-level attribution, used for the daily roll-up
    return sh, ev[ev_cols]


def build_hub_events_daily(events: pd.DataFrame, pairs: pd.DataFrame, region: dict) -> pd.DataFrame:
    """Hub-hour events rolled up to one row per (Date, Hub) - the permanent
    Packet_Loss_Hub_Events archive."""
    cols = ['Date', 'Hub', 'Node Type', 'Region', 'Event Hours', 'Loss Event Hours',
            'No Response Event Hours', 'Peak Sites Affected', 'Sites On Hub', 'Peak Share (%)',
            'Avg Loss During Events (%)', 'Event Type', 'First Hour', 'Last Hour', 'Affected Sites']
    if events is None or events.empty:
        return pd.DataFrame(columns=cols)
    members = events.attrs.get('members')
    ev = events.copy()
    ev['Date'] = ev['Time'].dt.strftime('%Y-%m-%d')
    ev['Hour'] = ev['Time'].dt.hour
    ev['LossDominant'] = ev['LossSites'] >= ev['NoRespSites']
    out = ev.groupby(['Date', 'Node'], as_index=False).agg(**{
        'Event Hours': ('Time', 'nunique'),
        'Loss Event Hours': ('LossDominant', 'sum'),
        'Peak Sites Affected': ('Affected', 'max'),
        'Sites On Hub': ('Total', 'max'),
        'Peak Share (%)': ('Share', 'max'),
        'Avg Loss During Events (%)': ('AvgLoss', 'mean'),
        'First Hour': ('Hour', 'min'),
        'Last Hour': ('Hour', 'max'),
    })
    out['No Response Event Hours'] = out['Event Hours'] - out['Loss Event Hours']
    out['Event Type'] = np.select(
        [out['No Response Event Hours'].eq(0), out['Loss Event Hours'].eq(0)],
        ['Packet loss', 'No response (link/power down)'], default='Mixed')
    out['Peak Share (%)'] = (out['Peak Share (%)'] * 100).round(1)
    out['Avg Loss During Events (%)'] = out['Avg Loss During Events (%)'].round(3)
    if members is not None and not members.empty:
        mem = members.assign(Date=members['Time'].dt.strftime('%Y-%m-%d'))
        sites = mem.groupby(['Date', 'Node'])['Site'].apply(lambda s: ', '.join(sorted(set(s))))
        out['Affected Sites'] = [sites.get((d, n), '') for d, n in zip(out['Date'], out['Node'])]
    else:
        out['Affected Sites'] = ''
    ntype = pairs.drop_duplicates('Node').set_index('Node')['Node Type'] if not pairs.empty else pd.Series(dtype=str)
    out['Node Type'] = out['Node'].map(ntype)
    out['Region'] = out['Node'].map(region)
    out['Region'] = out['Region'].fillna(_region_fallback(out['Node']))
    out = out.rename(columns={'Node': 'Hub'})
    return out[cols].sort_values(['Date', 'Event Hours'], ascending=[True, False]).reset_index(drop=True)


# ---------------------------------------------------------------- step 5

def build_site_daily(site_hours: pd.DataFrame, rules, region: dict, chain: dict) -> pd.DataFrame:
    """One row per (Date, Site) - the permanent Packet_Loss_Site_Daily
    archive. Everything the window classification needs is here as counts
    or means, so 7/30-day (or any range) analysis never needs hourly data."""
    sh = site_hours
    L, M = rules['loss_hour_pct'], rules['minor_hour_pct']
    w = sh.assign(
        _measured=sh['Loss'].notna(),
        _loss=sh['Loss'] >= L,
        _minor=(sh['Loss'] >= M) & (sh['Loss'] < L),
        _loss2=sh['Loss_2G'] >= L,
        _loss3=sh['Loss_3G'] >= L,
        _hub=sh['Hub Event Node'].notna(),
        _busy=sh['Loss'].where(sh['Hour'].isin(rules['busy_hours'])),
        _quiet=sh['Loss'].where(sh['Hour'].isin(rules['quiet_hours'])),
        _nr2=sh['NoResp_2G'].fillna(0).astype(bool),
        _nr3=sh['NoResp_3G'].fillna(0).astype(bool),
    )
    keys = ['Date', 'Site']
    out = w.groupby(keys, as_index=False).agg(**{
        'GBSC': ('GBSC', 'first'),
        'RATs': ('RATs', 'first'),
        'Hours Measured': ('_measured', 'sum'),
        'Loss Hours': ('_loss', 'sum'),
        'Minor Hours': ('_minor', 'sum'),
        'No Response Hours': ('NoResp', 'sum'),
        'No Response Hours 2G': ('_nr2', 'sum'),
        'No Response Hours 3G': ('_nr3', 'sum'),
        'High Delay Hours': ('HighDelay', 'sum'),
        'Hub Event Hours': ('_hub', 'sum'),
        'Loss Hours 2G': ('_loss2', 'sum'),
        'Loss Hours 3G': ('_loss3', 'sum'),
        'Avg Loss (%)': ('Loss', 'mean'),
        'Avg Loss 2G (%)': ('Loss_2G', 'mean'),
        'Avg Loss 3G (%)': ('Loss_3G', 'mean'),
        'Worst Hour Loss (%)': ('Loss', 'max'),
        'Busy Avg Loss (%)': ('_busy', 'mean'),
        'Quiet Avg Loss (%)': ('_quiet', 'mean'),
        'Avg Delay (ms)': ('Delay', 'mean'),
        'Max Delay (ms)': ('Delay', 'max'),
    })

    worst = (sh.dropna(subset=['Loss']).sort_values('Loss', ascending=False)
               .drop_duplicates(keys)[keys + ['Hour']].rename(columns={'Hour': 'Worst Hour'}))
    out = out.merge(worst, on=keys, how='left')

    hub = sh.dropna(subset=['Hub Event Node'])
    if not hub.empty:
        top = (hub.groupby(keys + ['Hub Event Node']).size().rename('n').reset_index()
                  .sort_values('n', ascending=False).drop_duplicates(keys)[keys + ['Hub Event Node']])
        out = out.merge(top, on=keys, how='left')
    else:
        out['Hub Event Node'] = None

    out['Region'] = out['Site'].map(region).fillna(_region_fallback(out['Site']))
    out['FN/HUB Chain'] = out['Site'].map(chain)
    for c in [c for c in out.columns if c.endswith('(%)') or c.endswith('(ms)')]:
        out[c] = out[c].round(4)
    cols = ['Date', 'Site', 'Region', 'GBSC', 'RATs', 'FN/HUB Chain', 'Hours Measured', 'Loss Hours',
            'Minor Hours', 'No Response Hours', 'No Response Hours 2G', 'No Response Hours 3G',
            'High Delay Hours', 'Hub Event Hours', 'Hub Event Node', 'Loss Hours 2G', 'Loss Hours 3G',
            'Avg Loss (%)', 'Avg Loss 2G (%)', 'Avg Loss 3G (%)', 'Worst Hour Loss (%)', 'Worst Hour',
            'Busy Avg Loss (%)', 'Quiet Avg Loss (%)', 'Avg Delay (ms)', 'Max Delay (ms)']
    return out[cols].sort_values(keys).reset_index(drop=True)


def build_site_hourly_archive(site_hours: pd.DataFrame) -> pd.DataFrame:
    """Compact hourly site detail (rolling retention) - only for drill-down
    heatmaps, never for classification."""
    sh = site_hours
    out = pd.DataFrame({
        'Time': sh['Time'].dt.strftime('%Y-%m-%d %H:00'),
        'Site': sh['Site'],
        'Loss (%)': sh['Loss'].round(4),
        'Loss 2G (%)': sh['Loss_2G'].round(4),
        'Loss 3G (%)': sh['Loss_3G'].round(4),
        'No Response 2G': sh['NoResp_2G'],
        'No Response 3G': sh['NoResp_3G'],
        'Delay (ms)': sh['Delay'],
        'High Delay': sh['HighDelay'].astype(int),
        'Hub Event Node': sh['Hub Event Node'],
    })
    return out.sort_values(['Time', 'Site']).reset_index(drop=True)


# ---------------------------------------------------------------- one-shot

def process_raw(raw: pd.DataFrame, rules=None, topology=None, today=None, only_dates=None):
    """Full pipeline for one raw export. Returns {sheet_name: DataFrame}
    restricted to complete past days (and to only_dates, if given - used by
    the backfill so an older export never overwrites a newer one's days)."""
    rules = rules or load_rules()
    pairs, region, chain = topology if topology is not None else load_topology()

    lh = normalize_raw(raw)
    days = complete_dates(lh, today)
    if only_dates is not None:
        days = [d for d in days if d in set(only_dates)]
    if not days:
        return {}
    lh = add_delay_flags(lh, rules)          # baseline uses the full 7 days
    site_hours = build_site_hours(lh)
    site_hours, events = detect_hub_events(site_hours, pairs, rules)

    keep = set(days)
    lh_days = lh[lh['Date'].isin(keep)]
    sh_days = site_hours[site_hours['Date'].isin(keep)]
    ev_days = events[events['Time'].dt.strftime('%Y-%m-%d').isin(keep)] if not events.empty else events
    if not events.empty:
        ev_days.attrs['members'] = events.attrs.get('members')

    return {
        'Transmission_KPIs': build_link_daily(lh_days, rules),
        'Packet_Loss_Site_Daily': build_site_daily(sh_days, rules, region, chain),
        'Packet_Loss_Hub_Events': build_hub_events_daily(ev_days, pairs, region),
        'Packet_Loss_Site_Hourly': build_site_hourly_archive(sh_days),
    }


# ---------------------------------------------------------------- step 6

def _window_days(start, end):
    return (pd.Timestamp(end) - pd.Timestamp(start)).days + 1


def classify_sites(site_daily: pd.DataFrame, start, end, rules=None) -> pd.DataFrame:
    """Roll the per-day site archive up over [start, end] and classify.

    Class (first match wins):
      ⚪ Not reporting  no valid ping sample in the whole window
      🔴 Chronic        loss hours >= chronic_share of measured hours
      🟠 Recurrent      loss on >= max(recurrent_min_days, recurrent_day_share x days)
                        different days (1-2 day windows: >= recurrent_min_hours_short loss hours)
      🟡 Sporadic       at least one loss hour
      ⚫ Outage only    no loss, but no-response hours (link/site down)
      🟢 Clean
    Pattern: Traffic-driven (busy-hour loss >> night loss -> capacity),
             Constant (flat -> link quality fault), Burst (< 3 loss hours).
    """
    rules = rules or load_rules()
    if site_daily is None or site_daily.empty:
        return pd.DataFrame()
    s, e = pd.Timestamp(start).strftime('%Y-%m-%d'), pd.Timestamp(end).strftime('%Y-%m-%d')
    d = site_daily[(site_daily['Date'] >= s) & (site_daily['Date'] <= e)].copy()
    if d.empty:
        return pd.DataFrame()
    ndays = _window_days(s, e)

    d['_lossw'] = d['Avg Loss (%)'] * d['Hours Measured']
    d['_lossday'] = d['Loss Hours'] > 0
    d = d.sort_values('Date')
    g = d.groupby('Site')
    out = g.agg(**{
        'Region': ('Region', 'last'), 'GBSC': ('GBSC', 'last'), 'RATs': ('RATs', 'last'),
        'FN/HUB Chain': ('FN/HUB Chain', 'last'),
        'Days Reported': ('Date', 'nunique'),
        'Hours Measured': ('Hours Measured', 'sum'),
        'Loss Hours': ('Loss Hours', 'sum'),
        'Loss Days': ('_lossday', 'sum'),
        'Minor Hours': ('Minor Hours', 'sum'),
        'No Response Hours': ('No Response Hours', 'sum'),
        'High Delay Hours': ('High Delay Hours', 'sum'),
        'Hub Event Hours': ('Hub Event Hours', 'sum'),
        'Loss Hours 2G': ('Loss Hours 2G', 'sum'),
        'Loss Hours 3G': ('Loss Hours 3G', 'sum'),
        '_lossw': ('_lossw', 'sum'),
        'Avg Loss 2G (%)': ('Avg Loss 2G (%)', 'mean'),
        'Avg Loss 3G (%)': ('Avg Loss 3G (%)', 'mean'),
        'Worst Hour Loss (%)': ('Worst Hour Loss (%)', 'max'),
        'Busy Avg Loss (%)': ('Busy Avg Loss (%)', 'mean'),
        'Quiet Avg Loss (%)': ('Quiet Avg Loss (%)', 'mean'),
        'Avg Delay (ms)': ('Avg Delay (ms)', 'mean'),
        'Max Delay (ms)': ('Max Delay (ms)', 'max'),
    }).reset_index()
    out['Avg Loss (%)'] = out['_lossw'] / out['Hours Measured'].replace(0, np.nan)
    out = out.drop(columns='_lossw')
    out['Persistence (%)'] = (out['Loss Hours'] / out['Hours Measured'].replace(0, np.nan) * 100).round(1)

    hub = d[d['Hub Event Hours'] > 0]
    if not hub.empty:
        top = hub.groupby(['Site', 'Hub Event Node'])['Hub Event Hours'].sum().reset_index()
        top = top.sort_values('Hub Event Hours', ascending=False).drop_duplicates('Site')
        out['Top Hub Node'] = out['Site'].map(top.set_index('Site')['Hub Event Node'])
    else:
        out['Top Hub Node'] = None

    # --- class
    if ndays >= 3:
        need_days = max(rules['recurrent_min_days'], int(np.ceil(rules['recurrent_day_share'] * ndays)))
        recurrent = out['Loss Days'] >= need_days
    else:
        recurrent = out['Loss Hours'] >= rules['recurrent_min_hours_short']
    not_rep = out['Hours Measured'].eq(0)
    chronic = (out['Hours Measured'] > 0) & (out['Loss Hours'] >= rules['chronic_share'] * out['Hours Measured']) \
        & (out['Loss Hours'] > 0)
    out['Class'] = np.select(
        [not_rep, chronic, recurrent, out['Loss Hours'] > 0, out['No Response Hours'] > 0],
        ['⚪ Not reporting', '🔴 Chronic', '🟠 Recurrent', '🟡 Sporadic', '⚫ Outage only'],
        default='🟢 Clean')

    # --- pattern
    busy, quiet = out['Busy Avg Loss (%)'].fillna(0), out['Quiet Avg Loss (%)'].fillna(0)
    traffic = (busy >= rules['pattern_ratio'] * quiet) & ((busy - quiet) >= rules['pattern_min_diff'])
    out['Pattern'] = np.select(
        [out['Loss Hours'].eq(0), out['Loss Hours'] < 3, traffic],
        ['—', 'Burst', 'Traffic-driven'], default='Constant')

    # --- affected RAT
    has2 = out['RATs'].fillna('').str.contains('2G')
    has3 = out['RATs'].fillna('').str.contains('3G')
    l2, l3 = out['Loss Hours 2G'] > 0, out['Loss Hours 3G'] > 0
    out['Affected RAT'] = np.select(
        [out['Loss Hours'].eq(0), l2 & l3, l2 & has3, l3 & has2, l2, l3],
        ['—', 'Both (2G+3G)', '2G only (3G clean)', '3G only (2G clean)', '2G', '3G'], default='—')

    # --- suspected cause
    bad_hours = out['Loss Hours'] + out['No Response Hours']
    hub_driven = (out['Hub Event Hours'] > 0) & (out['Hub Event Hours'] >= rules['hub_cause_share'] * bad_hours)
    rat_specific = out['Affected RAT'].isin(['2G only (3G clean)', '3G only (2G clean)']) & (out['Loss Hours'] >= 3)
    out['Suspected Cause'] = np.select(
        [out['Class'].eq('⚪ Not reporting'),
         hub_driven,
         out['Class'].eq('⚫ Outage only'),
         rat_specific & out['Affected RAT'].str.startswith('2G'),
         rat_specific,
         out['Pattern'].eq('Traffic-driven'),
         out['Pattern'].eq('Constant'),
         out['Pattern'].eq('Burst')],
        ['No ping response all period - link down, site off-air or ping not configured',
         'Shared transmission - hub/FN ' + out['Top Hub Node'].fillna('?').astype(str),
         'Link/site down hours (power or transmission cut) - check alarms',
         'ABIS only - 2G path/port/IP config (3G on same site is clean)',
         'IUB only - 3G path/port/IP config (2G on same site is clean)',
         'Congestion - backhaul capacity insufficient at busy hour',
         'Link quality fault - check MW alignment/interference, fiber or equipment',
         'Short event(s) - monitor'],
        default='—')

    out['_order'] = out['Class'].map({c: i for i, c in enumerate(CLASS_ORDER)})
    out = out.sort_values(['_order', 'Loss Hours', 'Avg Loss (%)', 'No Response Hours'],
                          ascending=[True, False, False, False]).drop(columns='_order')
    for c in [c for c in out.columns if c.endswith('(%)') or c.endswith('(ms)')]:
        out[c] = out[c].round(3)
    cols = ['Class', 'Site', 'Region', 'GBSC', 'Suspected Cause', 'Pattern', 'Affected RAT',
            'Loss Hours', 'Persistence (%)', 'Loss Days', 'Avg Loss (%)', 'Worst Hour Loss (%)',
            'Minor Hours', 'No Response Hours', 'High Delay Hours', 'Hub Event Hours', 'Top Hub Node',
            'Busy Avg Loss (%)', 'Quiet Avg Loss (%)', 'Avg Delay (ms)', 'Max Delay (ms)',
            'Loss Hours 2G', 'Loss Hours 3G', 'Avg Loss 2G (%)', 'Avg Loss 3G (%)',
            'Hours Measured', 'Days Reported', 'RATs', 'FN/HUB Chain']
    return out[cols].reset_index(drop=True)


def summarize_hub_events(hub_daily: pd.DataFrame, start, end, sites=None) -> pd.DataFrame:
    """Hub events over a window, one row per hub (optionally only hubs that
    touched any of `sites`)."""
    if hub_daily is None or hub_daily.empty:
        return pd.DataFrame()
    s, e = pd.Timestamp(start).strftime('%Y-%m-%d'), pd.Timestamp(end).strftime('%Y-%m-%d')
    h = hub_daily[(hub_daily['Date'] >= s) & (hub_daily['Date'] <= e)].copy()
    if sites:
        want = {x.upper() for x in sites}
        h = h[h['Affected Sites'].fillna('').apply(lambda v: bool(want & set(v.split(', '))))]
    if h.empty:
        return pd.DataFrame()
    g = h.groupby('Hub')
    out = g.agg(**{
        'Node Type': ('Node Type', 'last'), 'Region': ('Region', 'last'),
        'Event Days': ('Date', 'nunique'), 'Event Hours': ('Event Hours', 'sum'),
        'Loss Event Hours': ('Loss Event Hours', 'sum'),
        'No Response Event Hours': ('No Response Event Hours', 'sum'),
        'Peak Sites Affected': ('Peak Sites Affected', 'max'), 'Sites On Hub': ('Sites On Hub', 'max'),
        'Peak Share (%)': ('Peak Share (%)', 'max'),
        'Avg Loss During Events (%)': ('Avg Loss During Events (%)', 'mean'),
        'First Date': ('Date', 'min'), 'Last Date': ('Date', 'max'),
    }).reset_index()
    out['Affected Sites'] = out['Hub'].map(
        g['Affected Sites'].apply(lambda s: ', '.join(sorted({x for v in s.dropna() for x in v.split(', ') if x}))))
    out['Event Type'] = np.select(
        [out['No Response Event Hours'].eq(0), out['Loss Event Hours'].eq(0)],
        ['Packet loss', 'No response (link/power down)'], default='Mixed')
    out['Avg Loss During Events (%)'] = out['Avg Loss During Events (%)'].round(3)
    cols = ['Hub', 'Node Type', 'Region', 'Event Type', 'Event Days', 'Event Hours', 'Loss Event Hours',
            'No Response Event Hours', 'Peak Sites Affected', 'Sites On Hub', 'Peak Share (%)',
            'Avg Loss During Events (%)', 'First Date', 'Last Date', 'Affected Sites']
    return out[cols].sort_values(['Event Hours', 'Peak Sites Affected'], ascending=False).reset_index(drop=True)
