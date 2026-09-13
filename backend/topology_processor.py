#!/usr/bin/env python3
"""
Libyana NPM - FN/HUB Topology Processor
Converts the RF-team-maintained "FN-HUB Info-... FN Sites List.xlsx" (one
sheet per region, plus a couple of combined index sheets) into the flat
Region/Node_Name/Node_Type/Connected_Site/Remark shape that
ReportGenerator.build_topology_summary() already reads from
config/site_topology.csv.

Sheet layout (RF team's spreadsheet convention, not this pipeline's):
each Fiber Node/HUB "block" starts on a row with Region/Site name/Type/
Count filled in, then continuation rows for each site impacted by that
node leave those columns blank (Excel-style "looks merged" via blank
cells, not actual merged-cell ranges) and only fill in that row's
'List of sites impacted' value. Parsed by column NAME (not position) with
a per-column forward-fill, since column order/count differs between the
per-region sheets, the "*Related Sites" index sheets, and DARN (which adds
an extra 'Connected Sites (Datacom.)' column no other sheet has).

Several sheets overlap (e.g. "FN-Related Sites" appears to be a combined
index of the 11 per-region sheets, "South Related Sites" largely duplicates
"South") - rather than guess which is authoritative, every sheet is parsed
and the result de-duplicated on (Node_Name, Connected_Site), so nothing is
missed if two sheets disagree slightly.
"""

import os
import logging
import pandas as pd
import openpyxl

logger = logging.getLogger(__name__)

TOPOLOGY_XLSX_GLOB = 'FN-HUB Info*.xlsx'
CONFIG_DIR = 'config'
OUTPUT_CSV = os.path.join(CONFIG_DIR, 'site_topology.csv')

# The node's own row identifies it as an impacted site too (a site is
# "connected to" its own FN/HUB) - kept, matching the original hand-built
# CSV's convention (e.g. SHATGP -> SHATGP as its own first Connected_Site row).
NODE_COL = 'Site name'
TYPE_COL = 'Terminal or Hub or Fiber node'
COUNT_COL = 'No. of Sites impacted'
CONNECTED_COL = 'List of sites impacted'
REMARK_COL = 'Remark'
REGION_COL = 'Region'

# Columns that carry forward from the block's first row down through its
# continuation rows (Excel's "looks merged" blank-cell convention).
FORWARD_FILL_COLS = [REGION_COL, NODE_COL, TYPE_COL, COUNT_COL]


def find_topology_xlsx(config_dir=CONFIG_DIR):
    import glob
    matches = glob.glob(os.path.join(config_dir, TOPOLOGY_XLSX_GLOB))
    return matches[0] if matches else None


def _clean_header(name):
    return str(name).strip() if name else name


def _parse_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [_clean_header(h) for h in rows[0]]
    # A duplicate header name (e.g. "South"'s two "List of sites impacted"
    # columns) collapses to its last occurrence - both hold the same value
    # on every observed row, so nothing is lost.
    col_idx = {name: i for i, name in enumerate(header) if name}
    header_values_lower = {str(h).strip().lower() for h in header if h}

    def get(row, col_name):
        i = col_idx.get(col_name)
        return row[i] if i is not None and i < len(row) else None

    records = []
    carry = {c: None for c in FORWARD_FILL_COLS}
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        # Some sheets (e.g. "FN-Related Sites") were built by pasting
        # several exports together and repeat the header row mid-sheet at
        # each paste boundary - skip any row whose every non-blank cell is
        # itself literally a column header, or it leaks in as a bogus
        # "Region" node.
        row_values_lower = {str(v).strip().lower() for v in row if v is not None and str(v).strip()}
        if row_values_lower and row_values_lower.issubset(header_values_lower):
            continue
        rec = {}
        for c in FORWARD_FILL_COLS:
            val = get(row, c)
            if val is not None and str(val).strip():
                carry[c] = val
            rec[c] = carry[c]
        rec[CONNECTED_COL] = get(row, CONNECTED_COL)
        rec[REMARK_COL] = get(row, REMARK_COL)
        if rec[NODE_COL] and rec[CONNECTED_COL]:
            records.append(rec)
    return records


def build_site_topology_csv(xlsx_path=None, output_csv=OUTPUT_CSV, log_callback=None):
    """Regenerate config/site_topology.csv from the FN-HUB Excel reference.
    Returns the resulting DataFrame (empty if the xlsx couldn't be found/read)."""

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            logger.info(msg)

    xlsx_path = xlsx_path or find_topology_xlsx()
    if not xlsx_path or not os.path.exists(xlsx_path):
        log(f"⚠️ Topology Excel reference not found in {CONFIG_DIR}/")
        return pd.DataFrame()

    log(f"📖 Reading topology reference: {os.path.basename(xlsx_path)}")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)

    all_records = []
    for sheet_name in wb.sheetnames:
        recs = _parse_sheet(wb[sheet_name])
        log(f"   {sheet_name}: {len(recs)} relationship rows")
        all_records.extend(recs)

    if not all_records:
        log("⚠️ No relationship rows parsed from the topology Excel file")
        return pd.DataFrame()

    df = pd.DataFrame(all_records).rename(columns={
        REGION_COL: 'Region', NODE_COL: 'Node_Name', TYPE_COL: 'Node_Type',
        CONNECTED_COL: 'Connected_Site', REMARK_COL: 'Remark',
    })
    df = df[['Region', 'Node_Name', 'Node_Type', 'Connected_Site', 'Remark']]
    df['Node_Name'] = df['Node_Name'].astype(str).str.strip()
    df['Connected_Site'] = df['Connected_Site'].astype(str).str.strip()
    # "FN-Related Sites" (the combined index sheet) suffixes node names with
    # their own type, e.g. "BGZ005(FN)" for the same physical node the
    # per-region "Benghazi" sheet calls plain "BGZ005" - strip it so both
    # collapse to one canonical node instead of showing up as two (measured:
    # 407 raw node names -> 350 after stripping, i.e. 57 pure duplicates).
    df['Node_Name'] = df['Node_Name'].str.replace(r'\s*\((FN|HUB)\)\s*$', '', regex=True)

    before = len(df)
    df = df.drop_duplicates(subset=['Node_Name', 'Connected_Site']).reset_index(drop=True)
    log(f"✅ Parsed {before} rows, {len(df)} unique (Node, Connected Site) relationships, "
        f"{df['Node_Name'].nunique()} FN/HUB nodes")

    os.makedirs(CONFIG_DIR, exist_ok=True)
    df.to_csv(output_csv, index=False)
    log(f"💾 Wrote {output_csv}")
    return df


def build_site_ancestor_map(csv_path=OUTPUT_CSV) -> dict:
    """site -> set of every FN/HUB node upstream of it, walking the full
    chain rather than one hop. The topology isn't always a flat site->hub
    mapping: a HUB's own Connected_Site rows can in turn list it under a
    bigger FN (e.g. in the Kufra region, KUFR003's direct parent is the HUB
    "KUFR001", which is itself one of the sites the FN "KUFRAGP" lists as
    Connected_Site - so KUFR003's full ancestor set is {KUFR001, KUFRAGP}).
    A node's self-reference row (Connected_Site == Node_Name, see this
    module's docstring) isn't a real upstream edge and is skipped. Empty
    dict if config/site_topology.csv doesn't exist yet.
    Cycle-guarded since the source spreadsheet is hand-maintained by the RF
    team, not generated - a bad edit could otherwise create an infinite
    parent loop."""
    if not os.path.exists(csv_path):
        return {}
    df = pd.read_csv(csv_path)
    df = df[df['Connected_Site'] != df['Node_Name']]

    direct_parents: dict = {}
    for leaf, node in zip(df['Connected_Site'], df['Node_Name']):
        direct_parents.setdefault(leaf, set()).add(node)

    ancestors_cache: dict = {}

    def ancestors_of(site, visiting=frozenset()):
        if site in ancestors_cache:
            return ancestors_cache[site]
        if site in visiting:
            return set()
        result = set()
        for parent in direct_parents.get(site, ()):
            result.add(parent)
            result |= ancestors_of(parent, visiting | {site})
        ancestors_cache[site] = result
        return result

    return {site: ancestors_of(site) for site in direct_parents}


# ---------------------------- Test ----------------------------
if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO)
    df = build_site_topology_csv()
    print(df.head(20).to_string())
    print(f"\n{len(df)} total rows")
