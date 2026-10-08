# Libyana NPM — Project Documentation

Last compiled: 2026-09-13. This is a living reference — update it when a processor, scraper, or output path changes, since drift here is worse than no documentation at all.

## 1. What this project is

A network performance management (NPM) system for Libyana's mobile network (2G/3G/4G), combining:
- A **daily KPI reporting pipeline**: SFTP-delivered Huawei counter exports → cleaned/aggregated CSVs → an Excel/Word/email daily report + a persistent historical archive.
- A **live Streamlit dashboard** (`streamlit_dashboard.py`) reading the same archive, plus several independently-scraped data sources (NOC alarms, subscriber experience, device mix).
- A **NOC alarm pipeline**: Selenium scrapers logging into three Huawei OSS portals (MAE, NetEco, NCE) around the clock, feeding a native alarm-correlation engine that identifies which sites are down and why (power vs. transmission root cause).
- A **weekly/on-demand SmartCare CEM + Device Penetration pipeline**, a separate portal entirely.

## 2. Architecture at a glance

```
                    ┌─────────────────────────────────────────────┐
                    │              SFTP (Huawei MAE)               │
                    │  daily zips: site/network/cell/traffic/user  │
                    └───────────────────┬───────────────────────────┘
                                        │ sftp_downloader.py (paramiko)
                                        ▼
                    <local_root>/<date>/{zipped,unzipped}/*.csv
                                        │ csv_loader.read_csv_skip_metadata
                                        ▼
        ┌──────────────────────────────────────────────────────────────┐
        │  backend/*_processor.py  (site, network_kpi, cell_kpi,       │
        │  transmission_kpi, hourly_cell, traffic_kpi, user_kpi,       │
        │  interference, site_detail, device_penetration)              │
        └───────────────────────────┬────────────────────────────────┘
                                    │ csv_history_manager.py (dedupe/append/archive)
                                    ▼
                           output/csv/*.csv  (persistent history)
                                    │
                                    ▼
                    backend/report_generator.py (ReportGenerator)
                    │                                   │
                    ▼                                   ▼
        Excel/Word/email daily report          streamlit_dashboard.py (live)

        ─────────────────────────── separately ───────────────────────────

  scrapers/ (Selenium, this project)         scrapers/smartcare_cem_scraper.py
  MAE / NetEco / NCE, current + historical    scrapers/weekly_device_penetration_scraper.py
        │ scraper_watchdog.py keeps alive           │ run_daily_smartcare_reports.bat
        ▼                                           ▼
  Libyana_Data/Output/{Current,Historical}   Libyana_Data/Output/{SmartCare_Exports,
  _Alarms, NCE_Current/Historical_Alarms      Weekly_Device_Penetration_Exports}
        │ backend/noc_alarm_processor.py            │ backend/smartcare_cem_processor.py
        │ backend/topology_processor.py             │ backend/device_penetration_processor.py
        ▼                                           ▼
              streamlit_dashboard.py — 🚨 Alarms / 📱 CEM tabs
```

Two data roots exist and are **not the same tree**:
- `<local_root>` (from `ftp_config.json`, currently `C:\Users\user\Desktop\FTP files\Daily Report KPIs 2026 update`) — the SFTP-delivered KPI counter files.
- `C:\Users\user\Desktop\Libyana_Data\Output\...` (`NOC_ALARM_DATA_ROOT` / scraper `EXPORT_BASE_DIR`s) — everything the Selenium scrapers produce (alarms, CEM, device penetration). This is the same folder tree a now-retired sibling automation project used to write to; this project's own scrapers took over writing to it on 2026-09-10 so nothing downstream had to change paths.

## 3. Data ingestion

### 3.1 SFTP pipeline (KPI counters)

- **`backend/sftp_downloader.py`** — `SFTPDownloader.download_and_organize()` connects via `paramiko`, lists remote `.zip` files, dates each by **SFTP mtime** (not filename), downloads only if newer than what's local, unzips into `<local_root>/<date>/unzipped/`. Also runs local and remote duplicate cleanup (`cleanup_old_zipped_duplicates`, `cleanup_old_csv_duplicates`, `cleanup_old_remote_duplicates`) so neither side accumulates stale copies of the same report.
- **`backend/config_manager.py`** — loads/saves `ftp_config.json` (host/port/username/password/remote_path/local_root). Password is obfuscated with a single-byte XOR (key `0x5A`) + base64, **not real encryption** — treat `ftp_config.json` as a secret file regardless (it's gitignored).
- **`backend/csv_loader.py`** — `read_csv_skip_metadata()` is the universal reader every processor below calls: detects the delimiter, skips Huawei's metadata preamble (finds the first line starting with `Time`/`Date`), reads with `on_bad_lines='skip'` (falls back to latin-1), strips a trailing footer row, and maps Huawei's null tokens (`/0`, `NIL`, `NULL`, `N/A`, `--`) to `NaN` — **never to 0**, since `/0` means "zero attempts" (counter divide-by-zero) and `NIL` means "not collected," and either read as a literal 0 would misrepresent a KPI as a real failure.

### 3.2 Scraper pipeline (NOC alarms + SmartCare)

All scrapers live in `scrapers/`, use `project_config.py` (root-level: `env_str`/`env_int`/`env_path_str`/`load_env_file`, `LoginFailedError`, `mark_login_failed`/`is_login_failed_today`, `keep_only_latest_export`) and read credentials/URLs from the root `.env` (gitignored).

| Scraper | Portal | Cadence | Output |
|---|---|---|---|
| `mae_scraper.py` | Huawei iMaster MAE | continuous, ~5 min | `Current_Alarms/<date>/CurrentAlarms_MAE_<ts>.csv` (latest only) |
| `mae_historical_alarms_scraper.py` | MAE | continuous, ~5 min | `Historical_Alarms/<date>/HistoricalAlarms_MAE_<ts>.zip` (accumulates) |
| `neteco_all_alarms_scraper_launcher.py` → `neteco_continuous all alrams.py` | NetEco (power/environment) | continuous, ~5 min | `Current_Alarms/<date>/NetEco_All_Current_Alarm_<ts>.csv` (latest only) |
| `neteco_historical_alarms_scraper.py` | NetEco | continuous, ~7 min | `Historical_Alarms/<date>/NetEco_Historical_Alarm_<ts>.csv` (accumulates) |
| `nce_active_alarms_scraper.py` | NCE (transmission) | continuous, ~5 min | `NCE_Current_Alarms/<date>/CurrentAlarms_NCE_<ts>.csv` (latest only) |
| `nce_historical_alarms_scraper.py` | NCE | continuous, ~5 min | `NCE_Historical_Alarms/<date>/HistoricalAlarms_NCE_<ts>.csv` (accumulates) |
| `smartcare_cem_scraper.py` | SmartCare (DPI/CEM) | on-demand, meant daily | `SmartCare_Exports/Comprehensive_Analysis_<ts>.zip`, folded by `reports/run_smartcare_analysis_task.py` into `Processed_Analysis/Comprehensive_Analysis_Historical.xlsx` |
| `weekly_device_penetration_scraper.py` | SmartCare (same login, different dashboard) | on-demand, meant daily | `Weekly_Device_Penetration_Exports/Device_Penetration_Rate_<ts>.zip`, folded into `Weekly_Device_Penetration_Historical.xlsx` |

**`scrapers/scraper_watchdog.py`** is the supervisor for the 6 continuous alarm scrapers: checks every 60s whether each is alive (psutil, cmdline match), relaunches any that died, and — critically — **skips relaunching one whose login failed today** (`is_login_failed_today`, keyed to a SHA-256 fingerprint of the password that failed, so fixing `.env` unblocks it immediately rather than waiting for a new calendar day). It refuses to run a second copy of itself. It is currently started manually (see §7 — the corresponding scheduled task was never registered).

**Important operational rule**: never run one of the 6 continuous scraper scripts directly while the watchdog is active — see §8, finding 1.

### 3.3 Native alarm correlation (`backend/noc_alarm_processor.py`)

This module (rewritten 2026-09-10 to be fully self-contained — it originally read a separate automation suite's pre-computed rollups, which stopped being reliable once scraping moved here) computes "which sites are down and why" directly from the raw exports above:

- **`build_live_disconnected_sites()`** — right now, from each source's latest current-alarm export. Returns the rich per-site table: `Site Name, MO Name, Alarm Name, Last Occurred, Down Hours, Power Reason (NOC), Mains Failure Time, Transmission Reason (NOC), NCE Last Occurred` — deliberately matching the column layout of the original sibling suite's `Live_NOC_Report.xlsx` that NOC engineers were already used to.
- **`build_daily_noc_alarm_report(target_date)`** — the same table for one historical calendar day, from the accumulating historical exports, with per-site summed `Down Hours` for that day.
- **`build_historical_insights_native(lookback_days=14)`** — chronic-offender / site-downtime / daily-trend / category rollups, computed by running the daily logic across a rolling window.
- Shared logic (`_attach_alarm_reasons`, `_power_reason_and_time`, `_transmission_reason_and_time`, `_is_power_reason`, `_nce_reason_for_site`) ensures all three views agree on what counts as a Power vs. Transmission cause. `_nce_reason_for_site` and `_power_reason_and_time`/`_transmission_reason_and_time` are **hub-aware**: if a site itself shows no transmission alarm, they walk up `backend/topology_processor.py`'s `build_site_ancestor_map()` (parsed from `config/site_topology.csv`) to check whether its upstream FN/HUB is down instead — a single failed regional hub can otherwise leave 8+ downstream sites wrongly bucketed "Investigating."
- `SITE_DOWN_ALARM_NAME = "NE Is Disconnected"` is the single definitive "site is down" signal (MAE). `POWER_ALARM_TERMS = ("Mains Failure", "BLVD", "LLVD")` defines what counts as a real power cause — anything else (e.g. NetEco's own ambiguous "Communication Between NMS And NE Is Abnormal") reads as "Check TX/Link (No Power Alarm)," not Power.
- `_extract_site_code()` normalizes Huawei's per-technology device-naming decoration: 3G NodeB names get a `U` prefix, 4G eNodeB names get an `L` prefix, over the real bare site code (confirmed against `output/csv/SiteDetail.csv`) — stripped only when the decorating key is actually present, since a handful of real sites (`UMSF001`, `LAHB001`, ...) legitimately start with U/L on their own.

### 3.4 CEM and Device Penetration readers

- **`backend/smartcare_cem_processor.py`** reads `Processed_Analysis/Comprehensive_Analysis_Historical.xlsx` (3 sheets: `Top100 per day`, `Top10 Application per week`, `Metrics` — TCP connection quality/traffic volume trend). Feeds the 📱 CEM tab's "App Traffic & Quality" sub-tab.
- **`backend/device_penetration_processor.py`** reads `Weekly_Device_Penetration_Exports/Weekly_Device_Penetration_Historical.xlsx`. Feeds the 📱 CEM tab's "Device Penetration" sub-tab.
- Both are "stale after N days" aware (`*_STALE_DAYS`) and degrade to an info message rather than erroring if the file is missing or old.

## 4. Backend KPI processors (SFTP-fed)

Each reads specific file-pattern(s) from `<local_root>/<date>/unzipped/`, returns cleaned DataFrame(s), and is written to `output/csv/` by `csv_history_manager.py`.

| Processor | Raw file pattern(s) | Key columns | Output sheet(s) |
|---|---|---|---|
| `site_processor.py` | `* (2G).csv`, `* (3G).csv`, `* (4G).csv` | `Site Name`/`DL frequency`, `NodeB Name`/`Band Indicator`, `eNodeB Name`/`Frequency band`/`EARFCN` | `SiteSummary.csv` (band/overlap counts) |
| `site_detail_processor.py` | same 2G/3G/4G files | + `Cell Name`, `Downlink bandwidth` | `SiteDetail.csv` (per-site master: bands, scenario, RAT, sectors) |
| `network_kpi_processor.py` | `* (2G/3G/4G-NWBH).csv`, `* (2G/3G/4G-NW_Daily).csv`, `*Gi Interface Traffic*.csv` | whole-network busy-hour/daily KPIs | `2G_NWBH.csv`, `2G_NW_Daily.csv`, `3G_NWBH.csv`, `3G_NW_Daily.csv`, `4G_NWBH.csv`, `4G_NW_Daily.csv`, `Gi_Interface_Traffic.csv` |
| `cell_kpi_processor.py` | `* (2G cell-CSBH).csv`, `* (3G-cells -CSBH).csv`, `* (4G cell-BH).csv` | per-cell busy-hour KPIs | `2G_Cell_CSBH.csv`, `3G_Cell_CSBH.csv`, `4G_Cell_BH.csv` |
| `hourly_cell_processor.py` | `*all last hours all cells level*.csv` (3 files inside, identified by header column `GBSC`/`RNC`/`eNodeB Name`, not filename) | full hourly grain, no aggregation | `2G_Cell_Hourly.csv`, `3G_Cell_Hourly.csv`, `4G_Cell_Hourly.csv` (90-day retention, excluded from the combined Excel export) |
| `transmission_kpi_processor.py` | `*BSC6900*PACKETLOSS*.csv` (IUB/ABIS backhaul, ~260k rows/day) | `T7816`/`T7817` packet loss, `T7812`/`T7813` delay | Via `packet_loss_engine.py`: `Transmission_KPIs.csv` (1 row/link/day, key Date+GBSC+ID - IDs repeat across BSCs), `Packet_Loss_Site_Daily.csv` (1 row/site/day: loss/minor/no-response/high-delay/hub-event hour counts - permanent), `Packet_Loss_Hub_Events.csv` (FN/HUB shared-path events per day), `Packet_Loss_Site_Hourly.csv` (31-day rolling drill-down). Complete past days only; each run replaces the 7 dates it covers. Rebuild all history: `python -m backend.transmission_kpi_processor "<FTP root>" --backfill` |
| `traffic_kpi_processor.py` | `* (PS Traffic 2G/3G/4G).csv` | `PS Traffic(GB)`, `CS Traffic`, `VoLTE Traffic Volume (Erl)` | `Traffic_2G/3G/4G.csv` (per-site), `Traffic_Network_2G/3G/4G.csv` (whole-network) |
| `user_kpi_processor.py` | `* (CS Roaming users).csv`, `* (MSC Server KPI-CS Subscribers+total).csv`, `* (PS Roaming users).csv`, `* (PS users (2G-3G-4G)).csv`, `* (VoLTE users).csv` | daily MAX per KPI | `User_CS_Roaming.csv`, `User_CS_Subscribers.csv`, `User_PS_Roaming.csv`, `User_PS_Subscribers.csv`, `User_VoLTE.csv`, consolidated `User_Summary.csv` |
| `interference_processor.py` | `*interference*PRB*utilization*Automation*.csv` (weekly Huawei upload, 3 files inside identified by header column) | `Interference Band Proportion (4~5)(%)` (2G), `VS.MeanRTWP` (3G), `L.UL.Interference.Avg(dBm)` (4G) | `2G_Interference.csv`, `3G_Interference_Hourly.csv` + `3G_Interference_Daily.csv` rollup, `4G_Interference_Hourly.csv` + `4G_Interference_Daily.csv` rollup |

Special filters worth knowing: PS-Roaming is filtered to `Mobile country code==606, Mobile network code==1`; CS-Roaming is filtered to partner code `218919121253` (Almadar) — both hardcoded in `user_kpi_processor.py`.

## 5. Central orchestration

- **`backend/csv_history_manager.py`** — every processor's output funnels through here. `_append_with_dup_check` is the generic dedupe-and-append primitive; `update_hourly_cell_kpis`/`update_interference_kpis` use vectorized concat+`drop_duplicates` for million-row scale instead. `export_to_excel()` combines everything (except `_Cell_Hourly`/`_Interference_Hourly` detail sheets) into `output/Historical_Network_Data.xlsx`. Retries file writes up to 4× on `PermissionError` (Windows AV/indexer lock contention).
- **`backend/health_checker.py`** — evaluates any KPI DataFrame against `config/kpi_thresholds.csv`, producing pass/fail + a weighted health score + a vectorized worst-cells ranking (`sum(fail weight×10) + violations×5`). Used identically by the daily report and the dashboard so they never disagree on a KPI's status.
- **`backend/report_generator.py`** (`ReportGenerator`, ~1800 lines, ~45 methods) — the glue layer: builds scorecards, worst cells, traffic, site health/topology, alarm overview, CEM/device-penetration overview, packet loss, site inventory, 14-day trends, and generates the Excel/Word/email daily report. Both `scheduler.py` and `streamlit_dashboard.py` instantiate the same class so the live dashboard and the emailed report always agree.
- **`backend/topology_processor.py`** — regenerates `config/site_topology.csv` from the RF team's `FN-HUB Info*.xlsx` reference (`build_site_topology_csv`), and (added 2026-09-10) `build_site_ancestor_map()` for the hub-aware alarm correlation in §3.3.

## 6. Automation schedule

`scheduler.py` is a one-shot, argparse CLI (no internal loop) — Windows Task Scheduler invokes it via `run_scheduler.bat`, which just forwards `%*`.

| Task | Trigger | Command | What it does |
|---|---|---|---|
| Libyana NPM Daily Scheduler | daily 06:30 | `run_scheduler.bat` (no flag) | Full pipeline for **yesterday**: ensure dashboard alive → SFTP download (yesterday+1, since MAE uploads the next morning) → all 8 processors → archive NOC alarm summary → export Excel → generate Excel/Word/email report |
| Libyana NPM Hourly Cells | every 6h | `run_scheduler.bat --hourly-cells` | Rolling-window download + `hourly_cell_processor` only |
| Libyana NPM Weekly Interference | Sun 07:00 | `run_scheduler.bat --interference-weekly` | Rolling-window download + `interference_processor` → cleans up local raw interference files after archiving (~650MB/week unzipped) |
| Libyana NPM Dashboard | keep-alive | — | Keeps the persistent Streamlit process running |
| **Libyana NPM Scraper Watchdog** | *(never registered — see §8)* | `run_scraper_watchdog.bat` | Supervises the 6 continuous alarm scrapers |
| **Libyana NPM Daily SmartCare Reports** | *(never registered — see §8)* | `run_daily_smartcare_reports.bat` | CEM scrape → CEM analysis fold-in → Device Penetration scrape |
| **Firewall: Libyana NPM Dashboard (8501)** | *(never registered — see §8)* | inbound allow, TCP 8501 | LAN access to the dashboard from other devices |

`_send_email()` in `scheduler.py` is currently a placeholder — it logs report paths but does not send SMTP mail. `--auto-send` has no effect until that's implemented.

## 7. Dashboard (`streamlit_dashboard.py`)

Sidebar: report-date selector, auto-computed previous-date for comparison, "Refresh data" (clears `st.cache_data`), then the section radio. All data loads go through one shared `ReportGenerator` instance, cached (`@st.cache_resource` / `@st.cache_data(ttl=...)`).

Top-level sections, in order:
1. **📊 Overview** — Site Summary, Executive Summary
2. **📡 KPIs & Performance** — Scorecards, Worst Cells, Traffic & Capacity, 14-Day Trend, Data Freshness (each split further by 2G/3G/4G)
3. **🏗️ Sites & Infrastructure** — Site Health & Topology, Site Inventory, Site Detail, EPT
4. **📶 Packet Loss** — Network view (Day / 7 / 30 days / custom: class tiles, affected-site ranking by loss hours with suspected cause + hour-by-hour drill-down heatmap, FN/HUB shared-path events, network trend, core links & not-reporting) and Special report (pick sites / whole FN-HUB node / region + date range + label). Excel and Word export on both. Logic in `backend/packet_loss_engine.py`, thresholds in `config/packet_loss_rules.csv`
5. **🚨 Alarms** — live "Currently Disconnected Sites" + metric tiles, historical rollups (Chronic Offenders / Site Downtime / Daily Trend / Category Rollup), and a date-range-driven "Daily NOC Alarm Analysis": per-site Status (Cleared / Not Cleared / Unknown - historical exports only hold cleared alarms, so still-open ones come from the live feed, and a stale live snapshot gives Unknown), Outage Class (Long-Term Outage = an outage of `LONG_TERM_OUTAGE_DAYS`, 7, or more days, optionally excluded from KPIs), per-event list, daily summary, active outages, data-quality per day, combinable filters, and a formatted 5-sheet Excel workbook + Word export
6. **📱 CEM** — App Traffic & Quality, Device Penetration
7. **🔎 Investigate** — Cell Explorer, Special Reports
8. **📋 HQ Reports** — NQ Data Collection Template (own sub-tabs per sheet), Traffic & Availability (Tripoli HQ)
9. **📧 Reports** — Report & Export

## 8. Configuration reference

| File | Rows/shape | Purpose |
|---|---|---|
| `config/kpi_thresholds.csv` | 43 KPI rows × 12 cols (`Technology, KPI_Name, Column_Name, Source_Sheet, Aggregation, Threshold, Operator, Weight, Dimension, Severity, Definition, Effect`) | Single source of truth for every KPI's pass/fail threshold and health-score weight |
| `config/site_topology.csv` | 1785 rows × 5 cols (`Region, Node_Name, Node_Type, Connected_Site, Remark`) | FN/HUB transmission topology, regenerated from the RF team's Excel reference |
| `config/Golden 5 KPIs for Worst Cells.csv` | 3 rows (GSM/UMTS/LTE) × 5 KPI-category columns | Which KPI represents each of Accessibility/Retainability/Mobility/Congestion/Integrity, per technology |
| `.env` (root, gitignored) | — | Scraper credentials/URLs (MAE/NetEco/NCE/SmartCare), `DATA_ROOT`, per-scraper timing/timeout overrides |
| `ftp_config.json` (root, gitignored) | — | SFTP host/credentials (password XOR-obfuscated, not real encryption) |
| `requirements.txt` | — | pandas, numpy, openpyxl, paramiko, streamlit, plotly, schedule, xlwings, python-dotenv, python-docx, matplotlib, selenium, webdriver-manager, psutil |

## 9. Known issues / recommended next steps

1. **Scrapers have no self-duplicate guard.** The watchdog refuses to run a second copy of *itself*, but the 6 scraper scripts don't check whether another copy of *themselves* is already running. Manually launching one (e.g. via VS Code's "Run Python File") while the watchdog already manages it creates a second concurrent login into the same MAE/NetEco/NCE account — confirmed twice live (2026-09-10 03:00 and again 3:42 PM, the second time going undetected for 3 full days — see §10). **Recommendation**: add an `is_only_instance()`-style check to each scraper's own startup, refusing to proceed if another instance is already running.
2. **No retention/cleanup for accumulating historical alarm exports.** Unlike the current-alarm scrapers (`keep_only_latest_export`), the 3 historical scrapers keep every export forever. Confirmed: 2.3GB in the first 3 days alone (~0.75GB/day). **Recommendation**: a periodic job (or a change to `scraper_watchdog.py`) that deletes historical export files older than N days, matching the retention concept the old sibling project used but never actually enforced here.
3. **Three Task Scheduler / firewall entries were proposed but never registered** (see §6, §10) — I don't have the privilege to self-elevate and register them; each needs a one-time run of the given command in an **elevated** PowerShell.
4. **`config_manager.py`'s password "encryption" is a single-byte XOR** — obfuscation only. `ftp_config.json` should be handled as if it contained a plaintext password (it's gitignored, which is the real protection).
5. **`_send_email()` is an unimplemented placeholder** — `--auto-send` currently does nothing beyond logging.

## 10. Operational audit — last 3 days (2026-09-10 → 2026-09-13)

Performed 2026-09-13. Findings:

- **MAE/NetEco/NCE alarm scrapers**: the watchdog-managed set has been healthy and continuously producing data for all 3 days — zero login-failure markers, at most one restart per scraper (MAE Historical, once, on 2026-09-10 — the false-positive rate-limit/credentials bug fixed that same day). Historical export counts per day (MAE/NetEco/NCE): 09-10 318/269/360, 09-11 487/412/378, 09-12 525/412/276, 09-13 (partial) 199/156/105 — no gap days.
- **🔴 Found and fixed live**: a complete second set of all 6 scrapers had been running since **2026-09-10 15:42** (confirmed via process creation timestamps) — 3 days of doubled logins into every MAE/NetEco/NCE account, undetected until this audit. Cleaned up: 14 duplicate processes + 7 orphaned `chromedriver.exe` instances. Exactly 6 processes / 6 browsers remain, matching the 6 intended scrapers.
- **🔴 CEM + Device Penetration have been stale for 3 days.** Last successful collection: **2026-09-10 15:46** — the same timestamp as the duplicate-scraper incident above, strongly suggesting both came from the same manual re-run that afternoon. The `run_daily_smartcare_reports.bat` Task Scheduler entry was never registered, so nothing has collected fresh data since.
- **Disk**: `Libyana_Data/Output` is 2.5GB total after 3 days; `Historical_Alarms` (1.1GB) + `NCE_Historical_Alarms` (1.2GB) already account for 92% of it. See recommendation §9.2.
- **Dashboard**: running and healthy on port 8501; LAN access (`http://192.168.31.201:8501`) may still be blocked pending the firewall rule in §6/§9.3.
