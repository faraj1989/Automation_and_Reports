# KPI Additions Request (4G + 2G)

Source of 4G KPI definitions: `KPI_20261004173819938.xml` (Huawei PRS KPI library export, 744 KPIs),
converted to `KPI_20261004173819938.xlsx` in the project root.

The pipeline (`backend/network_kpi_processor.py`, `backend/csv_history_manager.py`) keeps every column
present in the Huawei CSV export, and history CSVs accept new columns (older rows stay blank).
So these KPIs only need to be added to the PRS report templates; no pipeline change is needed to archive them.

Templates to update: `4G-NWBH`, `4G-NW_Daily`, 4G cell BH/hourly; 2G cell CSBH/hourly, `2G-NWBH`, `2G-NW_Daily`.

---

## 4G (LTE) — new KPIs

All four already exist in the PRS KPI library; add them to the report templates by KPI ID.

### 1. Success Rate of Outgoing Handover (Cell)
- **Library KPI:** `Intra-RAT HO Success Rate` — ID **1741400008**
- **Formula:**
  ```
  ( L.HHO.IntraeNB.IntraFreq.ExecSuccOut + L.HHO.IntereNB.IntraFreq.ExecSuccOut
  + L.HHO.IntraeNB.InterFreq.ExecSuccOut + L.HHO.IntereNB.InterFreq.ExecSuccOut )
  / ( L.HHO.IntraeNB.IntraFreq.ExecAttOut + L.HHO.IntereNB.IntraFreq.ExecAttOut
    + L.HHO.IntraeNB.InterFreq.ExecAttOut + L.HHO.IntereNB.InterFreq.ExecAttOut ) * 100
  ```
- All outgoing intra-LTE HOs (intra-freq + inter-freq, intra/inter-eNodeB). The current report has these
  two parts separately (`Intra-HO execute Success Rate`, `Inter-Freq Handover Success Rate (%)`).

### 2. CSFB Success Rate (Redirection)
- **Library KPI:** `CSFB Success Rate` (GTS) — ID **1740642182**
- **Formula:**
  ```
  ( L.RRCRedirection.E2W.CSFB + L.RRCRedirection.E2G.CSFB + L.IRATHO.E2W.CSFB.ExecSuccOut )
  / L.CSFB.PrepAtt * 100
  ```
- Alternative `CSFB Success Rate_Guo` — ID 1740642096 — also adds `L.IRATHO.E2G.CSFB.ExecSuccOut`.
- If a **redirection-only** rate is required, it is not in the library; create a user-defined KPI:
  `( L.RRCRedirection.E2W.CSFB + L.RRCRedirection.E2G.CSFB ) / L.CSFB.PrepAtt * 100`
- Current report has only `CSFB Preparation Success Rate` (= L.CSFB.PrepSucc / L.CSFB.PrepAtt).

### 3. VoLTE Handover Success Rate (QCI1)(%)
- **Library KPI:** `VoIP Handover Success Rate exclude InterFddTdd-ZM` — ID **1740642080**
- **Formula:**
  ```
  ( L.HHO.IntereNB.IntraFreq.ExecSuccOut.VoIP + L.HHO.IntereNB.InterFreq.ExecSuccOut.VoIP
  + L.HHO.IntraeNB.IntraFreq.ExecSuccOut.VoIP + L.HHO.IntraeNB.InterFreq.ExecSuccOut.VoIP )
  / ( L.HHO.IntereNB.IntraFreq.PrepAttOut.VoIP + L.HHO.IntereNB.InterFreq.PrepAttOut.VoIP
    + L.HHO.IntraeNB.IntraFreq.ExecAttOut.VoIP + L.HHO.IntraeNB.InterFreq.ExecAttOut.VoIP ) * 100
  ```
- Use ID 1740642079 (`include InterFddTdd-ZM`) instead only if the network has FDD<->TDD handovers.

### 4. RRC Establishment Success Rate
- **Library KPI:** `RRC Setup Success Rate (Service)` — ID **1740640163**
- **Formula:**
  ```
  ( L.RRC.ConnReq.Succ.Emc + .HighPri + .Mt + .MoData + .DelayTol + .MoVoiceCall )
  / ( L.RRC.ConnReq.Att.Emc + .HighPri + .Mt + .MoData + .DelayTol + .MoVoiceCall ) * 100
  ```
- Service-based (excludes mo-Signalling), and includes MoVoiceCall, which matters with VoLTE.
- Same formula without MoVoiceCall: `RRC connection setup success rate-ZM GTS`, ID 1740642105.
- Existing `RRC Setup Success Rate(%)` = L.RRC.ConnReq.Succ / L.RRC.ConnReq.Att (all causes incl. signalling).

---

## 2G (GSM) — counters + formulas

Formulas below were checked against two real cell rows (SURT018-2 / SURT018-3, 2026-07-20):
the computed value matches the exported KPI to 4 decimals where marked ✅.

### 1. Immediate Assignment Success Rate ✅
```
CA303J:Call Setup Indications (Circuit Service) / CA300J:Channel Requests (Circuit Service) * 100
```
Counters: **CA303J, CA300J** (already in the export).
Check: 2310 / 2378 = 97.1405 ✅ · 11268 / 11660 = 96.6381 ✅

### 2. Interference Band Proportion (4~5)
Library KPI `Interference Band Proportion (4~5)`, ID 1741941021 (from the 2G KPI XML):
```
( AS4207D + AS4207E + AS4208D + AS4208E )
/ ( AS4207A + AS4207B + AS4207C + AS4207D + AS4207E
  + AS4208A + AS4208B + AS4208C + AS4208D + AS4208E ) * 100
```
Counters: **AS4207A–E** (Mean Number of TCHFs in Interference Band 1–5) and
**AS4208A–E** (Mean Number of TCHHs in Interference Band 1–5). These aren't in the current export.

### 3. SDCCH Assignment Successful Rate - Faraj
```
K3003:Successful SDCCH Seizures / K3000:SDCCH Seizure Requests * 100
```
Counters: **K3003, K3000** (already in the export; K3001 "Failed SDCCH Seizures due to Busy SDCCH" also there).
Check: 2378 / 2378 = 100 ✅ · 11660 / 11660 = 100 ✅
Confirmed in the 2G KPI XML (ID 1740642228): no K3001 adjustment.

### 4. DL TBF Drop Rate - Faraj ✅
```
( A9118:Number of Downlink GPRS Intermit Transfers + A9318:Number of Downlink EGPRS Intermit Transfers )
/ ( A9102:Successful Downlink GPRS TBF Establishments + A9302:Successful Downlink EGPRS TBF Establishments ) * 100
```
Counters: **A9118, A9318, A9102, A9302** (already in the export).
Check: (7 + 83) / (14357 + 1633) = 0.5629 ✅ · (20 + 440) / (24323 + 7811) = 1.4315 ✅

### 5. UL TBF Establishment Successful Rate - Faraj
```
( A9002:Successful Uplink GPRS TBF Establishments + A9202:Successful Uplink EGPRS TBF Establishments )
/ ( A9001:Uplink GPRS TBF Establishment Attempts + A9201:Uplink EGPRS TBF Establishment Attempts ) * 100
```
Counters: A9002, A9202, A9001 are in the export. **A9201 (UL EGPRS TBF Establishment Attempts) is missing**
and must be added, otherwise the KPI cannot be recomputed from counters.
(Back-calculated from the exported rate, A9201 ≈ 1811 for SURT018-3 and ≈ 8410 for SURT018-2.)

### Related (already in export, for reference)
- Downlink TBF Establishment Success Rate = (A9102 + A9302) / (A9101 + A9301) * 100 ✅
- UL TBF Drop Rate - Faraj = (A9006 + A9007 + A9206 + A9207) / (A9002 + A9202) * 100 ✅
- TCH Assignment Successful Rate (Exclude HO) - Faraj = K3013A / K3010A * 100 ✅

---

## Counters to add to the 2G template (summary)
| Counter | Needed for |
|---|---|
| A9201: Uplink EGPRS TBF Establishment Attempts | UL TBF Establishment Successful Rate |
| AS4207A–E, AS4208A–E (TCHF/TCHH in Interference Band 1–5, 10 counters) | Interference Band Proportion (4~5) |

Everything else is already exported.

## After the new columns arrive
- Archive: automatic (new columns appear in `output/csv/*` history).
- Still to do in code: show the new KPIs in the dashboard / Word report and add thresholds to
  `config/kpi_thresholds.csv` (thresholds to be agreed by the RF team).
