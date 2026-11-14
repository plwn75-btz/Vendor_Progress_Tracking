# Technical Specification (SPEC.md) — Weekly Project Management & Vendor Package Monitoring Dashboard (ASK & Z1F)

## 1. Problem Statement & Operational Context
Project Management and Engineering/Procurement Expediting teams for the **ASK** (Aung Sinkha Development Project Phase 1A — EPC-01) and **Z1F** (Zawtika Development Project Phase 1F — Topside & Jacket WP1) offshore projects track vendor package execution across distinct sets of dynamic weekly documents:
1. **Project Master Schedule (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf` Rev B1 & `MM-ZTK-1F-GEN-PJM-SCD-0002.pdf` Rev C1):** Master EPC Primavera P6 schedules containing 2,522 (ASK) and 4,110 (Z1F) activities that establish contractual Required On Site (ROS) delivery gates, key engineering approval milestones, yard topside erection windows, and critical path activities (`Total Float = 0` or tight float).
2. **Procurement Status Report / Procurement Plan (`Attachment 1 - Procurement Plan-Z1F_*.xlsx` & `Attachment 2-Procurement Plan-ASK-*.xlsx`):** Multi-sheet procurement tracking registers detailing PO award dates, ex-works delivery forecast, and companion ROS dates across all core packages and minor equipment items.
3. **Weekly Progress Reports (`WPR/*.pdf`, `WPR/*.docx`):** Detailed weekly or bi-weekly reports submitted by package vendors detailing milestone progress weights, planned vs. actual progress S-curves, sub-vendor material orders (Attachment 4), and fabrication Gantt schedules (Attachment 3).
4. **Minutes of Meeting (`MOM/*.pdf`):** Regular progress and technical alignment meetings containing contractual action items, action parties, target completion dates, status (`OPEN`, `CLOSED`, `INFO`), and sub-order procurement notes.
5. **Vendor Document Register Lists (`VDRL/*.xlsx`):** Tracking registers monitoring vendor document submittals (`SUBMITTED`, `NOT YET`, `OVERDUE`, `CLOSED`), Key vs. Non-Key documents (`Y`/`N`), and production-impact flags.
6. **Official Client Review & Return Log (`ASK Vendor Document Log Report*.xls/.xlsx`):** Contractual log exported from D2 tracking Client (PTTEPI) and Contractor (TJE/GCME) review turnarounds (`In Review`, `Submitted`, `In Approval`, `APPR`, `AWC`, `APPR-C`, `REC`, reviewer target vs. actual finish dates, and transmittals).
7. **Kick-Off Meeting Presentations (`KOM/*.pdf`, `KOM/*.pptx`):** Vendor kick-off decks containing manufacturer corporate entity, vendor code, factory location, KOM date, and contact rosters.

### Architectural Constraint (Web-Based Deployment & Asymmetric Environments)
The user operates from two distinct environments:
- **Local Working Computer:** Contains the complete raw multi-gigabyte document repository (`KOM/`, `WPR/`, `MOM/`, `VDRL/`, `SCD` master PDFs). Due to corporate network restrictions, Git CLI (`git add`, `git commit`, `git push`) is unavailable, requiring manual uploads through the GitHub web interface.
- **Admin Computer / Cloud Hosting (GitHub & Render.com):** GitHub Web UI enforces a strict 25 MB per-file and per-batch upload limit. Uploading thousands of weekly PDFs (~8,000 files, ~16 GB) causes browser timeouts, repository bloat, and Render container memory exhaustion on startup.
- **Solution:** A two-tier decoupled architecture:
  1. **Local Extraction & Synchronization Engine:** Raw PDFs are parsed locally. Essential structured data is extracted and non-destructively upserted into lightweight project-specific Excel workbooks: [`Export Data File_ASK.xlsx`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/Export%20Data%20File_ASK.xlsx) (~145 KB) and [`Export Data File_Z1F.xlsx`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/Export%20Data%20File_Z1F.xlsx) (~78 KB).
  2. **Pure Cloud / Web Deployment:** The GitHub repository and Render web server read the lightweight Excel export files directly, delivering 100% functionality with zero dependency on heavy raw PDFs.

---

## 2. Core Functional Specifications

### 2.1. Switchable Dual-Project Workspace (`ASK` | `Z1F`)
- The dashboard supports independent, switchable workspaces for **ASK** and **Z1F**, defaulting to **ASK** upon initial load.
- All executive KPIs, package inventories, VDRL counts, S-curves, Look-Ahead items, and MOM action items are strictly filtered by the selected project.
- Top bar contains an interactive project switch button that updates the URL state (`?project=ASK` or `?project=Z1F`) and re-renders the dashboard instantly.

### 2.2. Governing ROS & Master Schedule Resumption
- **Dual ROS Alignment:** Tracks vendor delivery against official Project Master Schedule ROS (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf` for ASK, `MM-ZTK-1F-GEN-PJM-SCD-0002.pdf` for Z1F) as **Priority #1**, side-by-side with Procurement Plan ROS (`Attachment 2-Procurement Plan-ASK-*.xlsx` / `Attachment 1 - Procurement Plan-Z1F_*.xlsx`) as **Priority #2**.
- **Clash Detection & Early Warning:** Detects delivery clashes against Yard Topside Erection windows and highlights Critical Path (`Total Float = 0` or tight float) packages (e.g. Pedestal Crane `PDC`, Diesel Transfer Pump `DTP`, Control Valve `CV`, Actuated Choke Valve `ACV`, Steel Plate `PLT`).
- **One-Picture Executive Summary:**
  - **4-Tier Classification:** Groups packages into `ALREADY DELAYED`, `HIGH RISK`, `LOW RISK`, and `VERY GOOD / HEALTHY`.
  - **Smart Mitigation Action Matrix:** Automatically prioritizes delayed and high-risk packages with actionable expediting focus areas.

### 2.3. Unified 7-Day & 14-Day Look-Ahead Hub
- Combines Master Schedule EPC level activities (Engineering deliverables, Procurement PO releases, Yard fabrication/erection dates) with Package level items (WPR Gantt activities, VDRL submittal deadlines, sub-order material holds, milestone gates).
- Slices deliverables dynamically into **7-Day** and **14-Day** time horizons based on the reporting cut-off date.
- Provides interactive filtering by window (7 Days vs. 14 Days), scope (Master Schedule vs. Package), and critical path status.
- Synced into dedicated sheets: `LookAhead_7Days` and `LookAhead_14Days` in the respective export Excel files.

### 2.4. Universal WPR & MOM Ingestion Engine & Live Execution Governance (`mode="live"`)
- **Live Execution Mode Default (`mode="live"`):** All packages strictly follow their own uploaded `WPR/` and `MOM/` reports (`Priority #1`) compared against the Master Schedule (`SCD`) and Procurement Plan (`PSR`). Packages without an uploaded WPR and/or MOM never borrow PLC template data in Execution Mode — they display only verified information available from the Master Schedule (`SCD`), `PSR`, `VDRL/`, and Client Document Log (`D2`).
- **Gantt Chart ETA Governance & Discrepancy Alerting:** When a package's WPR includes an Attachment 3 Gantt Chart with a shipment/delivery task (e.g., `CIP` Task 70 `4.2 Sea Freight / Transit to China Port (CIF)` finishing on `2027-02-23`), the governing Package ETA follows the Gantt Chart (`2027-02-23`) and compares against the Master Schedule ROS date (`2027-03-05`, `+10 days buffer`). If the WPR Section 5 Key Milestone table lists a different date (e.g., Milestone #10 `2027-02-08`, which is the Task 4.1 loading/departure date), the dashboard highlights and notifies the user of the **+15 days discrepancy**.
- **MOM Situation Summary & Key Concerns (`MOM-0010`):** Extracts meeting metadata, 48 action items (`OPEN`, `CLOSED`, `INFO`), and 14 sub-vendor procurement rows (e.g., `MM-ASK-1A-GEN01-CIP-MEC-MOM-0010.pdf` dated `2026-10-28`), summarizing the overall situation and highlighting critical concerns (such as 28-week manufacturing lead time / +6w impact, expediting Flow Switch & Metering Pump, unpriced sub-POs, and end-of-December pump dispatch).
- **Cut-Off Date & Reference Document Provenance Stamps:** Every Element (Elements 1–9, Attachments 3 & 4, and MOM Integration) displays its governing `Ref Doc` and `Cut-off Date`, and a top-level **Incorporated Uploaded Files Banner** confirms which WPR, MOM, SCD, PSR, and VDRL files are active in the dashboard.
- **Universal Project Detection:** Automatically identifies project and package code from document naming:
  - ASK: `MM-ASK-1A-GEN01-<PKG>-...`
  - Z1F: `MM-ZTK-1F-GEN-<PKG>-...`
- **WPR Extraction:** Extracts weekly report summary metadata, Key Milestones, milestone progress weights, weekly planned vs. actual S-curve coordinates, Attachment 3 fabrication Gantt tasks, and Attachment 4 sub-vendor material expediting orders.
- **MOM Extraction:** Extracts meeting metadata, vendor information, meeting date, action items table (item number, action party, target date, status, description), and sub-order procurement status.
- **Non-Destructive Upsert:** Synchronizes extracted records into `Export Data File_ASK.xlsx` and `Export Data File_Z1F.xlsx` using composite primary keys, updating modified items without removing existing historical records.

### 2.5. Status Vendor Document Dual Monitoring (Element 6)
- **Part 6A (Vendor Document Submission — `VDRL/`):** Monitors vendor submission compliance across `Key Documents (Y)` (`PJM`, `ENG`, `QAC`, `IOM`), `Non-Key Documents (N)`, and `TOTAL` (`Total Docs`, `Submitted`, `Not Yet`, `Overdue`, `Closed`, `% Submitted`, `% Closed`), highlighting Production-Impact and Sub-vendor Ordering key document holds.
- **Part 6B (Official Client Review & Return Log):** Parses official D2 log (`ASK Vendor Document Log Report*.xls/.xlsx`), calculating review turnaround KPIs (`In Review`, `Submitted`, `In Approval`, `APPR`, `AWC`, `APPR-C`, `REC`) and highlighting client/contractor review bottlenecks.

### 2.6. In-Dashboard Upload & One-Click Extraction Portal
- Drag-and-drop file upload modal accessible directly from the top navigation bar.
- Automatic routing of dropped files into their respective folders (`WPR/`, `MOM/`, `VDRL/`, `KOM/`, or workspace root).
- **One-Click Synchronization Button (`⚡ Sync All WPR/MOM to Excel`):** Triggers `/api/sync-all` to extract all newly dropped PDFs into the export Excel workbooks and reloads the dashboard cache.
- **Direct Download Button (`💾 Download Export Data File`):** Triggers `/api/download-export-file?project=ASK|Z1F` to download the updated Excel workbook directly to the local desktop.

### 2.7. Construction Monitoring, 4x2 Platform Grid & Offshore T&I Timeline (Nov 13, 2026 Release)
- **Construction Monitoring View (`#viewConstruction`):** Dedicated executive screen launched via the `Const. Monitoring` top-level button. Synthesizes Master EPC schedules (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf` for ASK and `MM-ZTK-1F-GEN-PJM-SCD-0002.pdf` for Z1F) into:
  1. *Vendor Delay Impact Area of Concern:* Cross-references vendor delivery dates against yard erection gates.
  2. *14-Day Lookahead Construction Activities:* Active engineering, fabrication, and erection deliverables.
  3. *Simplified 21-Day Gantt (7d past + 14d lookahead):* Interactive horizontal bars with duration, float badges, and Critical Path filters.
- **Master Schedule Deduplication:** Automatic grouping and filtering of paginated/repeated Primavera WBS activities (`_clean_and_dedup_master_activities`).
- **Critical Path Definition:** Strictly follows EPC Primavera criteria: `Total Float <= 0` defines critical path activities (`TF = 0.00d` or negative).
- **ASK Dual-Platform Deck Breakdown:** Separate milestone pipelines for `WPA-01` and `WPA-02`, replacing generic main deck lift with granular deck levels: `Lower Column`, `Sump Deck`, `Lower Deck`, `Mezzanine Deck`, `Upper Deck`, and `Vent Boom`.
- **Z1F 4x2 Responsive Platform Grid:** Displays all 4 platforms (`ZWP-21`, `ZWP-23`, `ZWP-22`, `ZWP-20`) in a 4-column by 2-row layout (Jacket & Pile row + Topside row) ordered chronologically by Ready for Transportation date. Topside cards break down total assembly into deck erection levels.
- **Offshore T&I Timeline:** Visualizes the full marine campaign timeline (mobilization, heavy transport, jacket installation, floatover/topside mating, demobilization) extracted from Attachment 3 of `MM-ZTK-1F-GEN-PJM-SCD-0002.pdf`.

---

## 3. Architecture & File Structure

```text
Vendor Tracking/
├── SPEC.md                                    # Technical specification (this document)
├── README.txt                                 # Method of Statement for weekly operations
├── MANUAL_UPLOAD_WRAPUP.md                    # Quick cheat-sheet for manual GitHub & Render uploads
├── Lesson_Learn.md                            # Comprehensive engineering lessons learned & design rules
├── Handoff_Report.md                          # Full system architecture, release notes & handoff guide
├── README.md                                  # Project overview and architecture documentation
├── .gitignore                                 # Git exclusions including heavy PDFs and superseded/
├── superseded/                                # Archived UI backups (e.g., backup_original_ui_20261113)
├── Export Data File_ASK.xlsx                  # Synchronized ASK WPR, Gantt, MOM, KOM, Lookahead data (~164 KB)
├── Export Data File_Z1F.xlsx                  # Synchronized Z1F MOM, KOM, Lookahead data (~80 KB)
├── index.html                                 # Single-page application shell
├── styles.css                                 # Glassmorphism design system & responsive 4x2 grid
├── app.js                                     # Frontend state, Construction Monitoring & T&I widgets
├── server.py                                  # Pure Python HTTP server & REST API
├── export_dashboard_data.py                   # Extraction & synchronization engine
├── tests/
│   └── test_parsers.py                        # Complete automated test suite (15 suites)
└── vendor_dashboard/
    ├── ingestors.py                           # Master schedule deduplication, T&I parser, PDF/Excel ingestors
    ├── reconciler.py                          # WPR > PSR priority, construction monitoring payload generator
    ├── models.py                              # Package catalogs & metadata
    └── extractors/                            # WPR & MOM extraction modules
├── server.py                                  # Zero-dependency Python HTTP server (REST API & static server)
├── export_dashboard_data.py                   # CLI synchronization script for ASK & Z1F
├── requirements.txt                           # Production Python dependencies (openpyxl, pymupdf, pandas, xlrd)
├── render.yaml                                # Render.com Infrastructure-as-Code deployment manifest
├── index.html                                 # Dark glassmorphic Single Page Application
├── styles.css                                 # Glassmorphic responsive styling & layout rules
├── app.js                                     # Frontend interactivity, SVG S-curves, Look-Ahead & modals
├── Export Data File_ASK.xlsx                  # Lightweight synchronized data workbook for ASK (~145 KB)
├── Export Data File_Z1F.xlsx                  # Lightweight synchronized data workbook for Z1F (~78 KB)
├── Attachment 1 - Procurement Plan-Z1F_*.xlsx # Z1F Procurement Plan baseline
├── Attachment 2-Procurement Plan-ASK-*.xlsx   # ASK Procurement Plan baseline
├── ASK Vendor Document Log Report*.xls        # Official D2 Client Review Log
├── vendor_dashboard/
│   ├── __init__.py
│   ├── ingestors.py                           # Multi-source ingestion & caching engine
│   ├── reconciler.py                          # Priority reconciliation, Look-Ahead Hub & executive summaries
│   └── extractors/
│       ├── __init__.py
│       ├── wpr_mom_extractor.py               # PyMuPDF extractor for WPR & MOM PDFs
│       └── excel_sync.py                      # Openpyxl non-destructive upsert engine
├── tests/
│   └── test_parsers.py                        # Automated regression test suite (15 test suites)
├── KOM/                                       # [LOCAL ONLY - Git Ignored] Heavy KOM presentation decks
├── WPR/                                       # [LOCAL ONLY - Git Ignored] Heavy vendor weekly progress reports
├── MOM/                                       # [LOCAL ONLY - Git Ignored] Heavy vendor meeting minutes
└── VDRL/                                      # [LOCAL ONLY - Git Ignored] Vendor document registers
```

---

## 4. Verification & Testing Standards
All modifications must pass the 15 automated verification suites in `tests/test_parsers.py`:
1. `test_vdrl_multi_and_single_package_parsing`: Validates VDRL parsing and multi-package workbook splitting.
2. `test_client_doc_log_parsing`: Validates 271-row binary `.xls` parsing and D2 status categorization.
3. `test_kom_manufacturer_rule`: Enforces strict KOM extraction with blank fallback when missing.
4. `test_kom_export_data_file_fallback`: Validates Excel-based manufacturer fallback.
5. `test_wpr_pdf_and_docx_parsing`: Validates WPR PDF & DOCX extraction across milestones, S-curves, and Gantt tasks.
6. `test_reconciler_priority_and_modes`: Enforces `WPR Priority #1 > PSR Priority #2` reconciliation.
7. `test_server_snapshot_cache_and_elements`: Tests API responses, fingerprint cache invalidation, and headers.
8. `test_feedback_items_1_to_7`: Validates user feedback UX rules (table styling, text wrapping, KPI layout).
9. `test_latest_5_feedback_items`: Validates client log filters, equipment-only packages, and weightings.
10. `test_procurement_plan_ros_priority_and_one_picture_executive_summary`: Validates One-Picture 4-tier cards and mitigation matrix.
11. `test_dynamic_weekly_file_updates_seamless`: Validates automatic freshness sorting `max(mtime, ctime)`.
12. `test_wpr_mom_extraction_and_excel_sync`: Validates WPR/MOM PDF extraction and Excel upsert.
13. `test_master_schedule_ros_and_7_14_days_lookahead_hub`: Validates Master Schedule ROS and Look-Ahead Hub.
14. `test_z1f_scd_and_mom_ingestion`: Validates Z1F Master Schedule (4,110 activities), MOM PDC extraction, and `Export Data File_Z1F.xlsx`.
15. `test_live_execution_mode_and_cip_wpr_mom_reconciliation`: Validates Live Execution Mode default (`mode="live"`), CIP WPR-0010 (`Cut-off 2026-10-27`, `Plan 23.27%`, `Actual 17.99%`), Gantt Chart ETA governance (`2027-02-23` vs Milestone #10 `2027-02-08`, `+15d discrepancy` & `+10d buffer` vs Master Schedule ROS `2027-03-05`), CIP MOM-0010 (`48 items`, `28 weeks / +6w impact`), Element reference/cut-off stamps, and incorporated files notification banner.

---

## 5. Non-Functional Requirements & Operational Rules
1. **Zero-Cache HTTP Headers (`Lesson_Learn.md` §2.8):** Server must always inject `Cache-Control: no-cache, no-store, must-revalidate` and static assets must carry query versioning.
2. **Text Visibility & Responsiveness (`Lesson_Learn.md` §2.4):** Zero text truncation or clipping in table cells, alert boxes, or detail cards.
3. **No Heavy Raw PDF Uploads to GitHub Web:** Keep `KOM/`, `WPR/`, `MOM/`, and large master schedule PDFs on the local working computer. Only the lightweight Excel export files (`Export Data File_ASK.xlsx` and `Export Data File_Z1F.xlsx`) and updated procurement/log spreadsheets are uploaded to GitHub.
4. **Governing Cut-Off Date Extraction (`REPORT DATE:` / `CUT OFF PERIOD:`):** When multiple weekly reports exist for the same package, candidate sorting must compare the contractual `REPORT DATE:` / `CUT OFF PERIOD:` extracted from document headers (ranking `(cut_off_date, rev_num, report_num, mtime)`). The report with the latest cut-off date strictly governs dashboard metrics and S-curve progression.
5. **Dynamic Project Scope Isolation in Look-Ahead Hub:** Look-Ahead Scope dropdown and register subtitle dynamically adapt to the active project (`MM-ZTK-1F-GEN-PJM-SCD-0002` Rev. C1 for Z1F; `MM-ASK-1A-GEN01-PJM-SCD-0001` Rev. B1 for ASK), preventing cross-project label leakage.
6. **Cloud Parity via Export Workbooks (`Export Data File_ASK.xlsx`, `Export Data File_Z1F.xlsx`):** All WPR formats (`.pdf` and `.docx`, including `WPR_Milestones` and `WPR_Progress`), `VDRL_Summary`, `Master_Construction`, and full-scope `LookAhead_7Days`/`LookAhead_14Days` rows are persisted into lightweight Excel workbooks so Render Cloud matches Local execution.
7. **Construction Monitoring Vendor Delay Impact Toggle:** `1. Area of Concern — Vendor Delay Impact to Construction` defaults to hidden (`Show All (N)` / `Hide All`) for both **ASK** and **Z1F**, auto-expanding when drilling down from KPI Card #1.
8. **1-Click GitHub Sync (`POST /api/push-github`):** Supports atomic single-commit pushes of updated Render data files via GitHub Git Data API when `GITHUB_TOKEN` and `GITHUB_REPO` are configured in Render Environment Variables.
