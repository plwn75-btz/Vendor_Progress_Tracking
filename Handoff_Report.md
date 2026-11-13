# Engineering Handoff Report: Weekly Project Management & Vendor Package Monitoring Dashboard (ASK & Z1F)

**Project Name:** Weekly Project Management & Vendor Package Monitoring Dashboard (Aung Sinkha `ASK` & Zawtika `Z1F`)  
**Date:** November 13, 2026 (Updated from November 07, 2026 baseline)  
**Status:** Live WPR/MOM Execution Release, Construction Monitoring & 4x2 Platform Grid Release — All 15 Test Suites Passing & Verified  
**Active Dynamic Data Sources (5 Designated Weekly Drop Locations):**  
1. **`WPR/` (`.pdf`, `.docx`) — Priority #1:**  
   - ASK Live Packages:
     - `PLC System`: `MM-ASK-1A-GEN01-PLC-INS-WPR-0001-3.pdf` (Report Date `2026-10-29`, Rev 02, Week 6, Planned `16.0%` vs. Actual `9.3%`, Variance `-6.7%`) superseding Rev 00 (`2026-10-16`)
     - `CIP Package`: `MM-ASK-1A-GEN01-CIP-MEC-WPR-0010.pdf` (Report Date `2026-10-27`, Cut-off `2026-10-27`, Gantt Task 70 Delivery ETA `2027-02-23`)
   - Z1F Live Package: `MM-ZTK-1F-GEN-CCV-PIP-WPR-0001 0916.docx` (`Compact Check Valve`, Planned `10.0%` vs. Actual `6.0%`)  
2. **`Procurement Plan / PSR` (`.xlsx`) — Priority #2 Fallback Baseline:**  
   - ASK: `Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx` (`98` procurement cycle rows)  
   - Z1F: `Attachment 1 - Procurement Plan-Z1F_20261002.xlsx` (`142` procurement cycle rows)  
3. **`VDRL/` (`23` `.xlsx` Workbooks $\rightarrow$ `25` ASK Packages) — Dynamic Vendor Document Submission Monitoring:**  
   - `1,277` total vendor documents tracked (`419` Submitted [`32.8%`], `221` Overdue)  
   - Multi-package workbooks handled automatically: `ABV`+`ACBV` and `CV`+`SRV`  
4. **`ASK Vendor Document Log Report - 25 Sep 26.xls` (Root / `VDLR/`) — Official & Contractual Client Review & Return Log (D2):**  
   - `271` contractual submittals (`138` In Review, `118` Submitted, `15` In Approval; `100` AWC, `27` APPR, `5` REC, `1` APPR-C, `12` Overdue Review)  
5. **`KOM/`, `WPR/`, `MOM/` (Local PC only) → Separated `Export Data File_ASK.xlsx` & `Export Data File_Z1F.xlsx` (Uploaded):**  
   - Heavy PDFs (`KOM/`, `WPR/`, `MOM/`, ~16 GB at scale) stay safely on local storage and are git-ignored.
   - Extracted data (Gantt tasks, S-curves, milestones, sub-orders, action items, and manufacturer details) is synchronized into **`Export Data File_ASK.xlsx`** (~164 KB) and **`Export Data File_Z1F.xlsx`** (~78 KB) via the Dashboard UI ("⚡ Sync All WPR/MOM to Excel") or `python export_dashboard_data.py --project ALL`.  
   - On Render/GitHub (where raw PDFs are excluded), the dashboard reads all weekly progress, S-curves, and MS Project Gantt schedules directly from these Excel files, ensuring ultra-fast (<5s) boot times, sub-second page loads, and zero memory crashes.  
   - The Dashboard UI modal includes 1-click download buttons (`💾 Download Export Data File (ASK)` and `💾 Download Export Data File (Z1F)`) for manual GitHub web upload workflow.  
**Primary Tech Stack:** Python 3.11 (`server.py`, `vendor_dashboard/`), Vanilla JavaScript (`app.js`), Modern CSS3 (`styles.css`), HTML5 (`index.html`), Specification (`SPEC.md`)

---

## 0. November 01, 2026 Release — Vendor Package Monitoring Dashboard Architecture & Handoff

### 0.1. Separated & Switchable Project Workspaces (`ASK` | `Z1F`)
- The dashboard starts on **ASK** (`Aung Sinkha Development Project Phase 1A — EPC-01`) by default and provides a top-level segmented switcher (`Project: ASK` | `Project: Z1F`) to switch cleanly between **ASK** (`28` active packages) and **Z1F** (`17` active packages) without mixing data across projects.

### 0.2. Priority Rule (`WPR Priority #1 > PSR Priority #2`) & Development Sample Mode
- **Execution Priority Rule:** Because the Procurement Status Report (`PSR`) may lag behind the latest weekly vendor submission, **WPR is Priority #1** whenever any deviation exists between WPR and PSR. If no WPR information exists for a package, the reconciler falls back to **PSR (`Priority #2`)**.
- **Development Phase Sample Mode (`PLC WPR Template`):** Since all packages share an identical WPR template, the top bar includes a mode switch:
  - **`Dev Mode (PLC WPR Sample)`:** Uses `MM-ASK-1A-GEN01-PLC-INS-WPR-0001 WEEKLY PROGRESS REPORTS REV 00.pdf` as a sample template for packages that do not yet have their own WPR uploaded in `WPR/`, while preserving each package's real `KOM/` Manufacturer details, real `VDRL/` submission metrics, real `ASK Vendor Document Log Report` Client review metrics, and real `PSR` dates.
  - **`Execution Mode (Live WPR > PSR)`:** Displays WPR charts strictly for packages with a live uploaded WPR in `WPR/`, and falls back to PSR (`Priority #2`) for all other packages.

### 0.3. Implementation of the 9 Required Dashboard Elements + 7 User Feedback Enhancements (Nov 01, 2026)
| # | Required Dashboard Element / Enhancement | Backend Implementation (`vendor_dashboard/`) | Frontend Rendering (`index.html` / `app.js`) |
| :- | :--- | :--- | :--- |
| **1** | **Project** | `build_project_snapshot(project='ASK'|'Z1F')` | Top-bar Project Switcher (`ASK` \| `Z1F`) + Element 1 KV box |
| **2** | **Package Name** | `ASK_PACKAGE_METADATA` / `Z1F_PACKAGE_METADATA` + VDRL/WPR/PSR header extraction | Package Header Banner (`Code`, `Full Name`, `Discipline`, `RFQ`, `Tag`, `PO`, `Cut-off Date`, `ROS/ETA`) |
| **3** | **Manufacturer Detail** | `parse_kom_folder('KOM')` (`.pdf` & `.pptx` parser) | Element 3 Card: Displays `Manufacturer Name`, `Vendor Code`, `Factory Location`, `KOM Date`, `PO`, `Contacts` from `KOM/`, or blank (`—`) if no KOM exists |
| **4** | **Key Milestone** | `_reconcile_milestones()` (`12` standard milestones, `WPR #1 > PSR #2`) | Element 4 Table (Filterable & Sortable): Planned, Forecast/Actual, Status badge, Source (`WPR` vs `PSR`), & Deviation notes |
| **4.1** | **Package Key Highlight Banner** | `_build_package_highlights()` | 4-Card Highlight Banner: **4.1A** Current Key Milestone Status, **4.1B** Remaining Days to Next Key Milestone, **4.1C** Days to Delivery (ETA), **4.1D** Immutable PSR ROS Date & Delay Check |
| **5** | **Progress Tracking by Key Milestone** | `_build_dynamic_package_progress_and_scurve()` (`17` rows, `100%` weight) | Element 5 Table (Filterable & Sortable): Weight Factor %, Planned %, Actual %, Weighted Actual %, and progress bars |
| **6** | **Status Vendor Document** | **6A:** `parse_all_vdrl_files('VDRL')` (Vendor submission)<br/>**6B:** `parse_client_doc_log()` (`ASK Vendor Document Log Report` Client review & return) | Element 6 Dual Card + Tab 3 Hub (All Filterable & Sortable): Side-by-side 6A VDRL Submission register & 6B Contractual D2 Client Review register |
| **7** | **Overall Progress (S-Curve): Package + Project Weighted Across All Packages** | `_build_dynamic_package_progress_and_scurve()` + `_build_project_weighted_scurve()` | Element 7 Interactive SVG S-Curve with switcher between **Selected Package S-Curve** (updates dynamically per package) and **Weighted All Packages S-Curve** + Package Weighting Breakdown Table |
| **8** | **Look Ahead Plan (7–14 Days)** | `_build_look_ahead_plan()` (Dynamic per package from WPR, Attachment 3 Gantt, VDRL 14d window, & upcoming milestones) | Element 8 Card (Filterable): Dynamically updates when switching packages |
| **9** | **Area of Concern** | `_build_areas_of_concern()` (ROS Delays + Attachment 3/4 Material Holds + Overdue VDRLs + Overdue Client Reviews) | Element 9 Card (Filterable): Severity-coded risk & bottleneck items |
| **Att. 3 & 4** | **WPR Attachment 3 (Fabrication Gantt) & Attachment 4 (Supplier Material Status)** | `_build_attachments_3_and_4_summary()` | Dedicated dual cards for **WPR Attachment 3 (Overall Fabrication Detail Plan — Gantt Situation Analysis)** and **WPR Attachment 4 (Supplier Material Status Expediting Report)** with filter & sort |
| **ROS Rule** | **Governing Procurement Plan ROS Date & Delivery Delay Alert (`ETA > ROS`)** | `_build_package_highlights()` & `_reconcile_milestones()` | Highlights Fixed Procurement Plan ROS Date (`Attachment 2-Procurement Plan-ASK...`) as the governing delivery reference; flags any package where `ETA > ROS Date` (e.g. `Coat`: `ETA 2027-04-08` vs `Fixed ROS 2026-11-23` $\rightarrow$ **`+136 Days Delayed`**) |
| **Exec Summary** | **One-Picture 4-Tier Health Bar + Smart Mitigation Matrix** | `build_project_snapshot()` + `_build_smart_issue_and_mitigation()` | Categorizes all packages into 🔴 Already Delayed, 🟠 High Risk Delay, 🟡 Low Risk (Watch), 🟢 Very Good / Healthy; provides single-window Smart Root-Cause Issue & Timely Mitigation Action Matrix |
| **UI** | **Universal Filter + Ascending/Descending Sort & Daylight Mode** | Zero-cache headers + `?v=20261101_8` | Clickable `▲` / `▼` sort headers and filter inputs on **all tables** + `☀️ Daylight Mode` / `🌙 Dark Mode` toggle button in top bar + Single-window fit without horizontal sliding bar |

### 0.4. Current Workspace File Inventory (Nov 07, 2026)
```text
Vendor Tracking/
├── SPEC.md                                        # Formal Technical Specification (Problem, Solution, 9 Elements, Priority Rules)
├── server.py                                      # Standalone HTTP server, zero-cache headers, mtime fingerprint cache, /api/upload
├── index.html                                     # Dashboard SPA shell (?v=20261101_8)
├── styles.css                                     # Dark-navy glassmorphism, scoped '>' table selectors, 34px aligned controls
├── app.js                                         # Frontend state, ASK/Z1F switcher, Dev/Live mode toggle, SVG S-Curve, Dual Doc Hub
├── export_dashboard_data.py                       # CLI & backend engine for extracting WPR/MOM/KOM/SCD into Export Data Files
├── Export Data File_ASK.xlsx                      # Uploaded: extracted ASK WPR, Gantt, MOM, KOM, Look-Ahead (~164 KB)
├── Export Data File_Z1F.xlsx                      # Uploaded: extracted Z1F MOM, KOM, Look-Ahead (~78 KB)
├── MANUAL_UPLOAD_WRAPUP.md                        # Method of statement & pre-flight wrap-up for manual GitHub web upload
├── requirements.txt                               # openpyxl, pandas, xlrd, pymupdf, python-docx, python-pptx, pytest
├── render.yaml                                    # Render.com web service configuration
├── Handoff_Report.md                              # Engineering handoff report (this document)
├── Lesson_Learn.md                                # Technical lessons learned (§§1–34)
├── README.md                                      # Project overview & quick start
├── vendor_dashboard/
│   ├── __init__.py
│   ├── models.py                                  # Domain models & ASK/Z1F package catalogs
│   ├── ingestors.py                               # Fast parsers for KOM/, VDRL/, Client Doc Log (.xls), PSR (.xlsx), WPR (.pdf/.docx), MOM/
│   ├── reconciler.py                              # Reconciler enforcing WPR > PSR priority, Live/Dev modes, & 9 elements
│   └── extractors/
│       ├── __init__.py
│       └── wpr_mom_extractor.py                   # PDF text & table parsers for WPR covers, Gantts, MOMs (with doc.close())
├── tests/
│   └── test_parsers.py                            # Automated verification suite (15 test suites — ALL PASSING)
├── WPR/                                           # Designated Drop Location #1: Weekly Progress Reports (.pdf, .docx)
├── MOM/                                           # Designated Drop Location: Minutes of Meeting (.pdf, .docx)
├── VDRL/                                          # Designated Drop Location #3: 23 VDRL workbooks (25 ASK packages, .xlsx)
├── KOM/                                           # LOCAL ONLY (git-ignored): Kick-Off Meeting PDF decks
├── ASK Vendor Document Log Report - 25 Sep 26.xls # Designated Drop Location #4: Official Contractual Client Review Log (271 rows)
├── Attachment 2-Procurement Plan-ASK - 20260626.xlsx # Designated Drop Location #2 (Active ASK Procurement Plan)
├── Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx # Designated Drop Location #2 (ASK Procurement Plan Baseline)
├── Attachment 1 - Procurement Plan-Z1F_20261002.xlsx  # Designated Drop Location #2 (Z1F PSR Fallback Baseline)
├── MM-ASK-1A-GEN01-PJM-SCD-0001.pdf               # Master Schedule for ASK Rev. B1 (2,522 activities, git-ignored)
└── MM-ZTK-1F-GEN-PJM-SCD-0002.pdf                 # Master Schedule for Z1F Rev. C1 (4,110 activities, git-ignored)
```

### 0.5. Ready-to-Resume & Manual GitHub / Render Guide

#### Option A — GitHub Web Upload WITHOUT `KOM/` (Nov 02, 2026)
1. Refresh the KOM extract on the local PC (only needed after new KOM decks arrive; the dashboard also does it automatically):
   ```powershell
   python export_kom_data.py
   ```
2. Upload in batches that each stay **under 25 MB** (do **not** upload `KOM/` or `__pycache__/`):

   | Batch | Contents | Approx. Size |
   | :- | :--- | :--- |
   | 1 | Root files (`*.py`, `*.html`, `*.js`, `*.css`, `*.md`, `requirements.txt`, `render.yaml`, `.gitignore`, **`Export Data File.xlsx`**, Procurement Plans, D2 Log, Master Schedule PDF) | ~9.1 MB |
   | 2 | `vendor_dashboard/`, `tests/`, `WPR/` | ~1.1 MB |
   | 3 | `VDRL/`, first half of files (A–M by name) | ~14 MB |
   | 4 | `VDRL/`, second half of files | ~14 MB |

   > `VDRL/` is 28.1 MB in total, so it **must** be split into two commits.

#### Option B — Git CLI (no size batching needed)

#### Why Web Browser Upload Failed:
GitHub's Web UI (`github.com/.../upload`) enforces a strict **25 MB limit** per single file and per batch commit upload. With converted presentation PDFs in `KOM/`, registers, and master schedules, the workspace is **~75 MB**, which exceeds the web browser upload limit and triggers GitHub's warning:
> *"The file is too large and cannot be uploaded. Consider creating the file in a local clone and pushing it to GitHub."*

#### The Permanent Solution: Git CLI Workflow
Pushing via Git CLI in your local PowerShell has a limit of **100 MB per file** (our largest file is only 6.8 MB), eliminating all web browser batch limits and allowing you to push everything in one command.

1. **One-Time Git Setup (Link Local Workspace to GitHub Repository):**
   ```powershell
   # Inside "Vendor Tracking" directory:
   git init
   git branch -M main

   # Add your GitHub remote repository URL (replace with your repo URL):
   git remote add origin https://github.com/<YOUR-USERNAME>/<YOUR-REPO-NAME>.git

   # If the remote already has files from previous web commits, sync them:
   git pull origin main --allow-unrelated-histories
   ```

2. **Weekly Commit & Push Steps (Takes <5 seconds):**
   ```powershell
   # 1. Stage all changes
   git add .

   # 2. Commit with descriptive summary
   git commit -m "feat: complete vendor dashboard with converted PDF KOMs"

   # 3. Push cleanly to GitHub (Render auto-builds immediately)
   git push -u origin main
   ```

3. **Start Local Server:**
   ```powershell
   python server.py
   ```
   Open `http://localhost:8080` in browser.

4. **Run Automated Verification Suite:**
   ```powershell
   python tests/test_parsers.py
   ```

### 0.6. Seamless Weekly Dynamic Updates Architecture & Trial Test Results
To ensure user weekly updates require zero manual configuration, the backend implements seamless automated updates:
1. **Designated Weekly Drop Locations:**
   - **`WPR/`:** Place new Weekly Progress Reports (`.pdf` or `.docx`). New reports for an existing package automatically supersede older ones based on `_file_freshness_key = max(mtime, ctime)`. In Live Execution mode, dropping a WPR immediately elevates that package from PSR Priority #2 to Priority #1 Live WPR.
   - **`VDRL/`:** Place new or updated vendor document registers (`.xlsx`). Workbooks are sorted by freshness key ascending so higher revisions overwrite earlier ones cleanly.
   - **`Workspace Root (or VDLR/)`:** Place the weekly updated Client Review & Return Log (`ASK Vendor Document Log Report*.xls/.xlsx`). The system automatically picks the newest file by freshness key.
   - **`Workspace Root (or PSR/)`:** Place the weekly updated Procurement Plan (`Attachment 2-Procurement Plan-ASK*.xlsx`). The newest file is automatically selected as governing data reference for ROS dates and milestone baselines.
   - **`KOM/`:** Place newly held Kick-Off Meeting slides (`.pdf` or `.pptx`). Immediately populates manufacturer and PO details.
2. **Dual-Tier Cache Busting & Hot Reload:**
   - `_compute_workspace_fingerprint()` generates a composite tuple across all watched directories and files.
   - When any file is added, modified, or removed, `server.py` invalidates the snapshot cache (`_SNAPSHOT_CACHE`) and clears the low-level parser cache (`_INGEST_CACHE`).
   - Browser requests carry `Cache-Control: no-cache, no-store, must-revalidate` zero-cache headers, guaranteeing the UI renders fresh data upon browser reload or clicking `↻ Refresh Data`.
3. **Automated Trial Verification Suite:**
   - Validated via `test_dynamic_weekly_file_updates_seamless()` in `tests/test_parsers.py`:
     - Simulated dropping a newer Procurement Plan $\rightarrow$ automatically detected and updated headline & dates.
     - Simulated dropping a new Live WPR for `CIP` $\rightarrow$ automatically switched from PSR to Priority #1 Live WPR.
     - Simulated dropping an updated VDRL file $\rightarrow$ automatically updated registers and counts.
     - Simulated dropping an updated Client Document Log $\rightarrow$ automatically updated review numbers.
     - Simulated server fingerprint invalidation $\rightarrow$ cache busted and returned fresh snapshot.
     - Cleaned up all trial artifacts in `try...finally` with 100% baseline restoration.

### 0.7. Master Schedule ROS Resumption & 7 & 14 Days Look-Ahead Hub (Nov 03, 2026)
1. **Governing ROS Resolution & Priority Rules:**
   - **Priority #1:** `MM-ASK-1A-GEN01-PJM-SCD-0001` (Project Master Schedule Rev. B1) ETA/ROS date.
   - **Priority #2 (Fallback):** Live Procurement Plan (`Attachment 2-Procurement Plan-ASK...`) ROS date.
   - **Reconciler Implementation (`vendor_dashboard/reconciler.py`):**
     $$\text{Governing ROS} = \text{Master Schedule ROS} \parallel \text{Procurement Plan ROS}$$
   - Both baselines are preserved explicitly (`master_eta_date` and `psr_ros_date`) and displayed side-by-side in Card 6 and Card 4.1C with discrepancy alerts.
2. **7 & 14 Days Look-Ahead Hub Engine:**
   - **Scope 1 (Master Schedule):** Parses all 2,522 Primavera P6 activities from `MM-ASK-1A-GEN01-PJM-SCD-0001.pdf`, indexes Total Float (`TF=0` Critical Path flag), and slices activities scheduled to start or finish within 7-day (`2026-10-16` to `2026-10-23`) and 14-day (`2026-10-16` to `2026-10-30`) windows.
   - **Scope 2 (Package Level):** Aggregates package-specific WPR Gantt tasks, upcoming VDRL document submittal due dates, Attachment 4 material holds, and procurement milestone gates.
   - **Volume:** Generates 599 activities in the 7-day look-ahead and 762 activities in the 14-day look-ahead.
3. **Multi-Tab Excel Synchronization (`Export Data File.xlsx`):**
   - Automatically writes to dedicated sheets `LookAhead_7Days` and `LookAhead_14Days` alongside `KOM_Manufacturer`, `WPR_Overview`, `WPR_Gantt_Schedules`, `MOM_Action_Items`, and `MOM_Discussions`.
   - File size remains extremely lightweight (~145 KB total) and is git-tracked while excluding gigabytes of raw PDFs.
4. **Interactive Dashboard Tab ("🔮 7 & 14 Days Look-Ahead Hub"):**
   - Accessible via the `#tabBtnLookAhead` tab button.
   - Includes 4 KPI summary cards (7-Day Total, 14-Day Total, Critical Path `TF=0` Count, Package Milestone Count).
   - Dynamic toggle buttons for switching between 7-Day and 14-Day windows.
   - Filter bar: Scope (All / Master Plan / Package Level), Critical Path Only (`TF=0`), Category, and full-text keyword search.
5. **In-Dashboard Sync & Direct Download:**
   - Click `🔄 Run Extraction & Sync Excel` to trigger `/api/sync-all` and rebuild all Excel sheets.
   - Click `📥 Download Export Data File` to immediately download the project's export file via `/api/download-export-file`.

### 0.8. Separated Project Export Data Files (`Export Data File_ASK.xlsx` & `Export Data File_Z1F.xlsx`) (Nov 03, 2026)
- **Independent Project Workbooks:**
  - **ASK:** [`Export Data File_ASK.xlsx`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/Export%20Data%20File_ASK.xlsx) (~145 KB) contains all ASK KOM manufacturer details, WPR weekly summaries, S-curves, Gantt schedules, sub-orders, MOM action items, and 7/14 days look-ahead tasks.
  - **Z1F:** [`Export Data File_Z1F.xlsx`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/Export%20Data%20File_Z1F.xlsx) (~38.5 KB) contains Z1F-specific KOM manufacturer details, WPR/MOM data, and 7/14 days look-ahead tasks.
- **Dynamic UI Download Routing:**
  - Switching between `Project: ASK` and `Project: Z1F` on the top bar dynamically updates the download button label (`💾 Download Export Data File (ASK)` vs `💾 Download Export Data File (Z1F)`) and routes to `/api/download-export-file?project=ASK|Z1F`.
- **Dual & Selective CLI Extraction:**
  - `python export_dashboard_data.py`: Syncs **both** ASK and Z1F in a single execution.
  - `python export_dashboard_data.py --project ASK`: Syncs only ASK.
  - `python export_dashboard_data.py --project Z1F`: Syncs only Z1F.
- **GitHub Web Browser Compatibility:**
  - Both files stay well below 200 KB, allowing drag-and-drop web commits to `github.com` within seconds from restricted working computers.

### 0.9. Z1F Master Schedule & MOM Ingestion Verification (Nov 03, 2026)
- **Z1F Project Master Schedule (`MM-ZTK-1F-GEN-PJM-SCD-0002.pdf` Rev. C1):**
  - **Scale & Coverage:** Successfully parsed 216-page Primavera master schedule containing **4,110 activities**.
  - **Catalog Mapping:** Mapped key Z1F packages including Pedestal Crane (`PDC`: Activity `1F-WP1-WU1.4.1-1250`, ETA `2027-06-22`, Yard Erection `2027-08-16`, Total Float `TF=10`), Closed Drain System (`CDS`), TEG Contactor (`TSP`), Storage Tank (`TANK`), Chemical Injection (`CIP`), Instrument Gas Filter (`IGFIL`), Actuated Choke Valve (`ACV`), and Wellhead Control Panel (`WHCP`).
  - **Look-Ahead Hub Generation:** Sliced master schedule and package tasks into immediate operational horizons:
    - **7-Day Window:** **344 activities** (Master Schedule Engineering/Procurement/Fabrication + Package gates).
    - **14-Day Window:** **550 activities**.
- **Z1F Minutes of Meeting (`MOM/MM-ZTK-1F-GEN-PDC-MEC-MOM-0010 Rev.00.pdf`):**
  - **Universal Regex Package Detection:** Broadened regex in `vendor_dashboard/extractors/wpr_mom_extractor.py` to identify ZTK/Z1F packages (`MM-ZTK-1F-GEN-PDC-MEC-MOM-0010 Rev.00.pdf` $\rightarrow$ Package `PDC`, MOM No `MOM-0010`, Meeting Date `2026-10-09`, Vendor Favelle Favco).
  - **Action Items Extraction:** Extracted all **11 action items** with statuses (`INFO`, `CLOSED`, `OPEN`) and detailed descriptions.
- **Excel Synchronization (`Export Data File_Z1F.xlsx`):**
  - File size: **78.3 KB** (completely independent from `Export Data File_ASK.xlsx`).
  - Populated Sheets: `KOM_Manufacturer` (10 packages), `README`, `MOM_Action_Items` (11 PDC action items), `MOM_Procurement_Status`, `LookAhead_7Days` (344 activities), and `LookAhead_14Days` (550 activities).
- **Automated Test Suite:** Added Test Suite 14 (`test_z1f_scd_and_mom_ingestion`) in `tests/test_parsers.py`. Confirmed all 14 test suites pass (`ALL TESTS PASSED SUCCESSFULLY`).

### 0.10. Project Toggle & Dropdown Alignment Resolution (Nov 04, 2026)
- **Data & State Alignment:**
  - When switching the Project toggle between **ASK** and **Z1F**, the Look-Ahead Scope dropdown (`#lookAheadScopeFilter`) and header subtitle (`#lookAheadSubtitle`) dynamically update to match the active project's master document:
    - **Project: Z1F:** Dropdown displays `Master Schedule Only (MM-ZTK-1F-GEN-PJM-SCD-0002)` and subtitle displays `Integrated from MM-ZTK-1F-GEN-PJM-SCD-0002 (Rev. C1) & Package Schedules`.
    - **Project: ASK:** Dropdown displays `Master Schedule Only (MM-ASK-1A-GEN01-PJM-SCD-0001)` and subtitle displays `Integrated from MM-ASK-1A-GEN01-PJM-SCD-0001 (Rev. B1) & Package Schedules`.
- **Visual & Layout Alignment (CSS):**
  - Standardized all interactive controls (`.segmented-control`, `.seg-btn`, `.select-control`, `.input-control`, and `.btn-action`) to a uniform `34px` height with `box-sizing: border-box` and flex vertical centering, eliminating vertical height discrepancies and baseline jitter.
  - Converted the Look-Ahead window toggle (`#lookAheadWindowToggle`) to standard `.seg-btn active` pill controls matching top navigation styling.
  - Restricted the quick package selector (`#quickPackageFilterBar`) visibility to Tab 1 (`9-Point Package Monitoring Workspace`), preventing stray controls from wrapping onto a second line across project-wide tabs (Look-Ahead Hub, Portfolio Matrix, and Client Review Hub).

### 0.11. Live WPR/MOM Execution Release & Report Date Cut-Off Priority (Nov 07, 2026)
- **Governing WPR Cut-Off Date via PDF Cover Header Scan:**
  - Resolved issue where file system modification time (`mtime`) incorrectly prioritized an earlier report (e.g. `16-OCT-2026` Rev 00) over a newly uploaded report (`29-OCT-2026` Rev 02).
  - The governing cut-off date is now strictly extracted from the WPR PDF cover page header (`REPORT DATE:` or `CUT OFF PERIOD:`), supporting alphanumeric date formats like `29-OCT-2026`.
  - Candidate WPR files are sorted by `(cut_off_date, rev_num, report_num, mtime)` descending so that the latest cut-off date always governs.
  - Implemented dynamic 17-milestone progress extraction anchoring S-curve values (e.g., Week 6 `2026-10-29`: planned `16.0%`, actual `9.3%`, variance `-6.7%`).
- **Live Execution Mode (`mode="live"`):**
  - Released all packages to follow their live WPR and MOM data rather than falling back to the Development PLC template.
  - Packages with live reports (`PLC`, `CIP`, `CCV`) display their extracted S-curves, milestones, Gantt schedules, material holds, and action items.
  - Packages without a WPR gracefully display their live baseline from Procurement Plan (PSR Priority #2) and VDRL registers.
- **CIP Package Live WPR & MOM Reconciliation:**
  - `MM-ASK-1A-GEN01-CIP-MEC-WPR-0010.pdf` (Report Date `2026-10-27`): MS Project Gantt Task 70 governs package delivery with ETA `2027-02-23` (CIF China port).
  - Compared with Master Schedule ROS (`2027-03-05`, +10d positive buffer) and Key Milestone #10 Contractual Delivery (`2027-02-08`, +15d discrepancy alert surfaced on dashboard).
  - `MM-ASK-1A-GEN01-CIP-MEC-MOM-0010`: Ingested meeting situation analysis, expediting actions, and action item register.
- **Dynamic Project Isolation in Look-Ahead Schedule Register:**
  - Neutralized hardcoded ASK document number in `index.html`.
  - Added `#optMasterScheduleScope` and dynamically bound dropdown and subtitle in `switchProject()` and `renderLookAheadHub()`:
    - **Z1F:** `Master Schedule Only (MM-ZTK-1F-GEN-PJM-SCD-0002)` (Rev. C1, Cut-off `2026-09-16`).
    - **ASK:** `Master Schedule Only (MM-ASK-1A-GEN01-PJM-SCD-0001)` (Rev. B1, Cut-off `2026-10-16`).
- **PyMuPDF Windows File Handle Safety:**
  - Added immediate `doc.close()` in `extract_wpr_pdf` and `extract_mom_pdf` right after `full_text_pages` extraction, eliminating Windows file locking (`WinError 32`) during temporary file cleanup.
- **Automated Verification:**
  - Test Suite 15 (`test_live_execution_mode_and_cip_wpr_mom_reconciliation`) added to `tests/test_parsers.py`.
  - All **15 test suites** pass cleanly.

---

## 1. Prior Baseline System Architecture Overview (Procurement Tracking Reference)

The **Procurement Tracking Dashboard** baseline principles carried forward into this project:
- **Z1F (Zawtika Project Development - Topside & Jacket - WP1)**
- **ASK (Aung Sinkha Project Development - Overall Cycle - WP2)**

### Architecture Design:
1. **Backend Service (`server.py`):**
   - Built with pure Python `http.server` (`DashboardHandler`), ensuring lightweight deployment with zero heavy WSGI framework dependencies (no Flask/Django).
   - Serves static web assets with zero-cache headers (`must-revalidate`) and exposes RESTful API endpoints:
     - `GET /api/dashboard?project=ASK|Z1F&mode=sample|live`: Returns full reconciled JSON snapshot.
     - `GET /api/package?project=ASK|Z1F&code=<PKG>`: Returns single-package 9-element detail.
     - `POST /api/upload`: Routes uploaded weekly files (`WPR`, `VDRL`, `VDLR`, `PSR`, `KOM`) into their designated drop folders and invalidates the in-memory snapshot cache immediately.

2. **Frontend SPA (`index.html`, `styles.css`, `app.js`):**
   - Single-page application styled with a sleek dark glassmorphism aesthetic (`#0a0f1d` background, subtle borders, cyan/blue/emerald gradients).
   - Dynamic interactivity via `app.js`: real-time search filtering, project switching, mode toggling, and interactive file drag-and-drop.

---

## 2. Core Features & Capabilities Carry-Over Reference

### 2.1. Expanded Detail View & Table Layout Precision
- **Strict Left-Alignment & Scoped Selector Isolation:** Outer data table (`.data-table`) rules are strictly scoped using direct child selectors (`>`) so they never leak into nested inner elements (`tr.detail-row td`).
- **Multi-Line Word Wrapping & Overflow Resilience:** Detail boxes and values (`.kv-val`, `.alert-title`, `.alert-detail`) enforce `white-space: normal !important; overflow-wrap: anywhere !important; word-break: break-word !important;`. Long identifiers wrap cleanly onto additional lines and naturally expand the card height, ensuring 100% of text is visible with zero clipping or truncation.

### 2.2. Static Asset Caching Prevention (Crucial Operational Note)
To prevent CDNs and browsers from caching old CSS/JS files after an update:
1. `server.py` includes custom headers inside `DashboardHandler.end_headers()`:
   ```python
   self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
   self.send_header("Pragma", "no-cache")
   self.send_header("Expires", "0")
   ```
2. `index.html` references versioned assets (`styles.css?v=20261101_1` and `app.js?v=20261101_1`). **Always bump the `?v=` version string in `index.html` whenever you modify CSS or JS.**

---

## 3. Operational Policies & Project Directives

1. **No Power BI Execution Required:**
   - The application runs exclusively as a standalone Python web service and browser SPA.
2. **Manual Git & Render.com Synchronization:**
   - Automated git commits and git pushes by AI agents are strictly forbidden per user instruction.
   - All version control commits and pushes to GitHub and Render.com deployments are performed manually by the repository owner on a weekly update basis.
3. **Verification Checklist Before Closing Any Session:**
   - [x] Run `python tests/test_parsers.py` and confirm all 15 test suites pass (`ALL TESTS PASSED SUCCESSFULLY`).
   - [x] Confirm `SPEC.md`, `Handoff_Report.md`, `Lesson_Learn.md`, and `MANUAL_UPLOAD_WRAPUP.md` are up-to-date.
   - [x] Synchronize export files: `python export_dashboard_data.py --project ALL`.

---

## 4. Manual Deployment & GitHub / Render Upload Guide

Because your working computer has corporate security restrictions preventing Git CLI operations, use the **GitHub Web Browser Upload Workflow** described in detail in [`MANUAL_UPLOAD_WRAPUP.md`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/MANUAL_UPLOAD_WRAPUP.md):

### 4.1. Primary Method: GitHub Web Browser Upload (Working Computer)
1. **Local Extraction:** Ensure both project export files are updated:
   ```powershell
   python export_dashboard_data.py --project ALL
   ```
2. **Navigate to GitHub:** Open `https://github.com/<YOUR-USERNAME>/<YOUR-REPO-NAME>` in your browser.
3. **Upload Files:** Click **Add file** $\rightarrow$ **Upload files**.
4. **Drag and Drop Modified Files (<25 MB per file):**
   - [`Export Data File_ASK.xlsx`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/Export%20Data%20File_ASK.xlsx) (~164 KB)
   - [`Export Data File_Z1F.xlsx`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Vendor%20Tracking/Export%20Data%20File_Z1F.xlsx) (~78 KB)
   - Updated Python & Web files if modified (`vendor_dashboard/extractors/wpr_mom_extractor.py`, `vendor_dashboard/ingestors.py`, `index.html`, `app.js`, `export_dashboard_data.py`, `SPEC.md`, `Lesson_Learn.md`, `Handoff_Report.md`, `MANUAL_UPLOAD_WRAPUP.md`)
   - Any new Procurement Plan (`Attachment 2-Procurement Plan-ASK...`) or Client Log (`ASK Vendor Document Log Report...`)
   - **DO NOT UPLOAD:** `KOM/`, `WPR/`, `MOM/`, or master schedule PDFs (`*.pdf` > 25 MB). They stay safely on your local machine; all their data is already cleanly captured in the lightweight `Export Data File_*.xlsx` workbooks.
5. **Commit:** Enter a commit message (e.g. `feat: live WPR cut-off date selection and Z1F look-ahead isolation`) and click **Commit changes**.
6. **Render Auto-Deployment:** Render automatically detects the commit via webhook, builds the environment (`pip install -r requirements.txt`), and starts `server.py` in under 3 minutes.

### 4.2. Alternative Method: Git CLI (Admin Computer Only)

Open PowerShell in the project directory (`c:\Users\pipes\OneDrive\Documents\Google_AntiGravity\Project\Vendor Tracking`):

```powershell
# 1. Initialize git (if not already a git repository)
git init

# 2. Add remote repository URL (replace with your GitHub repository URL)
# Example: git remote add origin https://github.com/your-username/vendor-tracking-dashboard.git
git remote add origin <YOUR_GITHUB_REPO_URL>

# 3. Check status (all temporary files, cache, and office locks are ignored by .gitignore)
git status

# 4. Stage all files
git add .

# 5. Commit with a meaningful message
git commit -m "feat: Add Master Schedule PJM-SCD-0001 priority tracking and One-Picture Executive Summary"

# 6. Set main branch and push to GitHub
git branch -M main
git push -u origin main --force
```

### 4.2. Deploy to Render.com (Web Service)

1. **Sign in to Render.com:**
   - Go to [dashboard.render.com](https://dashboard.render.com/).
2. **Create New Web Service:**
   - Click **New +** $\rightarrow$ **Web Service**.
   - Select **Build and deploy from a Git repository**.
   - Connect your GitHub repository (`vendor-tracking-dashboard`).
3. **Configure Service Settings (Render will auto-detect from `render.yaml` or fill manually):**
   - **Name:** `ask-z1f-vendor-tracking-dashboard`
   - **Region:** Singapore (`singapore`) or nearest to your users
   - **Branch:** `main`
   - **Runtime:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python server.py`
   - **Instance Type:** `Free` (or standard paid instance)
4. **Environment Variables:**
   - `PYTHON_VERSION`: `3.11.8`
   - `PORT`: `8080` (Render sets this automatically, and `server.py` reads `os.environ.get("PORT", "8080")`)
5. **Deploy:**
   - Click **Deploy Web Service**.
   - Once deployed, your dashboard will be live at `https://ask-z1f-vendor-tracking-dashboard.onrender.com`.

### 4.3. Weekly Manual Update Workflow
When you drop new weekly files (`WPR/`, `VDRL/`, `KOM/`, `MM-ASK-1A-GEN01-PJM-SCD-*.pdf`, or `Attachment 2-Procurement Plan-*.xlsx`):
1. Test locally: `python server.py` and open `http://localhost:8080`.
2. Commit and push:
   ```powershell
   git add .
   git commit -m "chore: Update weekly cut-off documents"
   git push origin main
   ```
3. Render will automatically detect the push and redeploy the latest dashboard.

---

## 5. November 13, 2026 Release — Construction Monitoring, Master Schedule SCD Lookahead, ASK Dual Platform (WPA-01/02) Deck Erection Breakdown, Z1F 4x2 Platform Grid with Topside Deck Breakdown & Attachment 3 T&I Timeline

### 5.1. Construction Monitoring Module (`#viewConstruction`)
- **Executive Access:** Added **`Const. Monitoring`** button on the Executive Summary navigation bar next to Core Packages / All Equipment toggles.
- **Key Capabilities:**
  1. **Vendor Delay Impact Area of Concern:** Correlates late vendor packages (where `ETA > Master Schedule ROS`) directly to downstream Yard Topside Erection and Hook-up activities.
  2. **14-Day Forecast Construction Activities:** Extracts and displays active EPC construction activities from `MM-ASK-1A-GEN01-PJM-SCD-0001.pdf` (ASK) and `MM-ZTK-1F-GEN-PJM-SCD-0002.pdf` (Z1F).
  3. **Simplified Gantt Chart Window (7 Days Before Today & 14 Days Lookahead):** Renders all construction activities within a 21-day window relative to project cut-off with duration badges, status pills, and Critical Path indicators.

### 5.2. Master Schedule Activity Deduplication
- Implemented `_clean_and_dedup_master_activities()` in `vendor_dashboard/ingestors.py`.
- Resolves duplicate entries from multi-page Primavera PDF extractions by grouping on normalized `(activity_id, normalized_name, start_date, finish_date)` composite keys and retaining the single most descriptive activity row.

### 5.3. Critical Path Clarification & Definition
- Clarified the contractual basis of Critical Path in Primavera P6: an activity is critical when **`Total Float <= 0 days`** (`TF = 0.00d` or negative), meaning any delay immediately shifts the project completion or sail-away date.
- Master schedule activities with positive float (`TF > 0`) are categorized as non-critical with available scheduling buffer.

### 5.4. ASK Dual-Platform Deck-Level Erection Breakdown (`WPA-01` & `WPA-02`)
- Updated ASK Executive Milestone cards to present both platforms independently:
  - Replaced generic "Main Deck Lift" with granular deck-level erection milestones:
    - `WPA-01 Lower Column Erection` (`2027-01-13..01-18`)
    - `WPA-01 Sump Deck Erection` (`2027-01-19..02-19`)
    - `WPA-01 Lower Deck Erection` (`2027-02-16..03-16`)
    - `WPA-01 Mezzanine Deck Erection` (`2027-03-16..04-08`)
    - `WPA-01 Upper Deck Erection` (`2027-04-10..05-08`)
    - `WPA-01 Vent Boom - Erection` (`2027-05-10..06-01`)
  - Replicated corresponding structure for `WPA-02` with critical equipment delivery dates at worksite (`2027-06-02`).
  - Milestones are sorted chronologically by start date.

### 5.5. Z1F Responsive 4x2 Platform Milestone Grid (`4 Jackets & Piles + 4 Topsides`)
- Re-architected Z1F key dates into a clean **`4×2` responsive grid** displaying all 4 platforms:
  - **Top Row (Jacket & Pile):** `ZWP-21 Jacket & Pile`, `ZWP-23 Jacket & Pile`, `ZWP-22 Jacket & Pile`, `ZWP-20 Jacket & Pile`.
    - Milestones: `1) First Cut`, `2) MC & Testing`, `3) Ready for Loadout`, `4) Pile Fab. Complete`.
  - **Bottom Row (Topside):** `ZWP-21 Topside`, `ZWP-23 Topside`, `ZWP-22 Topside`, `ZWP-20 Topside`.
    - Milestones: `1) First Cut`, `2) Deck Erection per level` (`Lower Col.`, `Sump Deck`, `Lower Deck`, `Mezzanine Deck`, `Upper Deck`, `Vent Boom`), `3) Crit. Equip. at Site`, `4) MC & Pre-Comm`, `5) Ready for Loadout`.
  - **Chronological Column Arrangement:** Ordered by Ready for Transportation date:
    - Column 1: `ZWP-21` (Jacket `2027-08-03` / Topside `2027-09-17`)
    - Column 2: `ZWP-23` (Jacket `2027-08-16` / Topside `2027-09-28`)
    - Column 3: `ZWP-22` (Jacket `2027-09-30` / Topside `2027-11-12`)
    - Column 4: `ZWP-20` (Jacket `2027-10-14` / Topside `2027-11-26`)

### 5.6. Transportation & Installation (T&I) Timeline Widget
- Integrated offshore Transportation & Installation campaign timeline extracted from **Attachment 3** of `MM-ZTK-1F-GEN-PJM-SCD-0002.pdf`:
  - `T&I Campaign Mobilization`: `2027-08-01`
  - `ZWP-21 T&I Window`: `2027-08-15..09-25`
  - `ZWP-23 T&I Window`: `2027-08-28..10-10`
  - `ZWP-22 T&I Window`: `2027-10-05..11-20`
  - `ZWP-20 T&I Window`: `2027-10-20..12-05`
  - `Offshore Demobilization & Final Handover`: `2027-12-15`
- Rendered as an interactive visual campaign bar in both Executive Summary and Construction Monitoring views.

### 5.7. Housekeeping: Superseded Backup Folder Cleanup
- Moved `backup_original_ui_20261113/` into `superseded/` directory.
- Added `superseded/` to `.gitignore` to prevent unnecessary upload bloat.

