# Lessons Learned: Vendor Package & Procurement Monitoring Dashboard (ASK & Z1F)

**Document Date:** November 01, 2026 (Updated from July 28 / October 09, 2026 baseline)  
**Project:** Weekly Project Management & Vendor Package Monitoring Dashboard (Aung Sinkha `ASK` & Zawtika `Z1F`)  
**Author:** Engineering Team  

---

## 1. Executive Summary

During the implementation and cloud deployment (Render) of the **Procurement & Vendor Package Monitoring Dashboards** and their interactive **Weekly Dynamic Drop & Upload Centers**, several technical challenges and edge cases were encountered and resolved. This document captures the key engineering lessons learned across full-stack architecture, pure Python HTTP server customization, high-speed multi-format document parsing (`.xlsx`, `.xls`, `.pdf`, `.docx`, `.pptx`), priority reconciliation (`WPR` vs. `PSR`), and browser/CDN static asset caching to guide future development and maintenance.

---

## 2. Key Lessons & Technical Takeaways

### 2.1. Cloud Deployment & Static Asset Caching (`SimpleHTTPRequestHandler` vs. Browser/CDN Cache)
- **The Challenge:** After pushing updated frontend code (`index.html`, `styles.css`, `app.js`) to Git and deploying to Render, newly added buttons or components rendered without styling or failed with `ReferenceError` because browsers executed stale cached CSS/JS.
- **Root Cause:** By default, Python's `http.server.SimpleHTTPRequestHandler` serves static files (`.css`, `.js`, `.html`) without explicit anti-caching headers (`Cache-Control: no-cache`).
- **The Solution & Best Practice:**
  1. **Server-Side Cache Control:** Override `end_headers(self)` inside `DashboardHandler` (`server.py`) to explicitly send cache-prevention headers (`Cache-Control: no-cache, no-store, must-revalidate`, `Pragma: no-cache`, `Expires: 0`).
  2. **Frontend Cache-Busting Versioning:** Always append version query parameters to `<link>` and `<script>` tags in `index.html` (`styles.css?v=20261101_1`, `app.js?v=20261101_1`) whenever changes are made.

### 2.2. Dynamic Date-Based File Naming (`YYYYMMDD`) & Multi-Project Sorting
- **The Challenge:** Source files arrive with varying date suffixes across weekly updates.
- **The Solution & Best Practice:** Use modification-time (`os.path.getmtime`) and date-suffix sorting so the newest weekly cut-off file is automatically selected, and fingerprint `(rel_path, mtime_ns, size)` across all watched folders so any file dropped in place immediately invalidates the in-memory cache.

### 2.3. Pure Python Upload Handling Without External WSGI Dependencies
- **The Challenge:** Keeping the server lightweight (`http.server`) without Flask or Django while supporting browser file uploads across 5 document categories (`WPR`, `VDRL`, `VDLR`, `PSR`, `KOM`).
- **The Solution & Best Practice:** Support JSON base64 payload uploads (`POST /api/upload`) that route files directly to their designated drop folder (`WPR/`, `VDRL/`, `KOM/`, or workspace root) and clear `_SNAPSHOT_CACHE` immediately.

### 2.4. Real-Time In-Memory Cache Invalidation upon File Drop or Upload
- **The Challenge:** Users drop new files directly into `WPR/`, `VDRL/`, `KOM/`, or the workspace root via Windows Explorer or upload them via the browser and expect the dashboard to update without restarting `server.py`.
- **The Solution & Best Practice:** Compute a lightweight `_compute_workspace_fingerprint()` tuple of `(relpath, st_mtime_ns, st_size)` across all watched directories on every `/api/dashboard` call. If any file is added, modified, or replaced, rebuild the snapshot automatically.

### 2.5. Drag-and-Drop Auto-Detection & Category Routing
- **The Solution & Best Practice:** Provide explicit category selection in the Upload Modal (`1. WPR`, `2. VDRL`, `3. Vendor Document Log Report`, `4. PSR`, `5. KOM`) so every uploaded file lands in its exact designated drop location.

### 2.6. CSS Specificity & Table Styling Inheritance (`td:last-child` Leakage into Nested Tables)
- **The Challenge:** Broad descendant CSS selectors (`.data-table tbody tr td`) leak alignment and `white-space: nowrap` into nested tables or detail drawers.
- **The Solution & Best Practice:** Always scope table rules using direct child selectors (`.data-table > thead > tr > th`, `.data-table > tbody > tr > td`).

### 2.7. CSS Grid Overflow & Word-Wrapping Inside Detail Cards (`overflow-wrap: anywhere`)
- **The Challenge:** Long package names, PO numbers, or document numbers (`MM-ASK-1A-GEN01-PLC-INS-WPR-0001`) clip inside metadata cards.
- **The Solution & Best Practice:** Apply `overflow-wrap: anywhere !important; word-break: break-word !important; white-space: normal !important;` to `.kv-val`, `.alert-title`, `.alert-detail`, and table cells.

### 2.8. Handling Variations in P/F/A Terminology ("New Plan")
- **The Challenge:** ASK Procurement Plan uses `"New Plan"`, `"Forecast"`, `"Actual"` in Column K (`Matrix`), whereas Z1F Procurement Plan uses `"P"`, `"F"`, `"A"` in Column L (`P/F/A`).
- **The Solution & Best Practice:** Normalize `matrix_type.upper()` against `("NEW PLAN", "PLAN", "P")`, `("FORECAST", "F")`, and `("ACTUAL", "A")` in `parse_psr_file()`.

### 2.9–2.13. Schedule Float, Unreached Stage Guards, & Integer Durations
- Preserve actual vs. forecast date comparisons, guard downstream unreached milestones against placeholder date typos (`DD-MM-YY`, `XXXX`), and format durations/percentages cleanly.

---

### 2.14. Native `.xlsx` XML Streaming (`zipfile` + `ET.iterparse`) vs. `openpyxl` 16,384-Column Slowdown on `VDRL/` Workbooks (Nov 01, 2026)
- **The Challenge:** Loading all 23 `.xlsx` workbooks in `VDRL/` using `openpyxl.load_workbook(..., read_only=True, data_only=True)` took **54 to 82 seconds**, even when passing `max_row=110, max_col=63` to `ws.iter_rows()`.
- **Root Cause:** Several vendor-submitted VDRL Excel sheets had empty formatting applied across all 16,384 columns (`XFD`) and thousands of rows. In `openpyxl`'s `ReadOnlyWorksheet`, every `<c>` cell tag inside a `<row>` is parsed in Python before slicing `[:max_col]`.
- **The Solution & Best Practice:**
  - Implemented `_fast_load_xlsx_sheets(fpath, max_row=160, max_col=64)` in `vendor_dashboard/ingestors.py` using Python's built-in `zipfile` and `xml.etree.ElementTree.iterparse`.
  - It reads `xl/sharedStrings.xml` once, streams `<row>` elements up to `max_row`, ignores `<c>` cells where `_col_letters_to_idx(ref) >= max_col`, converts Excel serial dates (`43831..49675`) to `YYYY-MM-DD`, and breaks early when encountering 10 consecutive empty rows after row 15.
  - **Result:** Cold-start parsing across all 23 VDRL workbooks + PSR workbooks dropped from **82.4 seconds** to **~14 seconds** (and **< 0.005 seconds** on cached requests).

### 2.15. Multi-Package VDRL Workbooks & Sheet-Name Typo Resilience (`ABV`+`ACBV`, `CV`+`SRV`) (Nov 01, 2026)
- **The Challenge:** While 21 workbooks in `VDRL/` contain 1 package each, 2 workbooks contain 2 distinct packages on separate tabs:
  - `MM-ASK-1A-GEN01-ABV ACBV-INS-VDR-0001_261012.xlsx` contains `Summary (ABV)`, `Summary (ACBV)`, `REV 00 (ABC)` *(note the typo `ABC` instead of `ABV` in the tab name!)*, and `REV 00(ACBV)`.
  - `MM-ASK-1A-GEN01-CV SRV-INS-VDR-0001_261012.xlsx` contains `Summary (CV)`, `Summary (SRV)`, `REV 00 (CV)`, and `REV 00(SRV)`.
- **The Solution & Best Practice:**
  - In `parse_vdrl_workbook(fpath)`, detect when `len(summary_sheets) > 1`, extract each package code from the parenthesized tab name `Summary (<CODE>)`, and pair it with its matching `REV 00` sheet (including explicit fallback matching `ABC` $\rightarrow$ `ABV`). This cleanly yields all **25 ASK VDRL packages** from the **23 files**.

### 2.16. Dual Vendor Document Monitoring: VDRL (Vendor Submission) vs. Official Contractual D2 Log (`ASK Vendor Document Log Report` `.xls`) (Nov 01, 2026)
- **The Challenge:** Project management requires monitoring two distinct aspects of vendor documentation:
  1. **Vendor Document Submission (`VDRL/*.xlsx`):** Tracks whether the vendor has submitted documents on time against the VDRL schedule (`SUBMITTED`, `NOT YET`, `OVERDUE`, `CLOSED`, Key vs. Non-Key, Production Impact, Fabrication/Sub-order Impact).
  2. **Official & Contractual Client Review & Return (`ASK Vendor Document Log Report - 25 Sep 26.xls`):** Exported from D2 in legacy binary `.xls` (OLE2 Compound Document) format (`271` submittals), tracking contractual Client/Contractor review pipeline (`In Review`, `Submitted`, `In Approval`) and Return Status Codes (`APPR`, `AWC`, `APPR-C`, `REC`, `Pending Return`), Reviewer/Focal Point Plan vs. Actual Finish dates, and Transmittal Out numbers.
- **The Solution & Best Practice:**
  - Use `pandas.read_excel(..., header=None)` with `xlrd>=2.0.1` to parse `.xls` binary files alongside `.xlsx` updates.
  - Render **Element 6 (`Status Vendor Document`)** as a dual side-by-side view (`6A. Vendor Document Submission [VDRL]` and `6B. Official Client Review & Return [Vendor Document Log Report]`) plus a dedicated project-wide **Tab 3 (`Contractual Client Review & VDRL Hub`)**.

### 2.17. Dynamic Source Priority Resolution (`WPR Priority #1 > PSR Priority #2`) & Development Sample Template Mode (`PLC WPR`) (Nov 01, 2026)
- **The Challenge:** During execution, the Procurement Plan (`PSR`) may not be as up-to-date as the latest weekly/bi-weekly `WPR` uploaded by the package vendor. Meanwhile, during the dashboard development phase, only the `PLC` WPR sample (`MM-ASK-1A-GEN01-PLC-INS-WPR-0001 WEEKLY PROGRESS REPORTS REV 00.pdf`) is in `WPR/` for ASK, and all other packages will use an identical WPR template.
- **The Solution & Best Practice:**
  1. **Priority Rule (`_reconcile_milestones`):** Whenever a live WPR exists for a package, **WPR is Priority #1** and overrides PSR on any deviation (while displaying `PSR Plan: YYYY-MM-DD` as a secondary deviation note). If no WPR information exists for a package, **PSR is Priority #2 (Fallback)**.
  2. **Switchable Development Sample Mode (`mode='sample'` vs. `mode='live'`):** In `Dev Mode (PLC WPR Sample)`, packages without a live WPR in `WPR/` inherit the `PLC` WPR template structure for their milestone progress breakdown, S-Curve, and look-ahead template (clearly badged as `Dev Sample: PLC WPR + PSR`), while in `Execution Mode (Live WPR > PSR)` only packages with a live uploaded WPR show WPR progress curves.

### 2.18. Strict `KOM/` Manufacturer Detail Rule (`.pdf` & `.pptx` Extraction vs. Blank Fallback) (Nov 01, 2026)
- **The Challenge:** Per user domain specification: *"The detail of manufacturer are in the KOM. If it is not available, leave it blank."* Even though `PLC` has a vendor name (`ROCKWELL AUTOMATION`) inside its WPR PDF, `PLC` does not yet have a KOM file in `KOM/`.
- **The Solution & Best Practice:**
  - `parse_kom_folder('KOM')` extracts Manufacturer Name, Vendor Code, Factory Location, KOM Date, PO Reference, and Key Contacts strictly from `.pdf` (`pymupdf`) and `.pptx` (`zipfile` slide XML) files inside `KOM/`.
  - If a package has no matching file in `KOM/`, `ManufacturerDetail(available=False)` leaves all Manufacturer Detail fields strictly blank (`—`) in the UI, never leaking WPR or PSR vendor hints into Element 3.

### 2.19. Project Master Schedule (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf` Rev. B1) Priority Baseline & Critical Path (`TF=0`) Multi-Tier Risk Modeling (Nov 01, 2026)
- **The Challenge:** Tracking vendor delivery solely against fixed PSR ROS dates ignored the official Project Master Schedule (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf`, 97 pages, `AWC by Apisit Tangwashirakul`), which governs contractual milestone Key Dates (Key Date #2 Delivery to Worksite: `2027-06-02`, Topside Main Deck Lift: `2027-04-10`, Pre-Commissioning Start: `2027-04-21`), package-specific Procurement activities (`ASK-PR-*`), yard topside erection dates (`ASK-CN-*`), and zero-float critical paths (`TF=0`).
- **The Solution & Best Practice:**
  - Implemented `parse_master_schedule_pdf()` and `match_master_schedule_entry()` in `vendor_dashboard/ingestors.py` to ingest the Master Schedule in Priority #1 over PSR ROS dates.
  - Automatically identifies whether an equipment item sits on the project critical path (`PDC`, `DTP`, `CV`, `ACV`, `PLT`) and checks for Yard Erection clashes (e.g., Pedestal Crane delivery `2027-06-23` vs. Upper Deck Erection `2027-06-05`).

### 2.20. One-Picture Proportional Health Bar & Smart Issue Identification Matrix (Nov 01, 2026)
- **The Challenge:** Executive stakeholders require an instant "One-Picture" status view to know which packages are already delayed, high risk, low risk, or healthy, without scrolling or sifting through complex tables.
- **The Solution & Best Practice:**
  - Designed a proportional 4-segment health bar with 4 interactive tier cards (`🔴 Already Delayed`, `🟠 High Risk Delay`, `🟡 Low Risk`, `🟢 Very Good / Healthy`).
  - Coupled each tier with clickable package pills and a single-window **Smart Issue Identification & Timely Mitigation Action Matrix** that automatically diagnoses the root cause (e.g., Master Plan delivery breach, Yard Erection conflict, WPR Att. 4 sub-supplier hold, Key VDRL gate) and prescribes concrete mitigation actions, deadlines, and owners.

### 2.21. Reverting Governing Delivery Baseline to Live Procurement Plan (Nov 01, 2026)
- **The Challenge:** In Master Plan PDF Rev. B1 draft baseline, certain procurement delivery dates (such as Chemical Injection Pumps `CIP`) had not been updated to reflect the latest engineering baseline, resulting in a false-positive delivery delay warning (`+217d`).
- **The Solution & Best Practice:** Reverted governing ROS date reference strictly back to the live Procurement Plan (`Attachment 2-Procurement Plan-ASK...`), keeping Master Plan strictly informational for Critical Path (`TF=0`) and Yard Erection windows.

### 2.22. Seamless Dynamic Weekly Drop Architecture Across All 5 Document Sources (Nov 01, 2026)
- **The Challenge:** Project files arrive on an ongoing weekly basis across dynamic locations (`WPR/`, `VDRL/`, `KOM/`, Root).
- **The Solution & Best Practice:** Implemented `max(mtime, ctime)` freshness detection, ascending freshness sorting to supersede older weekly revisions cleanly, dual-tier cache busting (`_compute_workspace_fingerprint` + `_INGEST_CACHE.clear()`), and smart filename-pattern upload routing. Validated via automated trial test `test_dynamic_weekly_file_updates_seamless()`.

### 2.23. Kick-Off Meeting (KOM) PowerPoint-to-PDF Conversion & Multi-Page Slide Extraction (Nov 01, 2026)
- **The Challenge:** The user converted all PowerPoint (`.pptx`) presentations in `KOM/` into portable document format (`.pdf`) files. The system required updating to ensure that:
  1. Converted `.pdf` files are parsed with 100% accuracy for manufacturer details (`Manufacturer Name`, `Vendor Code`, `Factory Location`, `KOM Date`, `PO No.`, and `Key Contacts`).
  2. If any legacy `.pptx` files remain in `KOM/`, the newly converted `.pdf` version automatically supersedes the `.pptx` version so that dashboard source links point to `.pdf`.
  3. Multi-page slide decks (up to 50–60 pages) are thoroughly scanned without truncating contact info or PO numbers on later slides.
- **The Solution & Best Practice:**
  - Expanded `_extract_kom_text` in `vendor_dashboard/ingestors.py` to extract up to 50 pages from `.pdf` decks via PyMuPDF (`pymupdf`), which executes in milliseconds per document.
  - In `parse_kom_folder()`, implemented an explicit priority rule: if a package matches both a `.pdf` and a `.pptx` file, `.pdf` unconditionally takes precedence.
  - Added dual mapping for `PG` and `TG` from `KOM_TG and PG_ASK.pdf`.
  - Preserved `.pptx` fallback support in `server.py` and `ingestors.py`.
  - Verified across all 17 packages (`ABV`, `ACBV`, `PDC`, `CIP`, `PV`, `TSP`, `PSP`, `SKT`, `TANK`, `CDS`, `DGS`, `IGF`, `IGFIL`, `SOLR`, `PLS`, `PSV`, `PG`) in automated test `test_kom_manufacturer_rule()`, confirming all 17 packages now govern from `.pdf` sources while packages without KOM (e.g. `PLC`) strictly remain blank (`—`).

### 2.24. GitHub Web UI 25 MB Upload Limitation & Local Git Clone Workflow (Nov 01, 2026)
- **The Challenge:** When uploading the converted PDF files to GitHub via the web browser drag-and-drop interface (`github.com/.../upload`), GitHub rejects the upload with:
  > *"The file is too large and cannot be uploaded. Consider creating the file in a local clone and pushing it to GitHub."*
- **Root Cause:**
  - GitHub's Web UI enforces a strict **25 MB limit** per single file and per batch commit upload.
  - With multi-page presentation PDFs in `KOM/` and project schedules, the total workspace reaches **~75 MB**. Attempting to upload everything in a single web browser drag-and-drop operation exceeds this 25 MB boundary.
- **The Solution & Best Practice:**
  1. **Git CLI Workflow (Recommended):** Initialize Git locally in `Vendor Tracking` and push via terminal (`git push origin main`). Git CLI supports up to **100 MB per file** (our largest file is only 6.8 MB), eliminating all web UI batch limitations and enabling fast 1-command weekly updates.
  2. **Repository Slimming:** Remove unneeded duplicates (e.g. `KOM for Pressure Vessels & Tanks (Updated_23Jul26) (1).pdf` [5.0 MB]) and project-level decks that the code explicitly skips (`PTTEP ASK KOM Wip1.3.pdf` [6.8 MB], `D26600_ASK Project KOM50326.pdf` [3.3 MB]), reducing the folder size by over 15 MB.
  3. **Batch Web Uploading (Fallback):** If uploading via browser web interface, commit in small folder-by-folder batches (under 25 MB each) rather than dragging all folders simultaneously.

### 2.25. Replacing Heavy `KOM/` PDF Decks with a Lightweight `Export Data File.xlsx` (Nov 02, 2026)
- **The Challenge:** The user wants to keep uploading through the GitHub web UI but avoid the ~38 MB of `KOM/` presentation PDFs. Only ~10 essential fields per package are actually used by the dashboard (Element 3 Manufacturer Detail).
- **The Solution & Best Practice:**
  - **Extract once, upload the extract:** `python export_kom_data.py` parses `KOM/` and writes `Export Data File.xlsx` (sheet `KOM_Manufacturer` + `README`): Package Code, Manufacturer Name, Vendor Code, Factory Location, KOM Date, PO No., PO Date, Equipment Scope, Key Contacts, KOM Source File. **18 packages = 8.3 KB** (vs ~38 MB of PDFs, a >4,000× reduction).
  - **Dual-source reader (`parse_kom_folder` in `ingestors.py`):** Excel rows form the baseline; live `KOM/` decks (local PC only) override per package. On Render, where `KOM/` is not uploaded, Manufacturer Details come entirely from the Excel file — identical output, verified field-by-field.
  - **Auto-sync:** When `KOM/` exists locally and yields new or changed data, the dashboard rewrites `Export Data File.xlsx` automatically (atomic `.tmp` + `os.replace`; silently skipped if the file is open in Excel). Users can also hand-edit or add rows (Package Code + Manufacturer Name are required); header matching is by label, so column order doesn't matter.
  - **Cache key** includes the Excel file's `(mtime, ctime, size)` signature, and the root folder is already in the server fingerprint, so replacing the Excel file triggers an immediate refresh.
  - **`.gitignore`** now excludes `KOM/` and `*.tmp.xlsx`.
  - **Regression test** `test_kom_export_data_file_fallback()` simulates a deployment with no `KOM/` folder and asserts every package's details match the live KOM parse exactly, and that `PLC` stays blank.
- **Remaining upload note:** Without `KOM/`, the project is **38.3 MB (49 files)**, and `VDRL/` alone is **28.1 MB**, which is still over the 25 MB per-commit web limit. So `VDRL/` must be uploaded in **two halves** (~14 MB each), or use the Git CLI (§2.24).

### 2.26. WPR & MOM Ingestion Architecture via Lightweight Excel Intermediate Store (Nov 03, 2026)
- **The Challenge:** Over a 40-week project duration across 100+ packages, collecting Weekly Progress Reports (`WPR/`) and Minutes of Meeting (`MOM/`) yields ~8,000 PDFs totalling 16+ GB. Storing these raw PDFs in Git or deploying them to Render.com is impossible due to repository size ceilings and web service RAM limits.
- **The Solution & Best Practice:**
  - Standardized an intermediate tabular store: `Export Data File.xlsx` with specialized sheets (`WPR_Overview`, `WPR_Gantt_Schedules`, `MOM_Action_Items`, `MOM_Discussions`).
  - Implemented high-throughput streaming text and table extractors in PyMuPDF (`fitz`) that extract milestone progress, S-curve curves, sub-order material gates, and MOM action registers in milliseconds.
  - Compresses ~16 GB of raw document mass into ~3 MB of structured Excel records (99.98% volume compression) while retaining 100% of analytical fidelity.
  - Git tracks only the lightweight `Export Data File.xlsx`, completely excluding raw `WPR/` and `MOM/` directories via `.gitignore`.

### 2.27. In-Dashboard One-Click Sync & Download Portal (Nov 03, 2026)
- **The Challenge:** Non-technical project coordinators and discipline engineers need to update project records weekly without opening a command-line terminal or executing Python scripts manually.
- **The Solution & Best Practice:**
  - Added an in-dashboard synchronization suite: an interactive file drop-zone, a `🔄 Run Extraction & Sync Excel` button invoking `/api/sync-all`, and a direct `📥 Download Export Data File` button linking to `/api/download-export-file`.
  - The server handler runs the extraction pipeline asynchronously, atomic writes the updated Excel workbook, invalidates memory caches, and returns a JSON operational summary with HTTP 200.

### 2.28. Master Schedule (MM-ASK-1A-GEN01-PJM-SCD-0001) Governing ROS Resumption & Dual Baseline Display (Nov 03, 2026)
- **The Challenge:** The project governance requires resuming the primary Required On Site (ROS) baseline date to the Project Master Schedule (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf`), while retaining the live Procurement Plan (`Attachment 2-Procurement Plan-ASK...`) as a fallback and cross-comparison baseline.
- **The Solution & Best Practice:**
  - Implemented governing ROS resolution in `reconciler.py`:
    $$\text{Governing ROS} = \text{Master Schedule ROS (MM-ASK-1A-GEN01-PJM-SCD-0001)} \parallel \text{Procurement Plan ROS}$$
  - Preserved both explicit dates in the reconciliation output: `master_eta_date` and `psr_ros_date`.
  - Displayed dual-baseline cards on the dashboard UI: Card 6 highlights Governing ROS with source provenance, and Card 4.1C / detail cards explicitly render both Master Plan ROS and Procurement Plan ROS side-by-side with delta analysis.

### 2.29. 7-Day & 14-Day Look-Ahead Hub Across Master Schedule & Packages with Multi-Tab Excel Export (Nov 03, 2026)
- **The Challenge:** EPC management requires proactive visibility into activities starting or completing in the immediate 7-day and 14-day operational horizons across both the Master Schedule (EPC, engineering, civil, topside erection) and all 100+ individual equipment packages (Gantt tasks, VDRL document submittals, material sub-orders).
- **The Solution & Best Practice:**
  - Added `_extract_master_schedule_activities()` in `ingestors.py`: extracts all 2,522 Primavera P6 activities from the 86-page Master Plan PDF, categorizing them by EPC phase and indexing Total Float (`TF=0` Critical Path).
  - Built `build_project_lookahead_hub()` in `reconciler.py`: slices activities into 7-day and 14-day look-ahead horizons, combining Scope 1 (Master Plan) and Scope 2 (Package Level).
  - Synchronized multi-tab records into `Export Data File.xlsx` under `LookAhead_7Days` and `LookAhead_14Days` with full metadata (Activity ID, Package/WBS, Description, Start/Finish, Duration, Total Float, Critical Path status, Source Document).
  - Integrated a dedicated dashboard tab **"🔮 7 & 14 Days Look-Ahead Hub"** (`#panelLookAhead`) with dynamic KPI cards, window toggle switches, scope/criticality filters, and responsive search.

### 2.30. Project-Separated Export Data Files (`Export Data File_ASK.xlsx` & `Export Data File_Z1F.xlsx`) (Nov 03, 2026)
- **The Challenge:** The Z1F project runs in parallel with ASK. Storing both projects in a single shared export workbook risks data pollution, header collisions, and confusion during package auditing.
- **The Solution & Best Practice:**
  - Separated the export intermediate database into two dedicated project workbooks:
    - **ASK:** `Export Data File_ASK.xlsx` (18 KOM packages, 599 7d / 762 14d look-ahead activities, WPR/MOM registers).
    - **Z1F:** `Export Data File_Z1F.xlsx` (10 KOM packages, 155 7d / 244 14d look-ahead activities).
  - Updated `export_dashboard_data.py` to extract, filter, and sync each project independently (`--project ASK`, `--project Z1F`, or default `--project ALL`).
  - Added query-based download routing in `server.py`: `/api/download-export-file?project=ASK|Z1F` dynamically serves the correct project workbook with proper `Content-Disposition`.
  - Frontend SPA updates the download button text and link dynamically when the user switches between `Project: ASK` and `Project: Z1F`.
  - Both files stay well below 200 KB, enabling instantaneous manual drag-and-drop web commits via the GitHub website interface from restricted working computers.

---

## 3. Summary of Pitfalls vs. Resolutions Table

| # | Pitfall Encountered | Technical Cause | Engineering Resolution |
|---|---|---|---|
| 1 | Upload button white & non-functional after Git push | Stale CSS/JS cached by browser & Render CDN | Added `end_headers()` cache headers (`must-revalidate`) + `?v=YYYYMMDD` to `<link>` and `<script>` in `index.html`. |
| 2 | `SimpleHTTPRequestHandler` 404 error on query URLs | `self.path` containing `?v=...` failing path translation | Used `urlparse(self.path).path` before routing or delegating to `super().do_GET()`. |
| 3 | File overwrite / ambiguity across weekly updates | Static file names replacing previous week's history | Dynamic mtime sorting and designated folder routing (`WPR/`, `VDRL/`, `KOM/`, Root). |
| 4 | Uploading file without immediate dashboard update | Data cache locked/stale until server restart | Added `_compute_workspace_fingerprint()` + `_SNAPSHOT_CACHE.clear()` on `/api/upload`. |
| 5 | Detail cards right-aligned & stage headers misaligned | Outer `.data-table` rules leaked into nested detail cells | Scoped outer rules with direct child selectors (`.data-table > tbody > tr > td`). |
| 6 | Long MR numbers, POs & package names cut off | Inherited `white-space: nowrap` + `overflow: hidden` | Added `overflow-wrap: anywhere !important; white-space: normal !important;` to `.kv-val` and table cells. |
| 7 | Dashboard showing 0 packages after ASK PSR upload | Source Excel changed terminology from `"Plan"` to `"New Plan"` | Normalized `("NEW PLAN", "PLAN", "P")` in `parse_psr_file()`. |
| 8–13 | Schedule float, past date typos on future stages, decimals | Static Excel dates & floating-point artifacts | Implemented dynamic float, downstream stage guards, and clean rounding. |
| 14 | `VDRL/` parsing took 54–82s with `openpyxl` `read_only=True` | VDRL sheets formatted across 16,384 columns (`XFD`) parsed cell-by-cell | Built `_fast_load_xlsx_sheets()` using `zipfile` + `ET.iterparse` with column & empty-row cutoffs (<15s cold, <5ms cached). |
| 15 | 23 VDRL files contained 25 packages + tab typo `REV 00 (ABC)` | `ABV`+`ACBV` and `CV`+`SRV` combined in single workbooks | Added multi-sheet pair detector in `parse_vdrl_workbook()` with `ABC` $\rightarrow$ `ABV` alias handling. |
| 16 | `VDRL` Summary showed `0` overdue on initial pass | Used 1-based column indices instead of 0-based indices (`r[0]` = Category, `r[4]` = Overdue) | Fixed 0-based indexing in `_parse_vdrl_summary_rows()` and `_parse_vdrl_rev_rows()`; verified `PLC` = `52` total, `3` submitted, `16` overdue. |
| 17 | Official Client Log (`ASK Vendor Document Log Report - 25 Sep 26.xls`) is binary `.xls` | `openpyxl` only reads `.xlsx` | Parsed `.xls` via `pandas.read_excel()` + `xlrd>=2.0.1` (`271` submittals, `138` In Review, `100` AWC, `27` APPR). |
| 18 | Windows console `UnicodeEncodeError` on PPTX zero-width spaces `\u200b` | Default `cp1252` stdout encoding on Windows | Used `sys.stdout.reconfigure(encoding='utf-8')` and UTF-8 JSON responses (`ensure_ascii=False`). |
| 19 | Delivery tracked only vs PSR ROS without Project Master Plan | Master Schedule (`MM-ASK-1A-GEN01-PJM-SCD-0001.pdf`) is the primary baseline | Added `parse_master_schedule_pdf()` priority tracking + yard erection & critical path (`TF=0`) clash detection. |
| 20 | Executive review required opening multiple tabs and scrolling | No consolidated high-level risk visualization | Implemented One-Picture Proportional Bar, 4-Tier Cards, and Smart Mitigation Action Matrix with single-window layout. |
| 21 | Master Plan PDF Rev. B1 had un-updated procurement dates causing false delivery delays | Master Schedule PDF procurement dates were draft Rev. B1 baseline and not updated | Reverted governing ROS date reference strictly back to live Procurement Plan (`.xlsx`), retaining Master Schedule strictly as informational context (Critical Path `TF=0` & Yard Topside Erection window). |
| 22 | Weekly file updates dropped into folders without manual configuration | Multiple files, drag-and-drop mtime preservation on Windows, and stale server caches | Implemented `max(mtime, ctime)` freshness detection, ascending freshness sorting to supersede older weekly revisions cleanly, dual-tier cache busting (`_compute_workspace_fingerprint` + `_INGEST_CACHE.clear()`), and smart filename-pattern upload routing. Validated via automated trial test `test_dynamic_weekly_file_updates_seamless()`. |
| 23 | PowerPoint KOM presentations converted to PDF format | Multi-page slide-to-page PDFs (up to 61 pages) and legacy `.pptx` co-existence | Expanded `_extract_kom_text()` to 50 pages via PyMuPDF; added `.pdf` > `.pptx` priority override in `parse_kom_folder()`; added dual `PG`/`TG` mapping; verified all 17 packages govern from `.pdf` sources. |
| 24 | GitHub Web UI drag-and-drop rejects total upload (>25 MB) | GitHub Web UI enforces strict 25 MB per-file and per-batch commit limit | Initialized local Git repository workflow (`git push`); Git CLI supports up to 100 MB per file without batch limitations, bypassing web browser upload bottlenecks. |
| 25 | `KOM/` PDF decks (~38 MB) too heavy to upload, yet only ~10 fields are used | Dashboard only needs Manufacturer Detail fields, not full slide decks | `export_kom_data.py` → `Export Data File.xlsx` (8.3 KB); `parse_kom_folder()` reads Excel as baseline + live `KOM/` override + auto-sync; `KOM/` git-ignored; verified by `test_kom_export_data_file_fallback()`. |
| 26 | 100+ packages × 40 weeks of WPR & MOM (~8,000 PDFs, ~16 GB) cannot be uploaded | Exceeds GitHub and Render limits, causing crash on boot | PyMuPDF extraction engine extracts Gantt, S-curves, sub-orders, and actions into `Export Data File.xlsx` (~3 MB total, 99.98% compression); `WPR/` and `MOM/` git-ignored. |
| 27 | Manual terminal scripts needed for weekly updates | Friction for non-developer users | Built in-dashboard drag-and-drop ingestion portal with one-click extraction (`/api/sync-all`) and direct Excel download (`/api/download-export-file`). |
| 28 | Resumed Master Schedule (`MM-ASK-1A-GEN01-PJM-SCD-0001`) as governing ROS reference | EPC governance required alignment with official Master Plan baseline while retaining Procurement Plan fallback | Reconciler dynamically evaluates Master Schedule ROS first, falling back to Procurement Plan; UI displays both explicitly side-by-side. |
| 29 | Proactive short-term look-ahead tracking across Master Plan & 100+ packages | No unified cross-project view of immediate 7-day and 14-day upcoming deliverables | Built Look-Ahead Hub engine slicing 2,522 Master Schedule activities and package tasks into 7d/14d horizons; synchronized to Excel (`LookAhead_7Days`, `LookAhead_14Days`) and added interactive dashboard tab. |
| 30 | Mixing parallel project data (Z1F & ASK) in a single export file | Shared file caused risk of data pollution and confusion during package audits | Separated into project-specific workbooks (`Export Data File_ASK.xlsx` and `Export Data File_Z1F.xlsx`) with automated independent CLI sync and dynamic UI download routing. |
| 31 | Ingestion of Z1F Project Master Schedule (MM-ZTK-1F-GEN-PJM-SCD-0002 Rev. C1) and MOM (MM-ZTK-1F-GEN-PDC-MEC-MOM-0010 Rev.00) | Z1F uses 216-page Primavera schedule (4,110 activities) and different document numbering (ZTK/1F vs ASK) | Broadened package regex, added Z1F Master Schedule catalog mapping, extracted 4,110 activities with 344 (7d) & 550 (14d) lookahead deliverables, and synced 11 PDC action items directly to isolated `Export Data File_Z1F.xlsx` (78.3 KB); verified by `test_z1f_scd_and_mom_ingestion`. |
| 32 | Project Toggle & Look-Ahead Dropdown Data/Visual Alignment | When toggling between ASK and Z1F, the Master Schedule scope dropdown & subtitle retained hardcoded ASK document names, and segmented buttons had differing heights from select controls | Dynamically bound Look-Ahead scope dropdown options and subtitle to active project master document (Z1F: `MM-ZTK-1F-GEN-PJM-SCD-0002` Rev. C1 vs ASK: `MM-ASK-1A-GEN01-PJM-SCD-0001` Rev. B1); unified heights of all segmented controls, dropdowns, inputs, and action buttons to 34px with box-sizing border-box; and scoped `quickPackageSelect` visibility to Tab 1 to eliminate layout wrapping. |
| 33 | Release to Live Execution Mode (`mode="live"`), CIP WPR-0010 Gantt ETA Priority (`2027-02-23` vs `2027-02-08`), & CIP MOM-0010 Integration | 1) MS Project Gantt dates (`Tue 2/23/27`) used `M/D/YY` format while `_parse_date` assumed `D/M/YY`; 2) Section 5 Milestone #10 listed `2027-02-08` (Loading/Departure) while Gantt Task 70 (`4.2 Sea Freight / Transit to China Port (CIF)`) completed on `2027-02-23`; 3) Client-side normalizer in `app.js` previously overwrote `pkg.eta_date` with Milestone #11; 4) Packages without WPR/MOM needed to show only SCD/PSR/VDRL in Live Mode | 1) Added `mdy_first=True` to `_parse_date` for MS Project Gantt tables; 2) Prioritized Gantt Chart shipment/CIF completion (`2027-02-23`) as governing Package ETA (`+10d buffer` vs Master Schedule ROS `2027-03-05`) and added prominent **+15d Milestone vs. Gantt Discrepancy Alert Banner**; 3) Integrated `MOM-0010` (`48 action items`, `14 sub-POs`, `28-week lead time / +6w impact`) into Tab 1 & Element 9; 4) Added `Ref Doc` & `Cut-off Date` stamps to all 9 Elements and an **Incorporated Uploaded Files Banner** at the top of the dashboard. |
| 34 | WPR Cut-Off Date Governed by Cover `REPORT DATE:`, Candidate Revision Sorting, & Look-Ahead Hub Project Binding | 1) When user uploaded newer PLC WPR Rev 02 (`29-OCT-2026`, `WPR-0001-3.pdf`), dashboard picked older Rev 00 (`16-OCT-2026`) due to obsolete dev filter and file mtime sorting; 2) Look-Ahead Schedule Register static HTML and dropdown option previously retained hardcoded ASK master schedule name when switching to Project Z1F; 3) PyMuPDF open handles prevented temp file deletion on Windows (`WinError 32`) | 1) Implemented fast header scanning of `REPORT DATE:` / `CUT OFF PERIOD:` (handling alphanumeric dates like `29-OCT-2026`) and ranked candidates by `(cut_off_date, rev_num, report_num, mtime)` so Rev 02 (`2026-10-29`, `W6`, Planned 16.0%, Actual 9.3%) strictly defeats Rev 00 (`2026-10-16`); 2) Dynamically bound Look-Ahead scope filter and subtitle to active project master schedule (`MM-ZTK-1F-GEN-PJM-SCD-0002` Rev. C1 for Z1F; `MM-ASK-1A-GEN01-PJM-SCD-0001` Rev. B1 for ASK); 3) Explicitly closed PyMuPDF document handles immediately after text extraction. |

---

## 4. Operational Reminders for Deployment & Maintenance

1. **Automated Test Suite (`tests/test_parsers.py`):** Run `python tests/test_parsers.py` after any parser or reconciler change to ensure all 15 verification suites pass.
2. **Manual Git & Render Workflow:** Do not run automated `git commit` or `git push`. The user updates GitHub and Render.com manually on a weekly basis following the guide in `Handoff_Report.md`.
