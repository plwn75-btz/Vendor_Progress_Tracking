# Procurement Tracking Dashboard (Z1F & ASK)

An executive-level, glassmorphism web dashboard for real-time tracking, float erosion modeling, and milestone progress analysis across offshore procurement packages for:
- **Z1F (Zawtika Project Development - Topside & Jacket - WP1)**
- **ASK (Aung Sinkhla Project Development - Overall Cycle - WP2)**

---

## 1. System Architecture

The project runs as a lightweight, standalone, full-stack web application designed for high performance and zero framework bloat:

- **Backend (`server.py`):**
  - Native Python HTTP server (`http.server.SimpleHTTPRequestHandler`) with custom REST endpoints.
  - No heavyweight dependencies (no Django, Flask, or FastAPI required).
  - Excel data extraction powered by `pandas` and `openpyxl`.
  - Multi-sheet parsing across 24 milestone stages with automatic column header detection and aliasing.
  - Zero-cache response headers (`no-cache, no-store, must-revalidate`) for rapid asset propagation.
- **Frontend SPA (`index.html`, `styles.css`, `app.js`):**
  - Dark glassmorphism design system (`#0a0e1a` deep background, cyan/blue accent palettes, backdrop filters).
  - Responsive tables, interactive detail drawers, 24-stage milestone pipelines, and real-time text/stage filtering.
  - In-browser dynamic float analysis and schedule risk simulation.

---

## 2. Key Business Logic & Schedule Modeling

### 2.1. Dynamic Float & Overdue PO Schedule Erosion
Package float represents the schedule buffer between projected equipment delivery and Required on Site (ROS) date:
- **Actual Float (`Act` Badge):** For awarded contracts where `Actual PO Date` exists:
  $$\text{Actual Delivery Date} = \text{Actual PO Date} + \text{Delivery Duration (Days)}$$
  $$\text{Actual Float} = \text{Plan ROS Date} - \text{Actual Delivery Date}$$
- **Dynamic Float for Overdue POs (`Fcst*` Badge):** If a PO is unawarded and the planned PO date has passed ($\text{Forecast PO Date} < \mathbf{Today}$), static Excel numbers create false comfort. The dashboard dynamically pins the earliest possible PO award date to $\mathbf{Today}$:
  $$\text{Earliest Delivery Date} = \mathbf{Today} + \text{Delivery Duration (Days)}$$
  $$\mathbf{Dynamic\ Forecast\ Float} = \text{Plan ROS Date} - (\mathbf{Today} + \text{Delivery Duration (Days)})$$
  This erodes float day-by-day until procurement acts and awards the PO.
- **Standard Forecast Float (`Fcst` Badge):** When the forecast PO date remains in the future ($\ge \mathbf{Today}$):
  $$\text{Forecast Float} = \text{Plan ROS Date} - (\text{Forecast PO Date} + \text{Delivery Duration})$$

### 2.2. Downstream Unreached Stage Guard
In raw Excel data cuts (e.g. Cut `20261002`), planners often enter placeholder or stale dates from previous years for downstream milestones (e.g., `FAT` and `Ready for Shipment` in 2025/2026, while their baseline plan is in 2027).
- The dashboard detects the package's active stage (`currentStageIdx`).
- For any stage occurring *after* the active stage where $\text{Plan Date} > \mathbf{Today}$, past forecast dates are recognized as unreached downstream steps.
- These are locked to **`● Upcoming`** with `0d` delay, completely preventing premature downstream alarms (such as false 288-day FAT delays).

### 2.3. Dual-Status Execution & Forecast Tracking
- **Package Status (Table Header & Main Table):** Reflects the execution status of the **latest completed milestone** (`latestCompletedStage`). If completed early or on time, the package displays **`● On Track`** (Green).
- **Active Stage Forecast Slip:** If the currently active, in-progress stage is running behind schedule ($\text{Forecast Date} > \text{Plan Date}$):
  - Displays an amber **`+Nd Fcst Slip`** pill in the `Current Stage` table column.
  - Displays a **`● Forecast Slip`** status badge in the stage detail drawer and milestone pipeline.
- **Max Delay Preservation:** The `Max Delay` column (yellow tabular digits) accurately shows the maximum completed or active delay experienced by the package without being poisoned by downstream placeholders.

### 2.4. Lookahead Due Horizon (+Nd) & Clean Precision
- **Lookahead Window:** Highlights milestones falling due within 7 days ($\mathbf{Today} \le \text{Forecast Date} \le \mathbf{Today} + 7\text{d}$).
- **Due Horizon:** For milestones marked as **`● Due Soon`** (`atrisk`), the Delay column indicates remaining days until forecast:
  $$\text{Days Remaining} = \text{Forecast Date} - \mathbf{Today}$$
  (e.g., $+4\text{d}$ for a milestone due 4 days from today).
- **Clean Integer Durations:** All lead times and durations are strictly rounded to whole integer days (e.g., `90d`), preventing floating-point precision artifacts (e.g. `90.00000000000003d`).

---

## 3. Quick Start & Local Setup

### 3.1. Prerequisites
- Python 3.10+ (tested on Python 3.11)
- Modern web browser (Chrome, Edge, Firefox, Safari)

### 3.2. Installation
Clone the repository and install the minimal dependencies:
```bash
pip install -r requirements.txt
```

### 3.3. Running Locally
Start the server:
```bash
python server.py
```
By default, the server runs on port **8080** (or port specified via `PORT` environment variable):
- Open **`http://localhost:8080`** in your browser.

---

## 4. REST API Reference

| Endpoint | Method | Description |
|---|---|---|
| `/api/data` | `GET` | Returns aggregated JSON data containing summary KPIs, project packages, 24-stage milestone schedules, and lookahead items. |
| `/api/refresh` | `GET` | Forces the backend to reload and re-parse the latest Excel spreadsheets from disk. |
| `/api/upload` | `POST` | Multipart form upload endpoint for uploading new weekly cut-off Excel files (`Attachment 1-...` or `Attachment 2-...`). Automatically runs data extraction upon save. |

---

## 5. Weekly Cut Ingestion & Upload Center

1. **Web Upload:** Click **📁 Upload Center** in the top navigation bar. Drag and drop your weekly Excel cuts (`Attachment 1 - Procurement Plan-Z1F_YYYYMMDD.xlsx` or `Attachment 2-Procurement Plan-ASK-YYYYMMDD.xlsx`) directly onto the drop zone.
2. **File Placement:** Alternatively, place the new `.xlsx` files directly in the root workspace directory and trigger `GET /api/refresh` or restart `server.py`.
3. **Verification Command:**
   ```bash
   python -c "from server import extract_all_data; d = extract_all_data(); print('Z1F:', d['projects']['Z1F']['package_count'], 'ASK:', d['projects']['ASK']['package_count'])"
   ```

---

## 6. Operational Guidelines & Policy Constraints

1. **Standalone Architecture (No Power BI Execution):**
   - The primary execution vehicle is this standalone Python web application.
   - Do NOT execute or run Power BI automation scripts in `PowerBI_Solution/`.
2. **Manual Git Synchronization:**
   - All Git commits and GitHub synchronization are executed manually by the repository owner. AI agents must not run automated `git commit` or `git push` commands.
3. **Cache Busting on Deployment:**
   - Whenever updating CSS or JavaScript, bump the asset version query strings in `index.html` (e.g. `styles.css?v=YYYYMMDD_N` and `app.js?v=YYYYMMDD_N`) to prevent browser and CDN caching.

---

## 7. Project Documentation

- [Lesson_Learn.md](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Procurement%20Tracking/Lesson_Learn.md): Technical pitfalls, root causes, formula derivations, and architectural solutions across all weekly cuts.
- [Handoff_Report.md](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Procurement%20Tracking/Handoff_Report.md): Detailed engineering handoff specification, deployment instructions, and operational verification checklists.
