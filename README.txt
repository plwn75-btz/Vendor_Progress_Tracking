================================================================================
METHOD OF STATEMENT (MOS)
WEEKLY VENDOR PACKAGE MONITORING DASHBOARD (ASK & Z1F)
OPERATIONAL GUIDE FOR MANUAL GITHUB WEB UPLOADS & RENDER.COM DEPLOYMENT
================================================================================

DOCUMENT REFERENCE: MOS-VTRK-2026-001
REVISION: 2.1 (Dual Project: ASK & Z1F — Live WPR/MOM & Report Date Cut-Off Priority)
DATE: November 07, 2026

--------------------------------------------------------------------------------
1. OBJECTIVE & ARCHITECTURAL OVERVIEW
--------------------------------------------------------------------------------
This Method of Statement outlines the step-by-step operating procedure for 
managing, updating, and deploying the Vendor Package Monitoring Dashboard across
Aung Sinkha (ASK) and Zawtika Phase 1F (Z1F).

Because your working computer cannot use the Git CLI (no 'git commit' or 'git push')
and raw project documents comprise thousands of heavy PDFs (up to 16 GB over time),
the system uses a Decoupled Two-Tier Architecture:

  [Tier 1: Local Working Computer (Heavy Document Repository)]
  • Stores all raw PDFs: KOM presentations, WPR reports, MOM minutes, Master 
    Primavera schedules (SCD), and VDRLs.
  • Runs the local extraction engine to extract all Gantt schedules, S-curves,
    milestone progress, MOM action items, and 7/14-day look-ahead activities.
  • Upserts extracted data non-destructively into two lightweight Excel files:
      - 'Export Data File_ASK.xlsx' (~164 KB)
      - 'Export Data File_Z1F.xlsx' (~78 KB)

  [Tier 2: Cloud Web Deployment (GitHub & Render.com)]
  • Receives ONLY the lightweight Excel files and core web code.
  • Pure Cloud Fallback: The cloud application reads 'Export Data File_*.xlsx' 
    directly, delivering 100% of dashboard features with zero raw PDFs needed.
  • Bypasses GitHub Web's 25 MB upload limit and guarantees Render will never 
    crash from memory exhaustion or timeouts.

--------------------------------------------------------------------------------
2. WEEKLY FILE UPDATE PROCEDURES ON LOCAL WORKING COMPUTER
--------------------------------------------------------------------------------

----------------------------------------------------------------------
SCENARIO A: You Have New KOM, WPR, MOM, or SCD in PDF Format
----------------------------------------------------------------------
1. DROP THE FILES INTO THEIR DESIGNATED LOCAL FOLDERS:
   • New KOM Presentations (.pdf)  --> Drop into folder: MOM/ or KOM/
   • New WPR Weekly Reports (.pdf) --> Drop into folder: WPR/
   • New MOM Meeting Minutes (.pdf)--> Drop into folder: MOM/
   • New Master Schedule (.pdf)    --> Drop into root folder: Vendor Tracking/

   File Naming Conventions (Auto-detected by system):
   - ASK: MM-ASK-1A-GEN01-<PKG>-... (e.g., MM-ASK-1A-GEN01-CIP-MEC-MOM-0010.pdf)
   - Z1F: MM-ZTK-1F-GEN-<PKG>-...   (e.g., MM-ZTK-1F-GEN-PDC-MEC-MOM-0010 Rev.00.pdf)

2. EXTRACT DATA INTO EXPORT FILES (WITHOUT DELETING EXISTING DATA):
   The extraction engine uses composite primary keys:
   - WPR: (Package Code, Report No)
   - MOM: (Package Code, MOM No, Item No)
   - Gantt: (Package Code, Task ID)
   - LookAhead: (Scope, Activity ID, Window)

   When you run the sync, newly uploaded files are parsed and inserted/updated.
   All previous weeks, other packages, and historical data REMAIN 100% INTACT.

   METHOD 1 (Via Terminal / PowerShell):
   Open PowerShell in the project directory and run:
     python export_dashboard_data.py --project ALL
     (Or specify '--project ASK' or '--project Z1F')

   METHOD 2 (Via Dashboard Web Interface):
   If your local server is running ('python server.py'):
     a. Open http://localhost:8080 in your browser.
     b. Click the "⚡ Sync All WPR/MOM to Excel" button on the top bar.
     c. Click "💾 Download Export Data File (ASK)" or "(Z1F)" to save.

   RESULT:
   'Export Data File_ASK.xlsx' and 'Export Data File_Z1F.xlsx' are updated 
   with the latest metrics, action items, and look-ahead activities.

----------------------------------------------------------------------
SCENARIO B: You Have New Procurement Plan or Client Log in Excel Format
----------------------------------------------------------------------
1. DROP NEW EXCEL FILES DIRECTLY INTO THE WORKSPACE ROOT:
   • New Procurement Plan:
     Drop into root: Vendor Tracking/
     Name format: 'Attachment 2-Procurement Plan-ASK-YYYYMMDD*.xlsx'
                  'Attachment 1 - Procurement Plan-Z1F_YYYYMMDD*.xlsx'
   • New Client Review & Return Log (D2 Report):
     Drop into root: Vendor Tracking/
     Name format: 'ASK Vendor Document Log Report - DD Mon YY.xls' (or .xlsx)

2. AUTOMATIC FRESHNESS DETECTION:
   You do NOT need to delete old procurement plans or log reports!
   The system automatically inspects all matching files in root, parses the 
   date embedded in the file name, checks filesystem timestamps (max of mtime 
   and ctime), and dynamically governs from the NEWEST file. Older files serve 
   as historical backups automatically.

--------------------------------------------------------------------------------
3. STEP-BY-STEP MANUAL UPLOAD TO GITHUB WEBSITE
--------------------------------------------------------------------------------
Because you are using GitHub's website interface from your working computer:

STEP 1: Open Your GitHub Repository
   Open your browser, navigate to your GitHub project repository, and log in.

STEP 2: Click "Add file" -> "Upload files"
   On the repository home page, click the "Add file" button near the top right 
   and choose "Upload files".

STEP 3: Drag and Drop the Required Files:
   [FILES YOU MUST UPLOAD]
   ✅ Export Data File_ASK.xlsx                  (~164 KB)
   ✅ Export Data File_Z1F.xlsx                  (~78 KB)
   ✅ New Procurement Plan (if updated this week) (~1.5 MB)
   ✅ New Client Document Log (if updated this week)(~450 KB)
   ✅ Updated source code & docs for this release:
      • vendor_dashboard/extractors/wpr_mom_extractor.py
      • vendor_dashboard/ingestors.py
      • index.html & app.js
      • export_dashboard_data.py
      • tests/test_parsers.py
      • SPEC.md, Lesson_Learn.md, Handoff_Report.md, MANUAL_UPLOAD_WRAPUP.md, README.txt

   [FILES YOU MUST NEVER UPLOAD TO GITHUB]
   ❌ DO NOT upload folder KOM/   (Heavy slide decks, up to 100 MB+)
   ❌ DO NOT upload folder WPR/   (Dozens of multi-megabyte PDFs)
   ❌ DO NOT upload folder MOM/   (Meeting minutes PDFs)
   ❌ DO NOT upload folder VDRL/  (Hundreds of spreadsheet registers)
   ❌ DO NOT upload *.SCD-*.pdf   (Master Schedule PDFs: 12 MB+, 216 pages)

   Why? 
   GitHub web rejects uploads over 25 MB. The cloud server on Render does NOT 
   need those raw PDFs because all necessary data has already been extracted into 
   'Export Data File_ASK.xlsx' and 'Export Data File_Z1F.xlsx'.

STEP 4: Commit Changes
   • In the "Commit changes" box at the bottom:
     Title: "Weekly update: Synchronized ASK & Z1F export data"
   • Choose "Commit directly to the main branch".
   • Click "Commit changes".

--------------------------------------------------------------------------------
4. RENDER.COM DEPLOYMENT & VERIFICATION
--------------------------------------------------------------------------------
1. AUTOMATIC BUILD & DEPLOY:
   As soon as you commit your files on GitHub, Render.com detects the update 
   and automatically triggers a new deployment.
   Build command: pip install -r requirements.txt
   Start command: python server.py

2. CHECK LIVE DEPLOYMENT STATUS:
   • Log into your Render dashboard at https://dashboard.render.com
   • Look at your Web Service log. Within 1 to 2 minutes, you will see:
       ==> Build successful 🎉
       ==> Deploying...
       🚀 Starting Vendor Tracking Dashboard on http://0.0.0.0:8080
       ... Live WPR & MOM loaded from Export Data File_*.xlsx

3. OPEN AND VERIFY IN BROWSER:
   • Navigate to your public Render URL (e.g., https://your-app.onrender.com).
   • Perform a Hard Refresh to bypass browser cache:
       - Windows/Chrome: Press Ctrl + F5 (or Ctrl + Shift + R)
   • Verify:
       a. Switch between ASK and Z1F using the top project toggle.
       b. Open the "7 & 14 Days Look Ahead" tab to confirm upcoming activities.
       c. Check that the Executive Summary reflects the latest ROS dates.
       d. Click "💾 Download Export Data File" to confirm downloading the sheet.

--------------------------------------------------------------------------------
5. QUICK TROUBLESHOOTING & FAQ
--------------------------------------------------------------------------------
Q1: When I add a new WPR or MOM, does it erase previous data in the export file?
A1: NO. The synchronization script performs an intelligent UPSERT. Existing rows 
    are kept intact, and only new or modified items are added/updated.

Q2: Can I download the updated Excel file directly from the live web dashboard?
A2: YES. Click "💾 Download Export Data File" in the top bar of the live web app.
    It downloads the active project's file ('Export Data File_ASK.xlsx' or 'Z1F').

Q3: What if I accidentally forget to upload the new Master Schedule PDF?
A3: The Master Schedule PDF is NEVER uploaded to GitHub. Its activities, dates, 
    and critical path milestones are already saved in 'Export Data File_*.xlsx' 
    under the 'LookAhead_7Days' and 'LookAhead_14Days' sheets.

Q4: How do I test the application locally before uploading to GitHub?
A4: Run 'python server.py' in PowerShell and open http://localhost:8080.
    Run 'python -u tests/test_parsers.py' to run all 15 automated tests.
================================================================================
END OF METHOD OF STATEMENT
================================================================================

