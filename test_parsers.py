"""Automated verification tests for the ASK & Z1F Vendor Package Monitoring Dashboard."""

import os
import sys

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from vendor_dashboard.ingestors import (
    parse_all_vdrl_files,
    parse_all_wpr_files,
    parse_client_doc_log,
    parse_kom_folder,
    parse_psr_file,
)
from vendor_dashboard.reconciler import build_project_snapshot


def test_vdrl_multi_and_single_package_parsing() -> None:
    vdrl_dir = os.path.join(ROOT_DIR, "VDRL")
    vdrls = parse_all_vdrl_files(vdrl_dir, project_filter="ASK")
    # 23 workbooks -> 25 ASK packages (including ABV+ACBV and CV+SRV)
    assert len(vdrls) == 25
    assert "ABV" in vdrls and "ACBV" in vdrls
    assert "CV" in vdrls and "SRV" in vdrls
    assert "PLC" in vdrls
    plc = vdrls["PLC"]
    assert plc["total_docs"] == 52
    assert plc["submitted"] == 3
    assert plc["overdue"] == 16
    assert plc["not_yet"] == 33
    assert len(plc["documents"]) == 52


def test_client_doc_log_parsing() -> None:
    log = parse_client_doc_log(ROOT_DIR, project="ASK")
    assert log["available"] is True
    assert log["total_submittals"] == 271
    assert log["status_counts"]["In Review"] == 138
    assert log["status_counts"]["Submitted"] == 118
    assert log["status_counts"]["In Approval"] == 15
    assert log["return_counts"]["AWC"] == 100
    assert log["return_counts"]["APPR"] == 27
    assert "CIP" in log["by_package"]
    assert log["by_package"]["CIP"]["total_submittals"] == 37


def test_kom_manufacturer_rule() -> None:
    kom_dir = os.path.join(ROOT_DIR, "KOM")
    kom_map = parse_kom_folder(kom_dir)
    # Packages with KOM files in KOM/ must have Manufacturer details populated
    for code in ("ABV", "ACBV", "PDC", "CIP", "PV", "TSP", "PSP", "SKT", "TANK", "SOLR", "PLS", "PSV", "PG"):
        assert code in kom_map
        assert kom_map[code].available is True
        assert kom_map[code].manufacturer_name != ""
        # Converted presentations must prioritize .pdf as the governing source file
        assert kom_map[code].kom_source_file.lower().endswith(".pdf"), f"{code} source file is not PDF: {kom_map[code].kom_source_file}"
    # PLC has no KOM file in KOM/ -> must NOT be in kom_map
    assert "PLC" not in kom_map


def test_kom_export_data_file_fallback() -> None:
    """Simulate GitHub/Render where KOM/ PDFs are NOT uploaded: Manufacturer Details
    must come from the lightweight project-specific Export Data File instead."""
    from vendor_dashboard.ingestors import (
        EXPORT_DATA_FILENAME,
        EXPORT_DATA_FILENAME_ASK,
        EXPORT_DATA_FILENAME_Z1F,
        _INGEST_CACHE,
        read_kom_export_file,
    )

    export_path = os.path.join(ROOT_DIR, EXPORT_DATA_FILENAME_ASK)
    if not os.path.isfile(export_path):
        export_path = os.path.join(ROOT_DIR, EXPORT_DATA_FILENAME)
    assert os.path.isfile(export_path), "Run 'python export_dashboard_data.py' to create Export Data File_ASK.xlsx"
    assert os.path.getsize(export_path) < 1024 * 1024, "Export Data File must stay small (< 1 MB)"

    live_map = parse_kom_folder(os.path.join(ROOT_DIR, "KOM"), project="ASK")
    export_map = read_kom_export_file(export_path)
    assert set(live_map) == set(export_map)

    # A sibling folder that does not exist == deployment without KOM/
    _INGEST_CACHE.clear()
    no_kom_map = parse_kom_folder(os.path.join(ROOT_DIR, "KOM__not_uploaded"), project="ASK")
    _INGEST_CACHE.clear()
    for code in ("ABV", "ACBV", "PDC", "CIP", "PV", "TSP", "PSP", "SKT", "TANK", "SOLR", "PLS", "PSV", "PG"):
        assert code in no_kom_map, f"{code} missing from Export Data File fallback"
        assert no_kom_map[code].to_dict() == live_map[code].to_dict(), f"{code} export differs from KOM/"
    assert "PLC" not in no_kom_map

    # Also verify Z1F separate export file exists and stays small
    z1f_path = os.path.join(ROOT_DIR, EXPORT_DATA_FILENAME_Z1F)
    assert os.path.isfile(z1f_path), f"Missing {EXPORT_DATA_FILENAME_Z1F}"
    assert os.path.getsize(z1f_path) < 1024 * 1024, "Z1F export file must stay small (< 1 MB)"


def test_wpr_pdf_and_docx_parsing() -> None:
    wpr_dir = os.path.join(ROOT_DIR, "WPR")
    ask_wprs = parse_all_wpr_files(wpr_dir, project_filter="ASK")
    assert "PLC" in ask_wprs
    plc = ask_wprs["PLC"]
    # User Rule: cut-off date is report date, select the latest one (2026-10-29 > 2026-10-16)
    assert plc["source_file"] == "MM-ASK-1A-GEN01-PLC-INS-WPR-0001-3.pdf"
    assert plc["cut_off_date"] == "2026-10-29"
    assert plc["report_date"] == "2026-10-29"
    assert plc["overall_planned_pct"] == 16.0
    assert plc["overall_actual_pct"] == 9.3
    assert len(plc["milestones"]) == 12
    assert len(plc["milestone_progress"]) == 17
    assert len(plc["gantt_tasks"]) >= 50

    z1f_wprs = parse_all_wpr_files(wpr_dir, project_filter="Z1F")
    assert "CCV" in z1f_wprs
    ccv = z1f_wprs["CCV"]
    assert ccv["overall_planned_pct"] == 10.0
    assert ccv["overall_actual_pct"] == 6.0


def test_reconciler_priority_and_modes() -> None:
    # 1. Sample mode (Dev phase: PLC WPR used as sample for other packages)
    snap_sample = build_project_snapshot(ROOT_DIR, project="ASK", mode="sample")
    assert snap_sample["project"] == "ASK"
    pkgs_sample = {p["package_code"]: p for p in snap_sample["packages"]}
    assert "PLC" in pkgs_sample and "ABV" in pkgs_sample

    # PLC has live WPR (Priority #1), but no KOM in KOM/ (Manufacturer blank)
    plc = pkgs_sample["PLC"]
    assert plc["has_live_wpr"] is True
    assert plc["manufacturer_detail"]["available"] is False
    assert plc["manufacturer_detail"]["manufacturer_name"] == ""
    assert plc["key_milestones"][0]["source"] == "WPR (Priority #1)"

    # ABV has KOM in KOM/ (KCON), uses PSR (Priority #2) for milestones, and PLC sample WPR in sample mode
    abv = pkgs_sample["ABV"]
    assert abv["manufacturer_detail"]["available"] is True
    assert "KCON" in abv["manufacturer_detail"]["manufacturer_name"].upper()
    assert abv["using_sample_wpr"] is True
    assert abv["wpr_active"] is True

    # 2. Live Execution mode: ABV falls back to PSR Priority #2 without sample WPR
    snap_live = build_project_snapshot(ROOT_DIR, project="ASK", mode="live")
    pkgs_live = {p["package_code"]: p for p in snap_live["packages"]}
    assert pkgs_live["ABV"]["has_live_wpr"] is False
    assert pkgs_live["ABV"]["using_sample_wpr"] is False

    # 3. Z1F separation
    snap_z1f = build_project_snapshot(ROOT_DIR, project="Z1F", mode="live")
    assert snap_z1f["project"] == "Z1F"
    z1f_pkgs = {p["package_code"]: p for p in snap_z1f["packages"]}
    assert "CCV" in z1f_pkgs
    assert z1f_pkgs["CCV"]["has_live_wpr"] is True


def test_server_snapshot_cache_and_elements() -> None:
    import server
    snap_ask = server.get_cached_snapshot(project="ASK", mode="sample", force_refresh=False)
    assert snap_ask["project"] == "ASK"
    assert snap_ask["portfolio_kpis"]["vdrl_total_docs"] == 1277
    assert snap_ask["portfolio_kpis"]["client_log_total_submittals"] == 271

    # Verify all 9 dashboard elements + new feedback items exist on every package
    for pkg in snap_ask["packages"]:
        assert "project" in pkg                     # 1. Project
        assert "package_name" in pkg                # 2. Package Name
        assert "manufacturer_detail" in pkg         # 3. Manufacturer Detail
        assert "key_milestones" in pkg              # 4. Key Milestone
        assert "milestone_progress" in pkg          # 5. Progress Tracking by Key Milestone
        assert "vdrl_status" in pkg                 # 6A. Status Vendor Document (VDRL)
        assert "client_doc_status" in pkg           # 6B. Status Vendor Document (Client Log)
        assert "scurve" in pkg                      # 7. Overall Progress (S-Curve)
        assert "look_ahead_plan" in pkg             # 8. Look Ahead Plan 7-14 Days
        assert "areas_of_concern" in pkg            # 9. Area of Concern
        assert "package_highlights" in pkg          # 4.1 Package Key Highlight Information
        assert "attachment3_summary" in pkg         # 5A. WPR Attachment 3 Gantt Summary
        assert "attachment4_summary" in pkg         # 5B. WPR Attachment 4 Supplier Material Status


def test_feedback_items_1_to_7() -> None:
    snap = build_project_snapshot(ROOT_DIR, project="ASK", mode="sample")
    pkgs = {p["package_code"]: p for p in snap["packages"]}

    # Comment #2: Project Overall S-Curve weighted across all packages
    proj_sc = snap["project_overall_scurve"]
    assert len(proj_sc["points"]) == 40
    weights_breakdown = proj_sc["package_weights_breakdown"]
    assert len(weights_breakdown) == len(snap["packages"])
    total_weight = sum(r["package_weight_pct"] for r in weights_breakdown)
    assert abs(total_weight - 100.0) < 0.2

    # Comment #3: Dynamic Element 7 (S-Curve) & Element 8 (Look-Ahead) on package change
    plc = pkgs["PLC"]
    cip = pkgs["CIP"]
    abv = pkgs["ABV"]
    # PLC has live WPR, whereas ABV dynamically computes progress from PSR & VDRL
    assert plc["has_live_wpr"] is True
    assert abv["has_live_wpr"] is False
    assert plc["overall_progress"]["actual_pct"] != abv["overall_progress"]["actual_pct"]
    assert plc["look_ahead_plan"]["wpr_next_activities"] != abv["look_ahead_plan"]["wpr_next_activities"]

    # Comment #4.1 & #6: Verify EVERY ASK package has non-empty Receipt at Worksite (ETA) and PSR ROS Date
    for pkg in snap["packages"]:
        hl = pkg["package_highlights"]
        assert hl["current_stage_summary"], f"Missing current_stage_summary for {pkg['package_code']}"
        assert hl["active_stage_summary"], f"Missing active_stage_summary for {pkg['package_code']}"
        assert hl["remaining_days_to_next_milestone"] is not None, f"Missing remaining_days for {pkg['package_code']}"
        assert hl["effective_delivery_date"] != "", f"Missing Receipt at Worksite (ETA) for {pkg['package_code']}"
        assert hl["days_to_delivery"] is not None, f"Missing days_to_delivery for {pkg['package_code']}"
        assert hl["ros_date"] != "", f"Missing PSR ROS Date for {pkg['package_code']}"

    # Verify PLC specifically tracks WPR Milestone #11 Receipt at Worksite (2027-04-08) vs Master Plan (2027-04-15) & PSR ROS Date
    plc_hl = plc["package_highlights"]
    assert plc_hl["effective_delivery_date"] == "2027-04-08", f"plc_hl={plc_hl}"
    assert plc_hl["days_to_delivery"] in (161, 174), f"plc_hl={plc_hl}"
    assert plc_hl["ros_date"] in ("2027-04-14", "2027-04-15", "2027-04-23", "2027-04-28"), f"plc_hl={plc_hl}"
    assert plc_hl["ros_margin_days"] >= 6
    assert plc_hl["is_ros_delayed"] is False

    # Comment #5: Attachment 3 (Fabrication Gantt) & Attachment 4 (Supplier Material Status)
    assert len(plc["attachment3_summary"]["phases"]) == 5
    assert len(plc["attachment4_summary"]["items"]) == 4
    assert plc["attachment4_summary"]["components_on_hold"] >= 2
    assert cip["attachment3_summary"]["situation_summary"] != plc["attachment3_summary"]["situation_summary"]

    # Comment #6: Immutable PSR ROS Date & Delivery Delay (ETA > ROS Date)
    cip_hl = cip["package_highlights"]
    assert cip_hl["ros_date"].startswith("2027-"), f"cip_hl={cip_hl}"
    assert cip_hl["effective_delivery_date"].startswith("2027-"), f"cip_hl={cip_hl}"
    assert isinstance(snap["portfolio_kpis"]["ros_delayed_packages_list"], list)

    # Comment #1 & #7: Verify frontend HTML/JS/CSS contain sorting, filtering, and Daylight Mode
    with open(os.path.join(ROOT_DIR, "index.html"), "r", encoding="utf-8") as fh:
        html = fh.read()
    assert "toggleTheme()" in html
    assert "packageHighlightBanner" in html
    assert "toggleSort(" in html
    assert "att3PhasesTableBody" in html
    assert "att4MaterialsTableBody" in html


def test_latest_5_feedback_items() -> None:
    snap = build_project_snapshot(ROOT_DIR, project="ASK", mode="sample")
    pkgs = {p["package_code"]: p for p in snap["packages"]}

    # 1. Verify PLT is "Plate" (Primary Steel - Plates) and NOT aliased/duplicated as PLT (PLC)
    breakdown = snap["client_doc_log_overview"]["package_breakdown"]
    plt_rows = [r for r in breakdown if r["raw_package_code"] == "PLT"]
    assert len(plt_rows) == 1, f"Expected exactly 1 PLT row in Client Log breakdown, got {plt_rows}"
    assert plt_rows[0]["package_code"] == "PLT"
    assert plt_rows[0]["package_name"] == "Plate"
    assert plt_rows[0]["total_submittals"] == 27
    assert "PLT" in pkgs
    assert "Plate" in pkgs["PLT"]["package_name"]

    # 2. Verify newly dropped PSR file ("Attachment 2-Procurement Plan-ASK - 20260626.xlsx") is automatically picked up
    assert snap["sources_summary"]["psr_file"] == "Attachment 2-Procurement Plan-ASK - 20260626.xlsx"

    # 3. Verify all equipment from PSR (e.g. Barred Tee, Corrosion Coupon, Sight Glass) are included in Tab 1 packages
    assert len(snap["packages"]) >= 100
    all_names_upper = " | ".join(p["package_name"].upper() for p in snap["packages"])
    assert "BARRED TEE" in all_names_upper
    assert "CORROSION" in all_names_upper and "COUPON" in all_names_upper
    assert "SIGHT GLASS" in all_names_upper

    # 4. Verify Executive Summary at top of Tab 1 exists and has all 3 pillars
    exec_sum = snap["executive_findings_summary"]
    assert "schedule_and_ros_finding" in exec_sum
    assert "vendor_management_finding" in exec_sum
    assert "document_control_finding" in exec_sum

    with open(os.path.join(ROOT_DIR, "index.html"), "r", encoding="utf-8") as fh:
        html = fh.read()
    assert "tab1ExecutiveSummaryCard" in html
    assert "tab1EquipmentDropdown" in html
    assert "portfolioMatrixTable" in html

    # 5. Verify ROS delayed list uses Package Name (not code like PIP-093) and dropdown shows only Package Name
    delayed_names = snap["portfolio_kpis"]["ros_delayed_packages_list"]
    for d_name in delayed_names:
        assert not d_name.startswith("PIP-0"), f"Expected package name instead of code, got {d_name}"
    with open(os.path.join(ROOT_DIR, "app.js"), "r", encoding="utf-8") as fh:
        js = fh.read()
    assert '>${esc(p.package_name)}</option>' in js


def test_procurement_plan_ros_priority_and_one_picture_executive_summary() -> None:
    snap = build_project_snapshot(ROOT_DIR, project="ASK", mode="sample")
    exec_sum = snap["executive_findings_summary"]

    # 1. Verify Master Schedule is the governing ROS reference together with Procurement Plan
    assert "Master Schedule ROS" in exec_sum["headline"]
    assert "Procurement Plan" in exec_sum["headline"]
    assert "Attachment 2-Procurement Plan-ASK" in exec_sum["ros_data_reference"]
    assert "MM-ASK-1A-GEN01-PJM-SCD-0001" in exec_sum["ros_data_reference"]
    assert snap["sources_summary"]["psr_file"] == "Attachment 2-Procurement Plan-ASK - 20260626.xlsx"

    # Master Schedule is preserved as governing Priority #1 reference
    assert exec_sum["active_master_schedule_file"] == "MM-ASK-1A-GEN01-PJM-SCD-0001.pdf"
    assert snap["sources_summary"]["master_schedule_file"] == "MM-ASK-1A-GEN01-PJM-SCD-0001.pdf"

    # 2. Verify One-Picture 4-Tier Health breakdown exists for both Core and All packages
    tiers_core = exec_sum["tiers_core"]
    tiers_all = exec_sum["tiers_all"]
    for tg in (tiers_core, tiers_all):
        assert tg["high_risk"]["count"] >= 5
        assert tg["low_risk"]["count"] >= 4
        assert tg["healthy"]["count"] >= 5
        total_check = (
            tg["already_delayed"]["count"]
            + tg["high_risk"]["count"]
            + tg["low_risk"]["count"]
            + tg["healthy"]["count"]
        )
        assert total_check == tg["total"]

    # 3. Verify Coat is detected as ALREADY_DELAYED against Procurement Plan ROS Date (2026-11-23 vs ETA 2027-04-08)
    delayed_codes = {p["package_code"] for p in tiers_all["already_delayed"]["packages"]}
    assert "Coat" in delayed_codes or "PIP-093" in delayed_codes or snap["portfolio_kpis"]["ros_delayed_packages_list"] == ["Coat"]

    # 4. Verify Smart Issue & On-Time Mitigation Measure are populated for all packages
    for p in snap["packages"]:
        hl = p["package_highlights"]
        assert hl["health_tier"] in ("ALREADY_DELAYED", "HIGH_RISK", "LOW_RISK", "HEALTHY")
        assert hl["smart_issue"] != ""
        assert hl["smart_mitigation"] != ""
        assert hl["mitigation_deadline"] != ""


def test_dynamic_weekly_file_updates_seamless() -> None:
    """Trial test simulating weekly dynamic file updates across all dynamic document sources:
    1. Dynamic Procurement Plan (.xlsx) drop detection & automatic headline update
    2. Dynamic Live WPR (.pdf/.docx) drop & automatic priority switch (Priority #2 -> Priority #1)
    3. Dynamic VDRL (.xlsx) drop & revision detection
    4. Dynamic Client Document Log (.xls/.xlsx) drop & submittal tracking
    5. Automatic workspace fingerprinting & server cache invalidation
    Ensures all temporary files are cleanly unlinked in try...finally.
    """
    import shutil
    import time
    from server import _compute_workspace_fingerprint, get_cached_snapshot
    from vendor_dashboard.ingestors import _INGEST_CACHE

    created_temp_files = []
    temp_psr = os.path.join(ROOT_DIR, "Attachment 2-Procurement Plan-ASK - 20261108 R4.xlsx")
    temp_wpr_99 = os.path.join(ROOT_DIR, "WPR", "MM-ASK-1A-GEN01-CIP-MEC-WPR-0099.pdf")
    temp_wpr_mvg = os.path.join(ROOT_DIR, "WPR", "MM-ASK-1A-GEN01-MVG-MEC-WPR-0001 WEEKLY PROGRESS REPORTS REV 00.pdf")
    temp_vdrl = os.path.join(ROOT_DIR, "VDRL", "MM-ASK-1A-GEN01-PLC-INS-VDR-0001 261019.xlsx")
    temp_vdlr = os.path.join(ROOT_DIR, "ASK Vendor Document Log Report - 02 Oct 26.xls")
    for f in (temp_psr, temp_wpr_99, temp_wpr_mvg, temp_vdrl, temp_vdlr):
        if os.path.exists(f):
            try:
                os.remove(f)
            except OSError:
                pass
    _INGEST_CACHE.clear()

    try:
        # 1. Baseline fingerprint check
        fp_initial = _compute_workspace_fingerprint()
        assert len(fp_initial) > 0

        # 2. Trial: Dropping a newer Procurement Plan file
        base_psr = os.path.join(ROOT_DIR, "Attachment 2-Procurement Plan-ASK - 20260626.xlsx")
        temp_psr = os.path.join(ROOT_DIR, "Attachment 2-Procurement Plan-ASK - 20261108 R4.xlsx")
        shutil.copyfile(base_psr, temp_psr)
        created_temp_files.append(temp_psr)

        future_time = time.time() + 120
        os.utime(temp_psr, (future_time, future_time))

        _INGEST_CACHE.clear()
        psr_data = parse_psr_file(ROOT_DIR, project="ASK")
        first_pkg = next(iter(psr_data.values()))
        assert first_pkg["source_file"] == "Attachment 2-Procurement Plan-ASK - 20261108 R4.xlsx"

        snap_new_psr = build_project_snapshot(ROOT_DIR, project="ASK", mode="sample")
        assert snap_new_psr["sources_summary"]["psr_file"] == "Attachment 2-Procurement Plan-ASK - 20261108 R4.xlsx"
        assert "Attachment 2-Procurement Plan-ASK - 20261108 R4.xlsx" in snap_new_psr["executive_findings_summary"]["headline"]

        # 3. Trial: Dropping a newer Live WPR for CIP (WPR-0099 dynamic file update)
        base_wpr = os.path.join(ROOT_DIR, "WPR", "MM-ASK-1A-GEN01-CIP-MEC-WPR-0010.pdf")
        temp_wpr = os.path.join(ROOT_DIR, "WPR", "MM-ASK-1A-GEN01-CIP-MEC-WPR-0099.pdf")
        shutil.copyfile(base_wpr, temp_wpr)
        created_temp_files.append(temp_wpr)
        os.utime(temp_wpr, (future_time + 10, future_time + 10))

        _INGEST_CACHE.clear()
        wpr_map = parse_all_wpr_files(os.path.join(ROOT_DIR, "WPR"), project_filter="ASK")
        assert "CIP" in wpr_map
        assert wpr_map["CIP"]["source_file"] == os.path.basename(temp_wpr)

        snap_live = build_project_snapshot(ROOT_DIR, project="ASK", mode="live")
        cip_pkg = next((p for p in snap_live["packages"] if p["package_code"] == "CIP"), None)
        assert cip_pkg is not None
        assert cip_pkg["has_live_wpr"] is True
        assert cip_pkg["data_priority_badge"] == "Priority #1: Live WPR"

        # 4. Trial: Dropping an updated VDRL file
        base_vdrl = os.path.join(ROOT_DIR, "VDRL", "MM-ASK-1A-GEN01-PLC-INS-VDR-0001 261012.xlsx")
        temp_vdrl = os.path.join(ROOT_DIR, "VDRL", "MM-ASK-1A-GEN01-PLC-INS-VDR-0001 261019.xlsx")
        shutil.copyfile(base_vdrl, temp_vdrl)
        created_temp_files.append(temp_vdrl)
        os.utime(temp_vdrl, (future_time + 20, future_time + 20))

        _INGEST_CACHE.clear()
        vdrl_map = parse_all_vdrl_files(os.path.join(ROOT_DIR, "VDRL"), project_filter="ASK")
        assert "PLC" in vdrl_map
        assert vdrl_map["PLC"]["source_file"] == os.path.basename(temp_vdrl)

        # 5. Trial: Dropping a newer Client Document Log Report
        base_vdlr = os.path.join(ROOT_DIR, "ASK Vendor Document Log Report - 25 Sep 26.xls")
        temp_vdlr = os.path.join(ROOT_DIR, "ASK Vendor Document Log Report - 02 Oct 26.xls")
        shutil.copyfile(base_vdlr, temp_vdlr)
        created_temp_files.append(temp_vdlr)
        os.utime(temp_vdlr, (future_time + 30, future_time + 30))

        _INGEST_CACHE.clear()
        client_log = parse_client_doc_log(ROOT_DIR, project="ASK")
        assert client_log["available"] is True
        assert client_log["source_file"] == os.path.basename(temp_vdlr)

        # 6. Trial: Fingerprint change & Server cache invalidation
        fp_modified = _compute_workspace_fingerprint()
        assert fp_modified != fp_initial, "Workspace fingerprint must change when files are added or updated!"

        cached_snap = get_cached_snapshot(project="ASK", mode="sample", force_refresh=True)
        assert cached_snap["sources_summary"]["psr_file"] == os.path.basename(temp_psr)

    finally:
        import gc
        gc.collect()
        for f in created_temp_files:
            for _ in range(15):
                if os.path.exists(f):
                    try:
                        os.remove(f)
                        break
                    except OSError:
                        time.sleep(0.2)
        import vendor_dashboard.ingestors
        vendor_dashboard.ingestors._INGEST_CACHE.clear()
        import server
        server._SNAPSHOT_CACHE.clear()

        # Confirm workspace restored cleanly to original baseline
        snap_restored = build_project_snapshot(ROOT_DIR, project="ASK", mode="sample")
        assert snap_restored["sources_summary"]["psr_file"] == "Attachment 2-Procurement Plan-ASK - 20260626.xlsx"


def test_wpr_mom_extraction_and_excel_sync() -> None:
    """Verify that WPR Gantt tasks, S-curves, sub-orders, and MOM actions are fully extracted
    and synchronized into 'Export Data File.xlsx', and that ingestors load them seamlessly without raw PDFs.
    """
    from vendor_dashboard.extractors.wpr_mom_extractor import extract_wpr_pdf, extract_mom_pdf
    from vendor_dashboard.extractors.excel_sync import sync_wpr_data_to_excel, sync_mom_data_to_excel
    from vendor_dashboard.ingestors import read_wpr_export_file, parse_all_wpr_files

    wpr_path = os.path.join(ROOT_DIR, "WPR", "MM-ASK-1A-GEN01-CIP-MEC-WPR-0010.pdf")
    mom_path = os.path.join(ROOT_DIR, "MOM", "MM-ASK-1A-GEN01-CIP-MEC-MOM-0010.pdf")
    export_path = os.path.join(ROOT_DIR, "Export Data File.xlsx")

    assert os.path.isfile(wpr_path), f"Missing WPR-0010 at {wpr_path}"
    assert os.path.isfile(mom_path), f"Missing MOM-0010 at {mom_path}"

    # 1. WPR Extraction
    wpr_rec = extract_wpr_pdf(wpr_path)
    assert wpr_rec["summary"]["package"] == "CIP"
    assert wpr_rec["summary"]["report_no"] == "WPR-0010"
    assert len(wpr_rec["gantt_tasks"]) == 70, f"Expected 70 tasks, got {len(wpr_rec['gantt_tasks'])}"
    assert len(wpr_rec["scurve"]) == 29
    assert len(wpr_rec["sub_orders"]) >= 9

    # 2. MOM Extraction
    mom_rec = extract_mom_pdf(mom_path)
    assert mom_rec["metadata"]["package"] == "CIP"
    assert mom_rec["metadata"]["mom_no"] == "MOM-0010"
    assert len(mom_rec["action_items"]) >= 10
    assert len(mom_rec["procurement_status"]) == 14

    # 3. Excel Sync
    wpr_stats = sync_wpr_data_to_excel(export_path, [wpr_rec])
    assert wpr_stats["summaries"] == 1
    mom_stats = sync_mom_data_to_excel(export_path, [mom_rec])
    assert mom_stats["procurement_status"] == 14

    # 4. Pure Cloud Fallback Verification (No raw WPR PDFs needed)
    exported_pkgs = read_wpr_export_file(export_path)
    assert "CIP" in exported_pkgs
    assert len(exported_pkgs["CIP"]["gantt_tasks"]) == 70

    cloud_wprs = parse_all_wpr_files(os.path.join(ROOT_DIR, "NON_EXISTENT_WPR_CLOUD_SIM"))
    assert "CIP" in cloud_wprs
    assert len(cloud_wprs["CIP"]["gantt_tasks"]) == 70


def test_master_schedule_ros_and_7_14_days_lookahead_hub() -> None:
    """Verify Master Schedule MM-ASK-1A-GEN01-PJM-SCD-0001 Rev. B1 is resumed as Priority #1 ROS
    reference together with Procurement Plan, and that 7 & 14 days look-ahead activities
    are generated across both Master Schedule and Package levels, synced to Excel, and wired to UI.
    """
    import openpyxl

    snap = build_project_snapshot(ROOT_DIR, project="ASK", mode="live")
    la_hub = snap["lookahead_hub"]
    assert la_hub is not None
    assert "summary_7d" in la_hub
    assert "summary_14d" in la_hub
    assert len(la_hub["activities_7d"]) > 0
    assert len(la_hub["activities_14d"]) >= len(la_hub["activities_7d"])

    # Check scopes
    scopes_7d = {x["scope"] for x in la_hub["activities_7d"]}
    assert "Master Schedule" in scopes_7d, "Missing Master Schedule scope in 7-day lookahead"
    assert "Package" in scopes_7d, "Missing Package scope in 7-day lookahead"

    # Check Critical Path activities flagged
    crit_acts = [x for x in la_hub["activities_7d"] if x["critical_path"] == "YES"]
    assert len(crit_acts) > 0, "Expected critical path activities flagged"

    # Verify Excel sync
    export_path = os.path.join(ROOT_DIR, "Export Data File.xlsx")
    assert os.path.isfile(export_path)
    wb = openpyxl.load_workbook(export_path)
    assert "LookAhead_7Days" in wb.sheetnames
    assert "LookAhead_14Days" in wb.sheetnames
    ws7 = wb["LookAhead_7Days"]
    assert ws7.max_row > 10, f"Expected >10 rows in LookAhead_7Days, got {ws7.max_row}"
    ws14 = wb["LookAhead_14Days"]
    assert ws14.max_row >= ws7.max_row

    # Verify Frontend wiring
    with open(os.path.join(ROOT_DIR, "index.html"), "r", encoding="utf-8") as fh:
        html = fh.read()
    assert "tabBtnLookAhead" in html
    assert "panelLookAhead" in html
    assert "btnWindow7d" in html
    assert "btnWindow14d" in html
    assert "lookAheadTableBody" in html

    with open(os.path.join(ROOT_DIR, "app.js"), "r", encoding="utf-8") as fh:
        js = fh.read()
    assert "renderLookAheadHub" in js
    assert "setLookAheadWindow" in js
    assert "panelLookAhead" in js


def test_z1f_scd_and_mom_ingestion() -> None:
    """Verify trial test and implementation for Z1F:
    1. Parse Master Schedule MM-ZTK-1F-GEN-PJM-SCD-0002 Rev. C1 (4,110 activities).
    2. Extract MOM MM-ZTK-1F-GEN-PDC-MEC-MOM-0010 Rev.00 (PDC package, 11 action items).
    3. Verify Z1F snapshot sources and 7-day / 14-day Lookahead Hub.
    4. Confirm Export Data File_Z1F.xlsx is generated, isolated from ASK, and populated.
    """
    import openpyxl
    from vendor_dashboard.ingestors import parse_master_schedule_pdf
    from vendor_dashboard.extractors.wpr_mom_extractor import extract_mom_pdf

    # 1. SCD Ingestion
    scd = parse_master_schedule_pdf(ROOT_DIR, project="Z1F")
    assert scd["available"] is True
    assert scd["document_no"] == "MM-ZTK-1F-GEN-PJM-SCD-0002"
    assert scd["revision"] == "C1"
    assert len(scd["all_activities"]) >= 4000, f"Expected >=4000 activities, got {len(scd['all_activities'])}"
    assert "PDC" in scd["catalog"]
    assert scd["catalog"]["PDC"]["activity_id"] == "1F-WP1-WU1.4.1-1250"
    assert scd["catalog"]["PDC"]["master_eta_date"] == "2027-06-22"

    # 2. MOM Ingestion
    mom_path = os.path.join(ROOT_DIR, "MOM", "MM-ZTK-1F-GEN-PDC-MEC-MOM-0010 Rev.00.pdf")
    assert os.path.isfile(mom_path), f"Missing Z1F MOM file at {mom_path}"
    mom_rec = extract_mom_pdf(mom_path)
    assert mom_rec["metadata"]["package"] == "PDC"
    assert mom_rec["metadata"]["mom_no"] == "MOM-0010"
    assert len(mom_rec["action_items"]) == 11

    # 3. Snapshot & LookAhead Hub
    snap = build_project_snapshot(ROOT_DIR, project="Z1F", mode="live")
    assert snap["project"] == "Z1F"
    assert snap["sources_summary"]["master_schedule_file"] == "MM-ZTK-1F-GEN-PJM-SCD-0002.pdf"
    la_hub = snap.get("lookahead_hub", {})
    assert len(la_hub.get("activities_7d", [])) > 300
    assert len(la_hub.get("activities_14d", [])) > 500

    # 4. Export File Isolation and Contents
    z1f_export_path = os.path.join(ROOT_DIR, "Export Data File_Z1F.xlsx")
    assert os.path.isfile(z1f_export_path), f"Missing {z1f_export_path}"
    wb = openpyxl.load_workbook(z1f_export_path, read_only=True)
    assert "MOM_Action_Items" in wb.sheetnames
    assert "LookAhead_7Days" in wb.sheetnames
    assert "LookAhead_14Days" in wb.sheetnames
    ws_mom = wb["MOM_Action_Items"]
    mom_rows = list(ws_mom.iter_rows(values_only=True))
    assert len(mom_rows) >= 12  # Header + 11 items
    assert mom_rows[1][0] == "PDC"
    assert mom_rows[1][1] == "MOM-0010"


def test_live_execution_mode_and_cip_wpr_mom_reconciliation() -> None:
    """Verify:
    1. Default mode is 'live' (Execution Mode) — packages do NOT borrow PLC sample data.
    2. CIP WPR-0010 (Cut-off 2026-10-27) is ingested with Planned 23.27%, Actual 17.99%, 11 milestones, 15 milestone progress items, 70 Gantt tasks, and 10 sub-orders.
    3. Governing ETA follows Gantt Chart (2027-02-23 from Task 70 '4.2 Sea Freight / Transit to China Port (CIF)'), highlights +15d discrepancy vs Milestone #10 (2027-02-08), and compares against Master Schedule ROS (2027-03-05, +10d buffer).
    4. CIP MOM-0010 (2026-10-28) is ingested with 48 action items (21 OPEN) and 14 procurement items, highlighting the 28-week manufacturing lead time (+6 weeks impact) and expediting of Flow Switch & Metering Pump.
    5. Each Element has cut_off_date & ref_doc metadata, and sources_summary['incorporated_files'] lists incorporated WPRs, MOMs, SCD, and PSR.
    6. Packages without WPR/MOM (e.g. ABV) in live mode show only SCD, PSR, and VDRL baseline data.
    """
    snap = build_project_snapshot(ROOT_DIR, project="ASK")
    assert snap["mode"] == "live", f"Expected default mode 'live', got {snap['mode']}"

    pkgs = {p["package_code"]: p for p in snap["packages"]}
    assert "CIP" in pkgs
    cip = pkgs["CIP"]

    # 1. CIP Live WPR-0010 verification
    assert cip["has_live_wpr"] is True
    assert cip["using_sample_wpr"] is False
    assert cip["report_no"] == "WPR-0010"
    assert cip["cut_off_date"] == "2026-10-27"
    assert cip["overall_progress"]["planned_pct"] == 23.27
    assert cip["overall_progress"]["actual_pct"] == 17.99
    assert len(cip["milestone_progress"]) == 15
    assert len(cip["gantt_tasks"]) >= 68
    assert cip["attachment4_summary"]["total_components"] == 10

    # 2. Gantt Chart ETA Priority & Milestone Discrepancy Alert
    assert cip["eta_date"] == "2027-02-23"
    assert cip["gantt_eta_date"] == "2027-02-23"
    assert cip["milestone_eta_date"] == "2027-02-08"
    assert cip["has_milestone_discrepancy"] is True
    assert cip["milestone_discrepancy_days"] == 15
    assert "2027-02-08" in cip["milestone_discrepancy_note"]
    assert "2027-02-23" in cip["milestone_discrepancy_note"]

    # 3. Comparison with Master Schedule (MM-ASK-1A-GEN01-PJM-SCD-0001) ROS Date
    hl = cip["package_highlights"]
    assert hl["effective_delivery_date"] == "2027-02-23"
    assert hl["master_eta_date"] == "2027-03-05"
    assert hl["governing_ros_date"] == "2027-03-05"
    assert hl["is_ros_delayed"] is False
    assert hl["ros_margin_days"] == 10

    # 4. CIP MOM-0010 Overall Situation & Concerns
    assert cip["has_live_mom"] is True
    mom = cip["mom_summary"]
    assert mom["available"] is True
    assert mom["mom_no"] == "MOM-0010"
    assert mom["cut_off_date"] == "2026-10-28"
    assert mom["total_items"] == 48
    assert mom["open_count"] == 21
    assert len(mom["procurement_status"]) == 14
    assert "28 weeks" in mom["overall_situation_summary"]
    assert any("28 Weeks" in c["title"] or "28 weeks" in c["detail"] for c in mom["key_concerns"])
    assert any("28 Weeks" in c["title"] or "Discrepancy" in c["title"] for c in cip["areas_of_concern"])

    # 5. Element Reference & Cut-Off Metadata + Incorporated Files Notification
    em = cip["element_metadata"]
    for key in ("element_1_2", "element_3", "element_4", "element_5", "element_6", "element_7", "element_8", "element_9", "mom"):
        assert key in em, f"Missing {key} in element_metadata"
        assert em[key].get("ref_doc"), f"Empty ref_doc for {key}"
        assert em[key].get("cut_off_date"), f"Empty cut_off_date for {key}"

    inc_files = snap["sources_summary"].get("incorporated_files", [])
    assert any(f["category"] == "WPR" and f["package"] == "CIP" and f["cut_off_date"] == "2026-10-27" for f in inc_files)
    assert any(f["category"] == "MOM" and f["package"] == "CIP" and f["cut_off_date"] == "2026-10-28" for f in inc_files)

    # 6. Packages without WPR (e.g. ABV) in live mode do NOT borrow PLC sample WPR
    abv = pkgs["ABV"]
    assert abv["has_live_wpr"] is False
    assert abv["using_sample_wpr"] is False
    assert "No WPR Uploaded" in abv["element_metadata"]["element_5"]["ref_doc"]
    assert "PLC" not in abv["element_metadata"]["element_4"]["ref_doc"]
    assert abv["vdrl_status"]["available"] is True


if __name__ == "__main__":
    for fn in [
        test_vdrl_multi_and_single_package_parsing,
        test_client_doc_log_parsing,
        test_kom_manufacturer_rule,
        test_kom_export_data_file_fallback,
        test_wpr_pdf_and_docx_parsing,
        test_reconciler_priority_and_modes,
        test_server_snapshot_cache_and_elements,
        test_feedback_items_1_to_7,
        test_latest_5_feedback_items,
        test_procurement_plan_ros_priority_and_one_picture_executive_summary,
        test_dynamic_weekly_file_updates_seamless,
        test_wpr_mom_extraction_and_excel_sync,
        test_master_schedule_ros_and_7_14_days_lookahead_hub,
        test_z1f_scd_and_mom_ingestion,
        test_live_execution_mode_and_cip_wpr_mom_reconciliation,
    ]:
        fn()
        print(f"[PASS] {fn.__name__}")
    print("ALL TESTS PASSED SUCCESSFULLY.")


