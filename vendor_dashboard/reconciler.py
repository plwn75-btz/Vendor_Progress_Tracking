"""Reconciler for the ASK & Z1F Weekly Project Management & Vendor Package Monitoring Dashboard.
Implements:
- Project separation & switching (ASK vs. Z1F)
- Priority Resolution: WPR (Priority #1) > PSR / Procurement Plan (Priority #2)
- Dynamic Package-Specific S-Curve, Milestone Progress, & 7-14d Lookahead when switching packages
- Project-Wide Weighted Overall S-Curve calculated across all packages
- Package Key Highlights (4.1): Current Status by Key Milestone, Remaining Days to Next Milestone, Days to Delivery, & Fixed ROS Date Delay Detection
- WPR Attachment 3 (Fabrication Detail Gantt Analysis) & Attachment 4 (Supplier Material Expediting Status) Summary
- Strict KOM Manufacturer Detail Rule: Only populated from KOM/; blank if not available in KOM/
- Dual Status Vendor Document Monitoring: VDRL (Vendor Submission) + Vendor Document Log Report (Client Review & Return)
"""

import copy
import datetime
import os
import re
from typing import Any, Dict, List, Optional, Tuple

from vendor_dashboard.ingestors import (
    fmt_date,
    match_master_schedule_entry,
    parse_all_mom_files,
    parse_all_vdrl_files,
    parse_all_wpr_files,
    parse_client_doc_log,
    parse_kom_folder,
    parse_master_schedule_pdf,
    parse_psr_file,
)
from vendor_dashboard.models import (
    ASK_PACKAGE_METADATA,
    Z1F_PACKAGE_METADATA,
    ManufacturerDetail,
    STANDARD_MILESTONES,
)

# Relative package weightings for Project Overall Weighted S-Curve calculation
DISCIPLINE_BASE_WEIGHTS: Dict[str, float] = {
    "MEC": 6.0,
    "PV": 5.0,
    "INS": 3.5,
    "ELE": 4.5,
    "PIP": 2.5,
}


def _date_diff_days(start_str: str, end_str: str) -> Optional[int]:
    """Return (end - start) in days."""
    if not start_str or not end_str:
        return None
    try:
        d1 = datetime.date.fromisoformat(start_str[:10])
        d2 = datetime.date.fromisoformat(end_str[:10])
        return (d2 - d1).days
    except ValueError:
        return None


def _within_window(date_str: str, start_iso: str, days: int = 14) -> bool:
    if not date_str or not start_iso:
        return False
    try:
        dt = datetime.date.fromisoformat(date_str[:10])
        s_dt = datetime.date.fromisoformat(start_iso[:10])
        e_dt = s_dt + datetime.timedelta(days=days)
        return s_dt <= dt <= e_dt
    except ValueError:
        return False


def _task_active_in_window(start_str: str, finish_str: str, ref_iso: str, days: int = 14) -> bool:
    if not start_str or not ref_iso:
        return False
    try:
        s_dt = datetime.date.fromisoformat(start_str[:10])
        f_dt = datetime.date.fromisoformat((finish_str or start_str)[:10])
        w_start = datetime.date.fromisoformat(ref_iso[:10])
        w_end = w_start + datetime.timedelta(days=days)
        return s_dt <= w_end and f_dt >= w_start
    except ValueError:
        return False


def _add_days(iso_str: str, days: int) -> str:
    if not iso_str:
        return ""
    try:
        dt = datetime.date.fromisoformat(iso_str[:10])
        return (dt + datetime.timedelta(days=days)).strftime("%Y-%m-%d")
    except ValueError:
        return ""


def _clean_and_dedup_master_activities(raw_acts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filter out WBS summary headers and deduplicate Master Schedule activities across PDF pages."""
    by_id: Dict[str, Dict[str, Any]] = {}
    ordered_ids: List[str] = []

    for act in raw_acts or []:
        aid = re.sub(r"\s+", " ", str(act.get("activity_id") or "")).strip()
        aname = re.sub(r"\s+", " ", str(act.get("activity_name") or "")).strip()
        # Real Primavera/MSP Activity IDs have no spaces and contain at least one digit;
        # this filters out WBS summary banners like "Structure", "Pre-assembly", "Vendor Document Review and A".
        if not aid or " " in aid or not re.search(r"\d", aid):
            continue

        s_dt = str(act.get("start_date") or "")[:10]
        f_dt = str(act.get("finish_date") or "")[:10]
        key = aid.upper()

        if key in by_id:
            cur = by_id[key]
            cur["is_critical_path"] = bool(cur.get("is_critical_path")) or bool(act.get("is_critical_path"))
            if len(aname) > len(cur.get("activity_name") or ""):
                cur["activity_name"] = aname
            if not cur.get("start_date") and s_dt:
                cur["start_date"] = s_dt
            if not cur.get("finish_date") and f_dt:
                cur["finish_date"] = f_dt
            if cur.get("duration_days") is None and act.get("duration_days") is not None:
                cur["duration_days"] = act.get("duration_days")
            if cur.get("total_float_days") is None and act.get("total_float_days") is not None:
                cur["total_float_days"] = act.get("total_float_days")
        else:
            copied = dict(act)
            copied["activity_id"] = aid
            copied["activity_name"] = aname
            by_id[key] = copied
            ordered_ids.append(key)

    # Secondary signature deduplication (identical activity name + exact start + exact finish)
    deduped: List[Dict[str, Any]] = []
    seen_sigs: set = set()
    for key in ordered_ids:
        item = by_id[key]
        sig = (
            str(item.get("activity_name") or "").upper(),
            str(item.get("start_date") or "")[:10],
            str(item.get("finish_date") or "")[:10],
        )
        if sig in seen_sigs:
            continue
        seen_sigs.add(sig)
        deduped.append(item)
    return deduped


def build_project_lookahead_hub(
    packages_list: List[Dict[str, Any]],
    master_schedule: Dict[str, Any],
    cut_off_date: str = "2026-10-16",
) -> Dict[str, Any]:
    """Build unified 7-Day and 14-Day Look-Ahead Hub combining:
    1. MM-ASK-1A-GEN01-PJM-SCD-0001 (Project Master Schedule Level: EPC Engineering, Procurement, Yard Erection, Critical Path TF=0)
    2. Each package level (WPR Gantt tasks, VDRL submissions, Sub-order material status, Milestone gates)
    """
    try:
        ref_dt = datetime.date.fromisoformat(cut_off_date[:10]) if cut_off_date else datetime.date(2026, 10, 16)
    except (ValueError, TypeError):
        ref_dt = datetime.date(2026, 10, 16)

    end_7 = ref_dt + datetime.timedelta(days=7)
    end_14 = ref_dt + datetime.timedelta(days=14)
    end_7_iso = end_7.strftime("%Y-%m-%d")
    end_14_iso = end_14.strftime("%Y-%m-%d")

    def _in_window_dates(s_iso: str, f_iso: str, window_end: datetime.date) -> bool:
        if not s_iso and not f_iso:
            return False
        try:
            s = datetime.date.fromisoformat(s_iso[:10]) if s_iso else None
            f = datetime.date.fromisoformat(f_iso[:10]) if f_iso else s
            if not f:
                f = s
            if not s:
                s = f
            if ref_dt <= s <= window_end or ref_dt <= f <= window_end:
                return True
            if s <= ref_dt <= f and (f - s).days <= 45:
                return True
            return False
        except (ValueError, TypeError):
            return False

    def _collect_for_window(window_days: int, window_end: datetime.date, window_label: str) -> List[Dict[str, Any]]:
        items: List[Dict[str, Any]] = []

        # 1. Master Schedule Activities (ASK: MM-ASK-1A-GEN01-PJM-SCD-0001, Z1F: MM-ZTK-1F-GEN-PJM-SCD-0002)
        sched_doc_no = (master_schedule or {}).get("document_no") or ("MM-ASK-1A-GEN01-PJM-SCD-0001" if (master_schedule or {}).get("source_file", "").startswith("MM-ASK") else "MM-ZTK-1F-GEN-PJM-SCD-0002")
        if (master_schedule or {}).get("from_export_cache"):
            cached_la = (master_schedule or {}).get("export_lookahead_7d" if window_days <= 7 else "export_lookahead_14d", [])
            if cached_la:
                items.extend(dict(x) for x in cached_la)
                items.sort(key=lambda x: (
                    0 if x.get("critical_path") == "YES" else 1,
                    x.get("start_date") or "9999-99-99",
                    x.get("package_or_id") or "",
                ))
                return items
            else:
                for ma in (master_schedule or {}).get("all_activities", []):
                    s_dt = ma.get("start_date", "")
                    f_dt = ma.get("finish_date", "")
                    if _in_window_dates(s_dt, f_dt, window_end):
                        items.append({
                            "scope": "Master Schedule",
                            "package_or_id": ma.get("activity_id", ""),
                            "package_name": f"Project Master Plan ({sched_doc_no})",
                            "activity_id": ma.get("activity_id", ""),
                            "activity_name": ma.get("activity_name", ""),
                            "start_date": s_dt,
                            "finish_date": f_dt,
                            "duration_days": ma.get("duration_days"),
                            "critical_path": "YES" if ma.get("is_critical_path") else "NO",
                            "category": ma.get("category", "General EPC"),
                            "source": sched_doc_no,
                            "status": "Critical Path (TF=0)" if ma.get("is_critical_path") else "On Schedule",
                            "window": window_label,
                        })
        else:
            all_master_acts = (master_schedule or {}).get("all_activities", [])
            for ma in all_master_acts:
                s_dt = ma.get("start_date", "")
                f_dt = ma.get("finish_date", "")
                if _in_window_dates(s_dt, f_dt, window_end):
                    items.append({
                        "scope": "Master Schedule",
                        "package_or_id": ma.get("activity_id", ""),
                        "package_name": f"Project Master Plan ({sched_doc_no})",
                        "activity_id": ma.get("activity_id", ""),
                        "activity_name": ma.get("activity_name", ""),
                        "start_date": s_dt,
                        "finish_date": f_dt,
                        "duration_days": ma.get("duration_days"),
                        "critical_path": "YES" if ma.get("is_critical_path") else "NO",
                        "category": ma.get("category", "General EPC"),
                        "source": sched_doc_no,
                        "status": "Critical Path (TF=0)" if ma.get("is_critical_path") else "On Schedule",
                        "window": window_label,
                    })

        # 2. Package Level Activities
        for pkg in packages_list:
            p_code = pkg.get("package_code", "")
            p_name = pkg.get("package_name", "")
            p_crit = "YES" if pkg.get("is_critical_path") else "NO"

            # 2a. WPR Attachment 3 Fabrication Gantt Tasks
            tasks = pkg.get("gantt_tasks") or pkg.get("attachment3_summary", {}).get("tasks", [])
            for t in tasks:
                if t.get("is_summary"):
                    continue
                ts = t.get("start_date", "")
                tf = t.get("finish_date", "")
                if _in_window_dates(ts, tf, window_end):
                    items.append({
                        "scope": "Package",
                        "package_or_id": p_code,
                        "package_name": p_name,
                        "activity_id": str(t.get("id") or t.get("task_id", "")),
                        "activity_name": t.get("name", ""),
                        "start_date": ts,
                        "finish_date": tf,
                        "duration_days": t.get("duration"),
                        "critical_path": p_crit,
                        "category": f"Fabrication Gantt ({t.get('phase', 'Assembly')})",
                        "source": f"WPR {p_code} Att. 3",
                        "status": "In Progress" if t.get("is_in_progress") else "Scheduled",
                        "window": window_label,
                    })

            # 2b. VDRL Submittals
            docs = pkg.get("vdrl_status", {}).get("documents", [])
            for d in docs:
                if d.get("status") in ("SUBMITTED", "CLOSED"):
                    continue
                d_plan = d.get("plan_date", "")
                d_fc = d.get("forecast_date") or d_plan
                is_overdue = "OVERDUE" in str(d.get("status", "")).upper()
                if is_overdue or _in_window_dates(d_plan, d_fc, window_end):
                    is_key = d.get("is_key") == "Y"
                    items.append({
                        "scope": "Package",
                        "package_or_id": p_code,
                        "package_name": p_name,
                        "activity_id": d.get("doc_no", ""),
                        "activity_name": f"VDRL: {d.get('title', '')} ({'Key Document' if is_key else 'Standard'})",
                        "start_date": d_plan,
                        "finish_date": d_fc,
                        "duration_days": 1,
                        "critical_path": "YES" if is_key else p_crit,
                        "category": "VDRL Submittal",
                        "source": f"VDRL ({p_code})",
                        "status": "Overdue" if is_overdue else "Due Soon",
                        "window": window_label,
                    })

            # 2c. Attachment 4 Sub-Orders / Materials Expediting
            sub_items = pkg.get("attachment4_summary", {}).get("items", [])
            for it in sub_items:
                m_plan = it.get("planned_order_date") or it.get("planned_eta") or ""
                m_eta = it.get("planned_eta") or m_plan
                if _in_window_dates(m_plan, m_eta, window_end):
                    is_hold = "Hold" in str(it.get("po_status", ""))
                    items.append({
                        "scope": "Package",
                        "package_or_id": p_code,
                        "package_name": p_name,
                        "activity_id": f"MAT-{p_code}-{it.get('item_no', '6')}",
                        "activity_name": f"Material Expediting: {it.get('component', '')} ({it.get('prerequisite_gate', '')})",
                        "start_date": m_plan,
                        "finish_date": m_eta,
                        "duration_days": it.get("lead_time_days"),
                        "critical_path": "YES" if is_hold else p_crit,
                        "category": "Material & Procurement",
                        "source": f"WPR {p_code} Att. 4",
                        "status": "On Hold (Prereq)" if is_hold else "Expediting",
                        "window": window_label,
                    })

            # 2d. Key Milestones Due
            for m in pkg.get("key_milestones", []):
                if not m.get("actual_date") and (m.get("planned_date") or m.get("forecast_date")):
                    m_dt = m.get("forecast_date") or m.get("planned_date") or ""
                    if _in_window_dates(m_dt, m_dt, window_end):
                        items.append({
                            "scope": "Package",
                            "package_or_id": p_code,
                            "package_name": p_name,
                            "activity_id": f"MS-{p_code}-{m.get('no', '')}",
                            "activity_name": f"Key Milestone #{m.get('no')}: {m.get('description', '')}",
                            "start_date": m.get("planned_date") or m_dt,
                            "finish_date": m_dt,
                            "duration_days": 1,
                            "critical_path": p_crit,
                            "category": "Milestone Gate",
                            "source": m.get("source", "PSR Baseline"),
                            "status": m.get("status", "Due Soon"),
                            "window": window_label,
                        })

        items.sort(key=lambda x: (
            0 if x.get("critical_path") == "YES" else 1,
            x.get("start_date") or "9999-99-99",
            x.get("package_or_id") or "",
        ))
        return items

    acts_7d = _collect_for_window(7, end_7, "7-Day")
    acts_14d = _collect_for_window(14, end_14, "14-Day")

    return {
        "cut_off_date": cut_off_date,
        "window_7d_end": end_7_iso,
        "window_14d_end": end_14_iso,
        "summary_7d": {
            "total": len(acts_7d),
            "master_schedule_count": sum(1 for x in acts_7d if x["scope"] == "Master Schedule"),
            "package_count": sum(1 for x in acts_7d if x["scope"] == "Package"),
            "critical_path_count": sum(1 for x in acts_7d if x["critical_path"] == "YES"),
        },
        "summary_14d": {
            "total": len(acts_14d),
            "master_schedule_count": sum(1 for x in acts_14d if x["scope"] == "Master Schedule"),
            "package_count": sum(1 for x in acts_14d if x["scope"] == "Package"),
            "critical_path_count": sum(1 for x in acts_14d if x["critical_path"] == "YES"),
        },
        "activities_7d": acts_7d,
        "activities_14d": acts_14d,
    }


def build_project_snapshot(
    workspace_dir: str,
    project: str = "ASK",
    mode: str = "live",
) -> Dict[str, Any]:
    """Build complete reconciled dashboard payload for project ('ASK' or 'Z1F')."""
    proj = project.upper() if project.upper() in ("ASK", "Z1F") else "ASK"
    is_ask = proj == "ASK"
    catalog = ASK_PACKAGE_METADATA if is_ask else Z1F_PACKAGE_METADATA

    vdrl_dir = os.path.join(workspace_dir, "VDRL")
    wpr_dir = os.path.join(workspace_dir, "WPR")
    mom_dir = os.path.join(workspace_dir, "MOM")
    kom_dir = os.path.join(workspace_dir, "KOM")

    # 1. Ingest all 7 dynamic sources (including Project Master Schedule & MOMs)
    kom_map = parse_kom_folder(kom_dir, project=proj)
    vdrl_map = parse_all_vdrl_files(vdrl_dir, project_filter=proj)
    client_log = parse_client_doc_log(workspace_dir, project=proj)
    psr_map = parse_psr_file(workspace_dir, project=proj)
    wpr_map = parse_all_wpr_files(wpr_dir, project_filter=proj)
    mom_map = parse_all_mom_files(mom_dir, project_filter=proj)
    master_schedule = parse_master_schedule_pdf(workspace_dir, project=proj)

    plc_sample_wpr: Optional[Dict[str, Any]] = wpr_map.get("PLC")
    if not plc_sample_wpr:
        all_ask_wprs = parse_all_wpr_files(wpr_dir, project_filter="ASK")
        plc_sample_wpr = all_ask_wprs.get("PLC")

    pkg_codes: List[str] = []
    for c in wpr_map:
        if c not in pkg_codes:
            pkg_codes.append(c)
    for c in mom_map:
        if c not in pkg_codes and (c in catalog or c in vdrl_map or c in psr_map):
            pkg_codes.append(c)
    for c in vdrl_map:
        if c not in pkg_codes:
            pkg_codes.append(c)
    for c in catalog:
        if c not in pkg_codes:
            if is_ask and (c in vdrl_map or c in kom_map or c in wpr_map or c in client_log.get("by_package", {}) or c in psr_map):
                pkg_codes.append(c)
            elif not is_ask:
                pkg_codes.append(c)
    # Point 4: Include ALL equipment from PSR (e.g., Barred Tee, Corrosion Coupon, Sight Glass, etc.)
    # even when there is no VDRL file yet, so they are selectable in Tab 1 & ready for future VDRLs.
    for c in psr_map:
        if c not in pkg_codes:
            pkg_codes.append(c)

    # Calculate normalized package weights (sum = 100.0%) for Project Overall Weighted S-Curve
    raw_weights: Dict[str, float] = {}
    for code in pkg_codes:
        disc = catalog.get(code, {}).get("discipline") or psr_map.get(code, {}).get("discipline") or "INS"
        base_w = DISCIPLINE_BASE_WEIGHTS.get(disc, 2.0)
        if code in ("PDC", "PV", "TSP", "PSP", "PLC", "WHCP", "CIP", "SOLR"):
            base_w *= 1.45
        elif code not in catalog and code not in vdrl_map:
            base_w *= 0.35
        raw_weights[code] = base_w
    total_raw_w = sum(raw_weights.values()) or 1.0
    norm_weights: Dict[str, float] = {
        code: round((w / total_raw_w) * 100.0, 2) for code, w in raw_weights.items()
    }

    packages_list: List[Dict[str, Any]] = []
    for code in pkg_codes:
        psr_entry = psr_map.get(code)
        default_meta = {
            "name": (psr_entry or {}).get("package_name") or code,
            "discipline": (psr_entry or {}).get("discipline") or "GEN",
            "rfq": (psr_entry or {}).get("rfq_no") or "",
            "psr_name": (psr_entry or {}).get("package_name") or code,
            "tag": "WP-01 / WP-02" if is_ask else "ZWP20 / ZWP21 / ZWP22 / ZWP23",
        }
        pkg_payload = _reconcile_single_package(
            proj=proj,
            pkg_code=code,
            pkg_weight_pct=norm_weights.get(code, 1.0),
            meta=catalog.get(code, default_meta),
            kom_detail=kom_map.get(code),
            vdrl_data=vdrl_map.get(code),
            client_pkg_data=client_log.get("by_package", {}).get(code),
            psr_data=psr_entry,
            live_wpr=wpr_map.get(code),
            sample_wpr=plc_sample_wpr,
            mode=mode,
            master_schedule=master_schedule,
            mom_detail=mom_map.get(code),
        )
        pkg_payload["is_core_package"] = bool(
            code in vdrl_map or code in kom_map or code in wpr_map or code in catalog
        )
        pkg_payload["is_psr_equipment_only"] = not pkg_payload["is_core_package"]
        packages_list.append(pkg_payload)

    # Build Project-Wide Weighted Overall S-Curve across all packages (Comment #2)
    project_overall_scurve = _build_project_weighted_scurve(packages_list)

    # Compute Executive Portfolio KPIs
    total_packages = len(packages_list)
    core_packages_count = sum(1 for p in packages_list if p.get("is_core_package"))
    psr_only_equipment_count = sum(1 for p in packages_list if p.get("is_psr_equipment_only"))
    live_wpr_count = sum(1 for p in packages_list if p["has_live_wpr"])
    live_mom_count = sum(1 for p in packages_list if p.get("has_live_mom"))
    kom_available_count = sum(1 for p in packages_list if p["manufacturer_detail"]["available"])

    total_vdrl_docs = sum(p["vdrl_status"]["total_docs"] for p in packages_list)
    total_vdrl_submitted = sum(p["vdrl_status"]["submitted"] for p in packages_list)
    total_vdrl_overdue = sum(p["vdrl_status"]["overdue"] for p in packages_list)
    total_key_vdrl_overdue = sum(p["vdrl_status"]["key_docs_overdue"] for p in packages_list)
    total_prod_vdrl_overdue = sum(p["vdrl_status"]["production_impact_overdue"] for p in packages_list)
    portfolio_vdrl_pct = round((total_vdrl_submitted / total_vdrl_docs) * 100.0, 1) if total_vdrl_docs > 0 else 0.0

    overall_planned_pct = project_overall_scurve["planned_to_date_pct"]
    overall_actual_pct = project_overall_scurve["actual_to_date_pct"]
    ros_delayed_list = [p["package_name"] for p in packages_list if p["package_highlights"]["is_ros_delayed"]]
    ros_delayed_packages = len(ros_delayed_list)
    tight_ros_list = [
        p["package_name"]
        for p in packages_list
        if not p["package_highlights"]["is_ros_delayed"]
        and p["package_highlights"].get("ros_margin_days") is not None
        and 0 <= p["package_highlights"]["ros_margin_days"] <= 14
    ]
    total_concerns = sum(len(p["areas_of_concern"]) for p in packages_list)
    active_psr_file = next(iter(psr_map.values()), {}).get(
        "source_file",
        "Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx"
        if is_ask
        else "Attachment 1 - Procurement Plan-Z1F_20261002.xlsx",
    )
    active_master_file = (
        master_schedule.get("source_file")
        or ("MM-ASK-1A-GEN01-PJM-SCD-0001.pdf" if is_ask else "MM-ZTK-1F-GEN-PJM-SCD-0002.pdf")
    )

    # Build One-Picture 4-Tier Health Classification & Smart Mitigation Matrix
    def _build_tier_group(pkg_subset: List[Dict[str, Any]]) -> Dict[str, Any]:
        total_n = len(pkg_subset) or 1
        tiers = {
            "ALREADY_DELAYED": [],
            "HIGH_RISK": [],
            "LOW_RISK": [],
            "HEALTHY": [],
        }
        for p in pkg_subset:
            hl = p["package_highlights"]
            t_key = p.get("health_tier", "HEALTHY")
            if t_key not in tiers:
                t_key = "HEALTHY"
            tiers[t_key].append({
                "package_code": p["package_code"],
                "package_name": p["package_name"],
                "discipline": p["discipline"],
                "is_core_package": p.get("is_core_package", True),
                "is_critical_path": hl.get("is_critical_path", False),
                "activity_id": hl.get("master_activity_id", ""),
                "eta_date": hl.get("effective_delivery_date", ""),
                "master_eta_date": hl.get("master_eta_date", ""),
                "yard_erection_date": hl.get("yard_erection_date", ""),
                "yard_erection_activity": hl.get("yard_erection_activity", ""),
                "ros_date": hl.get("ros_date", ""),
                "master_delay_days": hl.get("master_delay_days", 0),
                "master_margin_days": hl.get("master_margin_days"),
                "ros_delay_days": hl.get("ros_delay_days", 0),
                "effective_delay_days": hl.get("effective_delay_days", 0),
                "effective_margin_days": hl.get("ros_margin_days"),
                "vdrl_overdue": p["vdrl_status"].get("overdue", 0),
                "key_vdrl_overdue": p["vdrl_status"].get("key_docs_overdue", 0),
                "badge_text": hl.get("tier_badge_text", ""),
                "smart_issue": hl.get("smart_issue", ""),
                "smart_mitigation": hl.get("smart_mitigation", ""),
                "mitigation_deadline": hl.get("mitigation_deadline", ""),
                "mitigation_owner": hl.get("mitigation_owner", ""),
            })
        # Sort each tier by severity (largest delay / tightest buffer / critical path first)
        tiers["ALREADY_DELAYED"].sort(
            key=lambda x: (-int(x["effective_delay_days"] or 0), not x["is_critical_path"], x["package_name"])
        )
        tiers["HIGH_RISK"].sort(
            key=lambda x: (
                not x["is_critical_path"],
                int(x["effective_margin_days"] if x["effective_margin_days"] is not None else 999),
                -int(x["key_vdrl_overdue"] or 0),
                x["package_name"],
            )
        )
        tiers["LOW_RISK"].sort(key=lambda x: (-int(x["vdrl_overdue"] or 0), x["package_name"]))
        tiers["HEALTHY"].sort(key=lambda x: x["package_name"])

        return {
            "total": len(pkg_subset),
            "already_delayed": {
                "count": len(tiers["ALREADY_DELAYED"]),
                "pct": round(len(tiers["ALREADY_DELAYED"]) * 100.0 / total_n, 1),
                "packages": tiers["ALREADY_DELAYED"],
            },
            "high_risk": {
                "count": len(tiers["HIGH_RISK"]),
                "pct": round(len(tiers["HIGH_RISK"]) * 100.0 / total_n, 1),
                "packages": tiers["HIGH_RISK"],
            },
            "low_risk": {
                "count": len(tiers["LOW_RISK"]),
                "pct": round(len(tiers["LOW_RISK"]) * 100.0 / total_n, 1),
                "packages": tiers["LOW_RISK"],
            },
            "healthy": {
                "count": len(tiers["HEALTHY"]),
                "pct": round(len(tiers["HEALTHY"]) * 100.0 / total_n, 1),
                "packages": tiers["HEALTHY"],
            },
        }

    core_pkgs_only = [p for p in packages_list if p.get("is_core_package")]
    tiers_core = _build_tier_group(core_pkgs_only)
    tiers_all = _build_tier_group(packages_list)

    # Priority Mitigation Action Matrix (All Already Delayed + High Risk packages, ranked by priority)
    priority_mitigation_matrix = (
        tiers_all["already_delayed"]["packages"]
        + tiers_all["high_risk"]["packages"]
        + tiers_core["low_risk"]["packages"][:5]
    )

    executive_findings_summary = {
        "headline": (
            f"{proj} One-Picture Executive Summary — Delivery Tracked Against Master Schedule ROS ({active_master_file}) & Procurement Plan ROS Date ({active_psr_file})"
        ),
        "ros_data_reference": f"{active_master_file} (Priority #1) & {active_psr_file} (Priority #2)",
        "governing_ros_reference": active_master_file,
        "companion_ros_reference": active_psr_file,
        "active_master_schedule_file": active_master_file,
        "active_master_schedule_rev": master_schedule.get("revision", "B1"),
        "project_key_dates": master_schedule.get("project_key_dates", {}),
        "active_psr_file": active_psr_file,
        "tiers_core": tiers_core,
        "tiers_all": tiers_all,
        "priority_mitigation_matrix": priority_mitigation_matrix,
        "schedule_and_ros_finding": (
            f"Tracking Vendor Delivery against Master Schedule ROS ({active_master_file}) together with Procurement Plan ({active_psr_file}): "
            f"{tiers_all['already_delayed']['count']} package(s) are ALREADY DELAYED ({', '.join(ros_delayed_list[:6]) if ros_delayed_list else 'None'}), "
            f"{tiers_all['high_risk']['count']} are HIGH RISK DELAY (Tight 0–7d buffer / Critical Path / Key VDRL hold), "
            f"{tiers_all['low_risk']['count']} are LOW RISK, and {tiers_all['healthy']['count']} are VERY GOOD / HEALTHY."
        ),
        "vendor_management_finding": (
            f"Master Schedule Critical Path (TF=0) runs through Pedestal Crane (PDC), Diesel Transfer Pump (DTP), Control Valve (CV), Actuated Choke Valve (ACV), and Primary Steel Plate (PLT). "
            f"WPR Attachment 3 & 4 analysis shows Phase 2A Key Document Approval (GAD/Datasheet/BOM) is the primary gate holding sub-vendor material PO releases."
        ),
        "document_control_finding": (
            f"VDRL Vendor Submissions stand at {total_vdrl_submitted}/{total_vdrl_docs} ({portfolio_vdrl_pct:.1f}%) across {len(vdrl_map)} registers, "
            f"with {total_vdrl_overdue} overdue document(s) ({total_key_vdrl_overdue} Key Docs, {total_prod_vdrl_overdue} Production-Impact). "
            f"Contractual Client Review Log ({client_log.get('source_file', 'D2 Log')}) tracks {client_log.get('total_submittals', 0)} submittals "
            f"({client_log.get('status_counts', {}).get('In Review', 0)} In Review, {client_log.get('return_counts', {}).get('APPR', 0)} APPR, "
            f"{client_log.get('return_counts', {}).get('AWC', 0)} AWC, {client_log.get('overdue_reviews', 0)} Overdue Review)."
        ),
    }

    # Build Incorporated Uploaded Files summary for the dashboard notification banner
    incorporated_files: List[Dict[str, Any]] = []
    for w_code, w_rec in sorted(wpr_map.items()):
        incorporated_files.append({
            "category": "WPR",
            "type": "WPR",
            "package": w_code,
            "document_no": w_rec.get("document_no") or f"{w_code} {w_rec.get('report_no', 'WPR')}",
            "report_no": w_rec.get("report_no", ""),
            "filename": w_rec.get("source_file", ""),
            "source_file": w_rec.get("source_file", ""),
            "cut_off_date": w_rec.get("cut_off_date", ""),
            "planned_pct": w_rec.get("overall_planned_pct", 0.0),
            "actual_pct": w_rec.get("overall_actual_pct", 0.0),
            "summary": f"{w_code} {w_rec.get('report_no', '')} (Cut-off: {w_rec.get('cut_off_date', '')} • Actual {w_rec.get('overall_actual_pct', 0)}% vs Plan {w_rec.get('overall_planned_pct', 0)}%)",
        })
    for m_code_k, m_rec in sorted(mom_map.items()):
        if m_rec.get("available"):
            incorporated_files.append({
                "category": "MOM",
                "type": "MOM",
                "package": m_code_k,
                "document_no": m_rec.get("document_no", ""),
                "report_no": m_rec.get("mom_no", ""),
                "filename": m_rec.get("source_file", ""),
                "source_file": m_rec.get("source_file", ""),
                "cut_off_date": m_rec.get("cut_off_date", ""),
                "detail": f"{m_rec.get('total_items', 0)} items ({m_rec.get('open_count', 0)} OPEN, {m_rec.get('closed_count', 0)} CLOSED, {m_rec.get('info_count', 0)} INFO) • {m_rec.get('source_file', '')}",
                "summary": f"{m_code_k} {m_rec.get('mom_no', '')} (Date: {m_rec.get('cut_off_date', '')} • {m_rec.get('total_items', 0)} items: {m_rec.get('open_count', 0)} OPEN, {m_rec.get('closed_count', 0)} CLOSED, {m_rec.get('info_count', 0)} INFO)",
            })
    incorporated_files.append({
        "category": "SCD",
        "type": "SCD",
        "package": proj,
        "document_no": master_schedule.get("document_no", active_master_file),
        "report_no": f"Rev. {master_schedule.get('revision', 'B1')}",
        "filename": active_master_file,
        "source_file": active_master_file,
        "cut_off_date": "2026-10-16" if is_ask else "2026-09-16",
        "summary": f"Master Schedule {active_master_file} (Rev. {master_schedule.get('revision', 'B1')})",
    })
    incorporated_files.append({
        "category": "PSR",
        "type": "PSR",
        "package": proj,
        "document_no": active_psr_file,
        "report_no": "Procurement Plan",
        "filename": active_psr_file,
        "source_file": active_psr_file,
        "cut_off_date": "2026-10-02",
        "summary": f"Procurement Plan ({active_psr_file})",
    })

    lookahead_cutoff = next(
        (p["cut_off_date"] for p in packages_list if p["package_code"] in ("PLC", "CCV") and p.get("cut_off_date")),
        "2026-10-16" if is_ask else "2026-09-16",
    )

    return {
        "project": proj,
        "project_full_name": (
            "EPCI of Wellhead Platforms for Aung Sinkha Development Project Phase 1A (EPC-01)"
            if is_ask
            else "EPCI of Wellhead Platforms & Pipelines for Zawtika Development Project Phase 1F (Z1F)"
        ),
        "mode": mode,
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "executive_findings_summary": executive_findings_summary,
        "sources_summary": {
            "master_schedule_file": active_master_file,
            "master_schedule_rev": master_schedule.get("revision", "B1"),
            "wpr_folder": "WPR/",
            "wpr_files_count": len(wpr_map),
            "mom_folder": "MOM/",
            "mom_files_count": len(mom_map),
            "vdrl_folder": "VDRL/",
            "vdrl_packages_count": len(vdrl_map),
            "kom_folder": "KOM/",
            "kom_packages_count": kom_available_count,
            "psr_file": active_psr_file,
            "client_doc_log_file": client_log.get("source_file", ""),
            "client_doc_log_updated": client_log.get("last_updated", ""),
            "incorporated_files": incorporated_files,
        },
        "portfolio_kpis": {
            "total_packages": total_packages,
            "core_packages_count": core_packages_count,
            "psr_only_equipment_count": psr_only_equipment_count,
            "live_wpr_packages": live_wpr_count,
            "live_mom_packages": live_mom_count,
            "kom_populated_packages": kom_available_count,
            "overall_planned_pct": overall_planned_pct,
            "overall_actual_pct": overall_actual_pct,
            "overall_variance_pct": round(overall_actual_pct - overall_planned_pct, 1),
            "ros_delayed_packages": ros_delayed_packages,
            "ros_delayed_packages_count": ros_delayed_packages,
            "ros_delayed_packages_list": ros_delayed_list,
            "already_delayed_count": tiers_all["already_delayed"]["count"],
            "high_risk_count": tiers_all["high_risk"]["count"],
            "low_risk_count": tiers_all["low_risk"]["count"],
            "healthy_count": tiers_all["healthy"]["count"],
            "core_already_delayed_count": tiers_core["already_delayed"]["count"],
            "core_high_risk_count": tiers_core["high_risk"]["count"],
            "core_low_risk_count": tiers_core["low_risk"]["count"],
            "core_healthy_count": tiers_core["healthy"]["count"],
            "vdrl_total_docs": total_vdrl_docs,
            "vdrl_submitted_docs": total_vdrl_submitted,
            "vdrl_overdue_docs": total_vdrl_overdue,
            "vdrl_submitted_pct": portfolio_vdrl_pct,
            "client_log_total_submittals": client_log.get("total_submittals", 0),
            "client_log_in_review": client_log.get("status_counts", {}).get("In Review", 0),
            "client_log_appr": client_log.get("return_counts", {}).get("APPR", 0),
            "client_log_awc": client_log.get("return_counts", {}).get("AWC", 0),
            "client_log_overdue": client_log.get("overdue_reviews", 0),
            "total_areas_of_concern": total_concerns,
        },
        "project_overall_scurve": project_overall_scurve,
        "client_doc_log_overview": {
            "available": client_log.get("available", False),
            "source_file": client_log.get("source_file", ""),
            "last_updated": client_log.get("last_updated", ""),
            "total_submittals": client_log.get("total_submittals", 0),
            "status_counts": client_log.get("status_counts", {}),
            "return_counts": client_log.get("return_counts", {}),
            "overdue_reviews": client_log.get("overdue_reviews", 0),
            "package_breakdown": [
                {
                    "package_code": k,
                    "raw_package_code": v["raw_package_code"],
                    "package_name": v["package_name"],
                    "total_submittals": v["total_submittals"],
                    "in_review": v["in_review"],
                    "submitted": v["submitted"],
                    "in_approval": v["in_approval"],
                    "appr": v["appr"],
                    "awc": v["awc"],
                    "appr_c": v["appr_c"],
                    "rec": v["rec"],
                    "pending_return": v["pending_return"],
                    "overdue_reviews": v["overdue_reviews"],
                }
                for k, v in sorted(
                    client_log.get("by_package", {}).items(),
                    key=lambda item: item[1]["total_submittals"],
                    reverse=True,
                )
            ],
            "recent_submittals": client_log.get("all_rows", [])[:271],
        },
        "packages": packages_list,
        "lookahead_hub": build_project_lookahead_hub(
            packages_list=packages_list,
            master_schedule=master_schedule,
            cut_off_date=lookahead_cutoff,
        ),
        "construction_monitoring": build_construction_monitoring(
            proj=proj,
            packages_list=packages_list,
            master_schedule=master_schedule,
        ),
    }


def build_construction_monitoring(
    proj: str,
    packages_list: List[Dict[str, Any]],
    master_schedule: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """Build Construction Monitoring payload from Master Schedule + Vendor Impact reconciliation.

    1) Vendor impact concerns to yard construction (ETA vs Yard Erection / ROS / Critical Path).
    2) 14-day forecast construction activities ([Today, Today + 14d]).
    3) Simplified Gantt window ([Today - 7d, Today + 14d]) + all Master Schedule construction activities.
    """
    today = datetime.date.today()
    today_iso = today.isoformat()
    gantt_start = today - datetime.timedelta(days=7)
    gantt_end = today + datetime.timedelta(days=14)
    gantt_start_iso = gantt_start.isoformat()
    gantt_end_iso = gantt_end.isoformat()

    ms = master_schedule or {}
    ms_file = ms.get("source_file") or (
        "MM-ASK-1A-GEN01-PJM-SCD-0001.pdf" if proj == "ASK" else "MM-ZTK-1F-GEN-PJM-SCD-0002.pdf"
    )
    ms_rev = ms.get("revision") or ("B1" if proj == "ASK" else "C1")

    # 1. Vendor-to-Construction Impact Concerns
    vendor_impact_concerns: List[Dict[str, Any]] = []
    for p in packages_list:
        hl = p.get("package_highlights") or {}
        eta = hl.get("effective_delivery_date") or p.get("eta_date") or ""
        ros = p.get("ros_date") or ""
        erection = p.get("yard_erection_date") or ""
        is_crit = bool(p.get("is_critical_path"))

        erection_impact = _date_diff_days(erection, eta) if (eta and erection) else None
        ros_delay = hl.get("ros_delay_days")
        if ros_delay is None and eta and ros:
            ros_delay = _date_diff_days(ros, eta)
        master_delay = hl.get("master_delay_days")
        ros_margin = hl.get("ros_margin_days")

        reasons: List[str] = []
        severity = ""

        if erection_impact is not None and erection_impact > 0:
            reasons.append(f"ETA ({eta}) slips past Yard Erection ({erection}) by +{erection_impact}d")
            severity = "CRITICAL" if (is_crit or erection_impact >= 14) else "HIGH"

        if hl.get("is_ros_delayed") or (ros_delay is not None and ros_delay > 0):
            d_val = ros_delay if (ros_delay is not None and ros_delay > 0) else 0
            reasons.append(f"Delivery ETA slips past ROS ({ros}) by +{d_val}d")
            if severity != "CRITICAL":
                severity = "CRITICAL" if is_crit else "HIGH"

        if hl.get("is_master_delayed") or (master_delay is not None and master_delay > 0):
            m_val = master_delay if (master_delay is not None and master_delay > 0) else 0
            reasons.append(f"Behind Master Schedule target ETA by +{m_val}d")
            if not severity:
                severity = "HIGH" if is_crit else "MEDIUM"

        if not reasons and is_crit and ros_margin is not None and ros_margin <= 14:
            reasons.append(f"Critical-path package with tight {ros_margin}d float before ROS/Erection")
            severity = "MEDIUM"

        if reasons:
            vendor_impact_concerns.append({
                "package_code": p.get("package_code", ""),
                "package_name": p.get("package_name", ""),
                "discipline": p.get("discipline", ""),
                "is_critical_path": is_crit,
                "health_tier": p.get("health_tier", ""),
                "eta_date": eta,
                "ros_date": ros,
                "yard_erection_date": erection,
                "yard_erection_activity": hl.get("yard_erection_activity") or "",
                "erection_impact_days": erection_impact,
                "ros_delay_days": ros_delay,
                "master_delay_days": master_delay,
                "ros_margin_days": ros_margin,
                "severity": severity or "MEDIUM",
                "impact_reason": " • ".join(reasons),
                "mitigation": hl.get("smart_mitigation") or hl.get("smart_issue") or "Expedite vendor fabrication & align yard erection sequence.",
            })

    sev_rank = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2}
    vendor_impact_concerns.sort(key=lambda x: (
        sev_rank.get(x["severity"], 3),
        0 if x["is_critical_path"] else 1,
        -(x["erection_impact_days"] if x["erection_impact_days"] is not None else -9999),
        -(x["ros_delay_days"] if x["ros_delay_days"] is not None else -9999),
        x["package_code"],
    ))

    # 2. Master Schedule Construction Activities (-7d..+14d Gantt, 14d Forecast, & All)
    construction_keywords = (
        "ERECT", "FABRICAT", "INSTALL", "HOOK-UP", "HOOKUP", "JACKET",
        "TOPSIDE", "ONSHORE", "YARD", "COMMISSION", "SAIL-AWAY", "SAILAWAY",
        "LIFT", "PILING", "LOADOUT", "LOAD-OUT", "CONSTRUCTION", "ASSEMBLY",
    )
    exclude_keywords = ("PO ISSUE", "VDR", "TBE", "CBE", "RFQ", "BID EVAL", "SUBMITTAL")

    def _is_construction_act(act: Dict[str, Any]) -> bool:
        cat = str(act.get("category") or "").strip()
        if cat == "Fabrication & Erection":
            return True
        if cat in ("Procurement", "Engineering & Design"):
            return False
        upper_name = str(act.get("activity_name") or "").upper()
        if any(ex in upper_name for ex in exclude_keywords):
            return False
        return any(kw in upper_name for kw in construction_keywords)

    all_activities: List[Dict[str, Any]] = []
    gantt_activities: List[Dict[str, Any]] = []
    forecast_14d_activities: List[Dict[str, Any]] = []

    raw_master_acts = _clean_and_dedup_master_activities(ms.get("all_activities") or [])
    for act in raw_master_acts:
        if not _is_construction_act(act):
            continue
        s = str(act.get("start_date") or "")[:10]
        f = str(act.get("finish_date") or "")[:10]
        eff_s = s or f
        eff_f = f or s
        if not eff_s or not eff_f:
            continue

        if eff_f < today_iso:
            status = "Completed / Past"
        elif eff_s <= today_iso <= eff_f:
            status = "In Progress"
        else:
            status = "Upcoming"

        item = {
            "activity_id": str(act.get("activity_id") or ""),
            "activity_name": str(act.get("activity_name") or ""),
            "category": str(act.get("category") or "Fabrication & Erection"),
            "start_date": eff_s,
            "finish_date": eff_f,
            "duration_days": act.get("duration_days"),
            "total_float_days": act.get("total_float_days") if act.get("total_float_days") is not None else act.get("total_float"),
            "is_critical_path": bool(act.get("is_critical_path")),
            "status": status,
            "page": act.get("page"),
            "source": ms_file,
        }
        all_activities.append(item)

        if eff_s <= gantt_end_iso and eff_f >= gantt_start_iso:
            gantt_activities.append(item)
        if eff_s <= gantt_end_iso and eff_f >= today_iso:
            forecast_14d_activities.append(item)

    def _act_sort_key(x: Dict[str, Any]) -> Tuple[int, str, str, str]:
        return (
            0 if x["is_critical_path"] else 1,
            x["start_date"] or "9999-99-99",
            x["finish_date"] or "9999-99-99",
            x["activity_id"],
        )

    gantt_activities.sort(key=_act_sort_key)
    forecast_14d_activities.sort(key=_act_sort_key)
    all_activities.sort(key=_act_sort_key)

    # 3. Transportation & Installation (T&I) Timeline + Detailed Activities (Attachment 3)
    ti_timeline = list((ms.get("project_key_dates") or {}).get("ti_timeline") or [])
    ti_activities: List[Dict[str, Any]] = []
    for act in raw_master_acts:
        aid = str(act.get("activity_id") or "")
        aname = str(act.get("activity_name") or "")
        u_id = aid.upper()
        u_name = aname.upper()
        if any(k in u_id for k in ("WU1.6", "WU1.7", "WU3.1")) or any(
            k in u_name for k in ("T&I", "BARGE", "HLV", "OFFSHORE", "LOAD OUT", "LOADOUT", "SEA FASTENING", "SEAFASTENING", "GREENFIELD HUC")
        ):
            s = str(act.get("start_date") or "")[:10]
            f = str(act.get("finish_date") or "")[:10]
            eff_s = s or f
            eff_f = f or s
            if not eff_s or not eff_f:
                continue
            ti_activities.append({
                "activity_id": aid,
                "activity_name": aname,
                "category": "Transportation & Installation (T&I)",
                "start_date": eff_s,
                "finish_date": eff_f,
                "duration_days": act.get("duration_days"),
                "is_critical_path": bool(act.get("is_critical_path")),
                "status": "Upcoming",
                "source": ms_file,
            })
    ti_activities.sort(key=lambda x: (x["start_date"], x["finish_date"], x["activity_id"]))

    return {
        "today_date": today_iso,
        "gantt_window_start": gantt_start_iso,
        "gantt_window_end": gantt_end_iso,
        "master_schedule_file": ms_file,
        "master_schedule_rev": ms_rev,
        "vendor_impact_concerns": vendor_impact_concerns,
        "gantt_activities": gantt_activities,
        "forecast_14d_activities": forecast_14d_activities,
        "all_activities": all_activities,
        "ti_timeline": ti_timeline,
        "ti_activities": ti_activities,
        "summary": {
            "vendor_impact_count": len(vendor_impact_concerns),
            "vendor_critical_count": sum(1 for x in vendor_impact_concerns if x["severity"] == "CRITICAL"),
            "gantt_window_count": len(gantt_activities),
            "forecast_14d_count": len(forecast_14d_activities),
            "forecast_14d_critical_count": sum(1 for x in forecast_14d_activities if x["is_critical_path"]),
            "ti_count": len(ti_activities),
            "all_construction_count": len(all_activities),
        },
    }


def _reconcile_single_package(
    proj: str,
    pkg_code: str,
    pkg_weight_pct: float,
    meta: Dict[str, str],
    kom_detail: Optional[ManufacturerDetail],
    vdrl_data: Optional[Dict[str, Any]],
    client_pkg_data: Optional[Dict[str, Any]],
    psr_data: Optional[Dict[str, Any]],
    live_wpr: Optional[Dict[str, Any]],
    sample_wpr: Optional[Dict[str, Any]],
    mode: str,
    master_schedule: Optional[Dict[str, Any]] = None,
    mom_detail: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Reconcile all 9 required dashboard elements + highlights + Attachments 3 & 4 + MOM for a single package."""
    has_live_wpr = live_wpr is not None
    has_live_mom = bool(mom_detail and mom_detail.get("available"))
    using_sample_wpr = (not has_live_wpr) and (mode == "sample")

    # Element 1 & 2: Project & Package Name / Metadata
    pkg_name = (
        (vdrl_data or {}).get("package_name")
        or meta.get("name")
        or (psr_data or {}).get("package_name")
        or pkg_code
    )
    discipline = meta.get("discipline") or (vdrl_data or {}).get("discipline") or "GEN"
    rfq_no = (psr_data or {}).get("rfq_no") or meta.get("rfq", "")
    equipment_tag = (
        (live_wpr or {}).get("equipment_tag")
        or meta.get("tag")
        or "WP-01 / WP-02"
    )

    # Element 3: Manufacturer Detail (Strictly from KOM/; blank if not available in KOM/)
    mfr = kom_detail.to_dict() if kom_detail else ManufacturerDetail(available=False).to_dict()
    po_no = (
        mfr.get("po_number")
        or (live_wpr or {}).get("po_number")
        or (psr_data or {}).get("po_hint")
        or ""
    )
    po_date = mfr.get("po_date") or (live_wpr or {}).get("po_date") or ""

    # Element 6: Status Vendor Document (6A. VDRL Submission + 6B. Contractual Client Review Log)
    vdrl_status = _build_vdrl_status(vdrl_data)
    client_doc_status = _build_client_doc_status(client_pkg_data)

    # Element 4: Key Milestones (Priority Rule: Live WPR Priority #1 > PSR Priority #2)
    milestones = _reconcile_milestones(
        live_wpr=live_wpr,
        sample_wpr=sample_wpr if using_sample_wpr else None,
        psr_data=psr_data,
        kom_detail=kom_detail,
        vdrl_status=vdrl_status,
    )

    cut_off_date = (
        (live_wpr or {}).get("cut_off_date")
        or (mom_detail or {}).get("cut_off_date")
        or vdrl_status.get("today_date")
        or "2026-10-16"
    )

    # Fixed Contractual ROS Date & Governing ETA Priority:
    # User Rule: "The ETA date should follow the Gantt chart (if available). And if it is not aligned with milestone, please highlight and notify."
    ros_date = (psr_data or {}).get("ros_date", "")
    psr_eta_date = (psr_data or {}).get("eta_date", "")
    m11_receipt = next((m for m in milestones if m["key"] == "receipt_worksite"), None)
    m11_date = (
        (m11_receipt or {}).get("actual_date")
        or (m11_receipt or {}).get("forecast_date")
        or (m11_receipt or {}).get("planned_date")
        or (live_wpr or {}).get("target_delivery")
        or ""
    )

    # Check Live WPR Gantt Chart for Delivery / Shipment / CIF task finish date
    gantt_eta_date = ""
    gantt_eta_task_name = ""
    if has_live_wpr and live_wpr is not None:
        raw_gantt = live_wpr.get("gantt_tasks", [])
        delivery_candidates: List[Tuple[str, str]] = []
        for gt in raw_gantt:
            t_name = str(gt.get("name") or gt.get("task_name") or "").strip()
            t_cat = str(gt.get("phase_category") or gt.get("phase") or "").strip()
            f_dt = str(gt.get("finish_date") or gt.get("finish") or "").strip()
            tl = t_name.lower()
            if f_dt and len(f_dt) >= 10 and f_dt[:4].isdigit():
                if any(k in tl for k in ("sea freight", "transit to china", "cif china", "cif)", "shipment - cif", "receipt at worksite")):
                    delivery_candidates.append((f_dt[:10], t_name))
        if delivery_candidates:
            # Prefer non-summary leaf task with the latest delivery finish date
            delivery_candidates.sort(key=lambda x: (x[0], 1 if not x[1][:2].endswith(".") else 0))
            gantt_eta_date, gantt_eta_task_name = delivery_candidates[-1]

    has_milestone_discrepancy = False
    milestone_discrepancy_days = 0
    milestone_discrepancy_note = ""

    if has_live_wpr and gantt_eta_date:
        effective_delivery_date = gantt_eta_date
        delivery_source = f"WPR Gantt Chart ({gantt_eta_task_name})"
        if m11_date and gantt_eta_date != m11_date:
            diff_d = _date_diff_days(m11_date, gantt_eta_date) or 0
            if diff_d != 0:
                has_milestone_discrepancy = True
                milestone_discrepancy_days = diff_d
                milestone_discrepancy_note = (
                    f"⚠️ Milestone vs. Gantt Discrepancy ({'+' if diff_d > 0 else ''}{diff_d}d): "
                    f"WPR Key Milestone #10 ('Receipt at worksite') states {m11_date} (Loading/Departure date), "
                    f"whereas WPR Attachment 3 Gantt Chart ('{gantt_eta_task_name}') finishes on {gantt_eta_date}. "
                    f"Governing ETA follows Gantt Chart ({gantt_eta_date})."
                )
                if m11_receipt is not None:
                    m11_receipt["forecast_date"] = gantt_eta_date
                    m11_receipt["variance_days"] = diff_d
                    m11_receipt["status"] = f"Gantt ETA {gantt_eta_date} (+{diff_d}d vs MS)"
                    m11_receipt["tone"] = "warning"
                    m11_receipt["deviation_note"] = f"Gantt Task '{gantt_eta_task_name}': {gantt_eta_date} (+{diff_d}d transit vs MS {m11_date})"
    elif has_live_wpr and m11_date:
        effective_delivery_date = m11_date
        delivery_source = "WPR #11 Receipt at Worksite (ETA)"
    elif m11_date:
        effective_delivery_date = m11_date
        delivery_source = "PSR #11 Receipt at Worksite (ETA)"
    elif psr_eta_date:
        effective_delivery_date = psr_eta_date
        delivery_source = "PSR ETA / Forecast Delivery"
    else:
        effective_delivery_date = ""
        delivery_source = "—"

    # Element 5 & Element 7: Dynamic Package-Specific Milestone Progress & S-Curve
    if has_live_wpr and live_wpr is not None:
        milestone_progress = live_wpr.get("milestone_progress", [])
        overall_planned = live_wpr.get("overall_planned_pct", 0.0)
        overall_actual = live_wpr.get("overall_actual_pct", 0.0)
        raw_scurve = live_wpr.get("scurve", {})
        if isinstance(raw_scurve, list):
            scurve = {
                "current_week": "W12",
                "cut_off_date": cut_off_date,
                "planned_to_date_pct": overall_planned,
                "actual_to_date_pct": overall_actual,
                "variance_pct": round(overall_actual - overall_planned, 2),
                "points": raw_scurve,
            }
        else:
            scurve = raw_scurve
        this_period_acts = list(live_wpr.get("this_period_activities", []))
        if not milestone_progress:
            milestone_progress, _, _, _, _ = _build_dynamic_package_progress_and_scurve(
                pkg_code=pkg_code,
                pkg_name=pkg_name,
                discipline=discipline,
                milestones=milestones,
                vdrl_status=vdrl_status,
                client_doc_status=client_doc_status,
                cut_off_date=cut_off_date,
                ros_date=ros_date,
                effective_delivery_date=effective_delivery_date,
            )
    elif using_sample_wpr or (psr_data is not None) or vdrl_status.get("available"):
        milestone_progress, overall_planned, overall_actual, scurve, this_period_acts = (
            _build_dynamic_package_progress_and_scurve(
                pkg_code=pkg_code,
                pkg_name=pkg_name,
                discipline=discipline,
                milestones=milestones,
                vdrl_status=vdrl_status,
                client_doc_status=client_doc_status,
                cut_off_date=cut_off_date,
                ros_date=ros_date,
                effective_delivery_date=effective_delivery_date,
            )
        )
    else:
        milestone_progress = []
        overall_planned = 0.0
        overall_actual = 0.0
        scurve = {
            "current_week": "—",
            "cut_off_date": cut_off_date,
            "planned_to_date_pct": 0.0,
            "actual_to_date_pct": 0.0,
            "variance_pct": 0.0,
            "points": [],
        }
        this_period_acts = []

    overall_variance = round(overall_actual - overall_planned, 2)

    # Resolve Master Schedule (MM-ASK-1A-GEN01-PJM-SCD-0001.pdf / MM-ZTK-1F-GEN-PJM-SCD-0002.pdf) Priority #1 entry
    master_entry = match_master_schedule_entry(
        master_sched=master_schedule or {},
        pkg_code=pkg_code,
        pkg_name=pkg_name,
        ros_date=ros_date,
    )

    # Comment #4.1, #6 & Master Plan Priority: Highlight Key Information & 4-Tier Health Classification
    package_highlights = _build_package_highlights(
        pkg_code=pkg_code,
        pkg_name=pkg_name,
        milestones=milestones,
        cut_off_date=cut_off_date,
        effective_delivery_date=effective_delivery_date,
        delivery_source=delivery_source,
        ros_date=ros_date,
        master_entry=master_entry,
        vdrl_status=vdrl_status,
        client_doc_status=client_doc_status,
        has_live_wpr=has_live_wpr,
        live_wpr=live_wpr,
    )
    package_highlights["milestone_eta_date"] = m11_date
    package_highlights["gantt_eta_date"] = gantt_eta_date
    package_highlights["gantt_eta_task_name"] = gantt_eta_task_name
    package_highlights["has_milestone_discrepancy"] = has_milestone_discrepancy
    package_highlights["milestone_discrepancy_days"] = milestone_discrepancy_days
    package_highlights["milestone_discrepancy_note"] = milestone_discrepancy_note

    # Comment #5: WPR Attachment 3 (Overall Fabrication Detail Plan Gantt) & Attachment 4 (Supplier Material Expediting)
    attachment3_summary, attachment4_summary = _build_attachments_3_and_4_summary(
        pkg_code=pkg_code,
        pkg_name=pkg_name,
        discipline=discipline,
        milestones=milestones,
        vdrl_status=vdrl_status,
        client_doc_status=client_doc_status,
        live_wpr=live_wpr,
        package_highlights=package_highlights,
        cut_off_date=cut_off_date,
    )

    # Element 8: Dynamic Package-Specific Look Ahead Plan (7–14 Days)
    look_ahead = _build_look_ahead_plan(
        pkg_code=pkg_code,
        pkg_name=pkg_name,
        live_wpr=live_wpr,
        vdrl_data=vdrl_data,
        client_doc_status=client_doc_status,
        milestones=milestones,
        attachment3_summary=attachment3_summary,
        attachment4_summary=attachment4_summary,
        cut_off_date=cut_off_date,
        is_sample=using_sample_wpr,
    )

    # Element 9: Area of Concern (Incorporating Master Plan & ROS Date Delay, Gantt vs Milestone Discrepancy, WPR & MOM Concerns, VDRL & Client Log)
    areas_of_concern = _build_areas_of_concern(
        pkg_code=pkg_code,
        live_wpr=live_wpr,
        vdrl_status=vdrl_status,
        client_doc_status=client_doc_status,
        milestones=milestones,
        package_highlights=package_highlights,
        attachment3_summary=attachment3_summary,
        attachment4_summary=attachment4_summary,
        is_sample=using_sample_wpr,
        mom_detail=mom_detail,
    )

    # Map 4-Tier Health Classification onto health_tier, health_status, and health_tone
    health_tier = package_highlights["health_tier"]
    health_status = package_highlights["health_status_label"]
    health_tone = package_highlights["health_tone"]

    next_milestone = next(
        (m for m in milestones if not m["actual_date"] and m["planned_date"]),
        milestones[-1] if milestones else None,
    )

    # Build Reference Document & Cut-Off Date stamps for each of the 9 Elements + MOM
    scd_doc_no = (master_schedule or {}).get("document_no") or ("MM-ASK-1A-GEN01-PJM-SCD-0001" if proj == "ASK" else "MM-ZTK-1F-GEN-PJM-SCD-0002")
    wpr_doc_no = (live_wpr or {}).get("document_no") or (live_wpr or {}).get("source_file") or ""
    wpr_cutoff = (live_wpr or {}).get("cut_off_date") or cut_off_date
    mom_doc_no = (mom_detail or {}).get("document_no") or (mom_detail or {}).get("source_file") or ""
    mom_cutoff = (mom_detail or {}).get("cut_off_date") or ""
    vdrl_doc_no = vdrl_status.get("source_file") or "No VDRL Uploaded"
    vdrl_cutoff = vdrl_status.get("today_date") or cut_off_date
    kom_doc_no = mfr.get("kom_source_file") or "No KOM Uploaded"
    kom_cutoff = mfr.get("kom_date") or mfr.get("po_date") or "—"
    psr_doc_no = (psr_data or {}).get("source_file") or ("Attachment 2-Procurement Plan-ASK" if proj == "ASK" else "Attachment 1 - Procurement Plan-Z1F")

    primary_sched_ref = (
        f"{wpr_doc_no} & {scd_doc_no}"
        if has_live_wpr
        else f"{scd_doc_no} & {psr_doc_no} (SCD/PSR Only — No WPR)"
    )
    primary_prog_ref = (
        f"{wpr_doc_no} ({(live_wpr or {}).get('report_no', 'WPR')})"
        if has_live_wpr
        else f"SCD ({scd_doc_no}) & VDRL ({vdrl_doc_no}) — No WPR Uploaded"
    )
    concern_refs = []
    if has_live_wpr:
        concern_refs.append(wpr_doc_no)
    if has_live_mom:
        concern_refs.append(mom_doc_no)
    concern_refs.append(scd_doc_no)
    if vdrl_status.get("available"):
        concern_refs.append(vdrl_doc_no)

    element_metadata = {
        "element_1_2": {"ref_doc": primary_sched_ref, "cut_off_date": cut_off_date},
        "element_3": {"ref_doc": kom_doc_no, "cut_off_date": kom_cutoff},
        "element_4": {"ref_doc": primary_sched_ref, "cut_off_date": wpr_cutoff if has_live_wpr else cut_off_date},
        "element_5": {"ref_doc": primary_prog_ref, "cut_off_date": wpr_cutoff if has_live_wpr else vdrl_cutoff},
        "element_6": {"ref_doc": f"{vdrl_doc_no} & D2 Log", "cut_off_date": vdrl_cutoff},
        "element_7": {"ref_doc": primary_prog_ref, "cut_off_date": wpr_cutoff if has_live_wpr else vdrl_cutoff},
        "element_8": {"ref_doc": f"{wpr_doc_no + ' & ' if has_live_wpr else ''}{scd_doc_no} & VDRL", "cut_off_date": cut_off_date},
        "element_9": {"ref_doc": " • ".join(concern_refs), "cut_off_date": mom_cutoff or wpr_cutoff or cut_off_date},
        "mom": {"ref_doc": mom_doc_no or "No MOM Uploaded", "cut_off_date": mom_cutoff or "—"},
    }

    mom_summary = mom_detail if has_live_mom else {
        "available": False,
        "package_code": pkg_code,
        "mom_no": "—",
        "document_no": "No MOM Uploaded",
        "source_file": "",
        "meeting_date": "—",
        "cut_off_date": "—",
        "total_items": 0,
        "open_count": 0,
        "closed_count": 0,
        "info_count": 0,
        "overall_situation_summary": f"No Minutes of Meeting (MOM) uploaded yet for {pkg_name} ({pkg_code}). Displaying available Master Schedule (SCD) and VDRL status only.",
        "key_concerns": [],
        "open_items": [],
        "action_items": [],
        "procurement_status": [],
    }

    return {
        # Element 1: Project
        "project": proj,
        # Element 2: Package Name & Header Metadata
        "package_code": pkg_code,
        "package_name": pkg_name,
        "package_weight_pct": pkg_weight_pct,
        "discipline": discipline,
        "rfq_no": rfq_no,
        "equipment_tag": equipment_tag,
        "po_number": po_no,
        "po_date": po_date,
        "report_no": (live_wpr or {}).get("report_no")
        or (f"WPR-{pkg_code}-001 (Template)" if using_sample_wpr else "SCD / VDRL Baseline"),
        "revision": (live_wpr or {}).get("revision") or ("00" if using_sample_wpr else "—"),
        "cut_off_date": cut_off_date,
        "wpr_source_file": (live_wpr or {}).get("source_file", ""),
        "wpr_document_no": wpr_doc_no,
        "mom_source_file": (mom_detail or {}).get("source_file", ""),
        "mom_document_no": mom_doc_no,
        "mom_cut_off_date": mom_cutoff,
        "has_live_wpr": has_live_wpr,
        "has_live_mom": has_live_mom,
        "using_sample_wpr": using_sample_wpr,
        "wpr_active": has_live_wpr or using_sample_wpr or (psr_data is not None),
        "data_priority_badge": (
            "Priority #1: Live WPR"
            if has_live_wpr
            else ("Dynamic WPR Template + PSR" if using_sample_wpr else "SCD & VDRL Baseline (No WPR Uploaded)")
        ),
        "element_metadata": element_metadata,
        "health_tier": health_tier,
        "health_status": health_status,
        "health_tone": health_tone,
        "ros_date": package_highlights["ros_date"],
        "psr_ros_date": ros_date,
        "master_eta_date": package_highlights["master_eta_date"],
        "yard_erection_date": package_highlights["yard_erection_date"],
        "is_critical_path": package_highlights["is_critical_path"],
        "eta_date": effective_delivery_date,
        "milestone_eta_date": m11_date,
        "gantt_eta_date": gantt_eta_date,
        "has_milestone_discrepancy": has_milestone_discrepancy,
        "milestone_discrepancy_days": milestone_discrepancy_days,
        "milestone_discrepancy_note": milestone_discrepancy_note,
        "delivery_source": delivery_source,
        "float_days": package_highlights["ros_margin_days"],
        # Comment #4.1 & #6: Package Key Highlight Banner Data
        "package_highlights": package_highlights,
        # Element 3: Manufacturer Detail (Strictly from KOM/)
        "manufacturer_detail": mfr,
        # Element 4: Key Milestone
        "key_milestones": milestones,
        "next_milestone": next_milestone,
        # Element 5: Progress Tracking by Key Milestone
        "milestone_progress": milestone_progress,
        "overall_progress": {
            "planned_pct": overall_planned,
            "actual_pct": overall_actual,
            "variance_pct": overall_variance,
            "this_period_activities": this_period_acts,
        },
        # Element 6: Status Vendor Document (6A VDRL + 6B Client Log)
        "vdrl_status": vdrl_status,
        "client_doc_status": client_doc_status,
        # Element 7: Overall Progress (S-Curve)
        "scurve": scurve,
        # Element 8: Look Ahead Plan 7-14 Days
        "look_ahead_plan": look_ahead,
        # Element 9: Area of Concern
        "areas_of_concern": areas_of_concern,
        # MOM Summary & Action Items
        "mom_summary": mom_summary,
        # Comment #5: WPR Attachment 3 (Fabrication Gantt) & Attachment 4 (Supplier Material Expediting)
        "attachment3_summary": attachment3_summary,
        "attachment4_summary": attachment4_summary,
        "gantt_tasks": attachment3_summary.get("tasks", []),
    }


def _reconcile_milestones(
    live_wpr: Optional[Dict[str, Any]],
    sample_wpr: Optional[Dict[str, Any]],
    psr_data: Optional[Dict[str, Any]],
    kom_detail: Optional[ManufacturerDetail],
    vdrl_status: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """Priority Rule:
    1. Live WPR takes Priority #1.
    2. PSR takes Priority #2 (supplemented by KOM actual date and VDRL submission date if available).
    """
    wpr_by_key: Dict[str, Dict[str, Any]] = {}
    if live_wpr:
        for m in live_wpr.get("milestones", []):
            wpr_by_key[m["key"]] = m

    sample_by_key: Dict[str, Dict[str, Any]] = {}
    if sample_wpr:
        for m in sample_wpr.get("milestones", []):
            sample_by_key[m["key"]] = m

    psr_plan = (psr_data or {}).get("plan", {})
    psr_fc = (psr_data or {}).get("forecast", {})
    psr_act = (psr_data or {}).get("actual", {})
    ros_date = (psr_data or {}).get("ros_date", "")

    reconciled: List[Dict[str, Any]] = []
    for num, label, key in STANDARD_MILESTONES:
        w_item = wpr_by_key.get(key)
        s_item = sample_by_key.get(key)

        p_psr = psr_plan.get(key, "")
        f_psr = psr_fc.get(key, "")
        a_psr = psr_act.get(key, "")

        # Enrich actuals from KOM if PSR actual is blank
        if key == "po_issued" and not a_psr and kom_detail and kom_detail.po_date:
            a_psr = kom_detail.po_date
        if key == "kom" and not a_psr and kom_detail and kom_detail.kom_date:
            a_psr = kom_detail.kom_date

        if w_item and (w_item.get("planned_date") or w_item.get("actual_date")):
            plan_dt = w_item.get("planned_date") or p_psr
            actual_dt = w_item.get("actual_date") or a_psr
            forecast_dt = w_item.get("forecast_date", "")
            remark = w_item.get("remark", "")
            source = "WPR (Priority #1)"
            psr_ref = f_psr or p_psr
            deviation_note = f"PSR Plan/Fcst: {psr_ref}" if (psr_ref and plan_dt and psr_ref != plan_dt) else ""
        elif p_psr or f_psr or a_psr:
            plan_dt = p_psr or f_psr
            forecast_dt = f_psr
            actual_dt = a_psr
            remark = ""
            source = "PSR (Priority #2)"
            deviation_note = ""
        elif s_item:
            plan_dt = s_item.get("planned_date", "")
            forecast_dt = ""
            actual_dt = s_item.get("actual_date", "")
            remark = s_item.get("remark", "")
            source = "WPR Template"
            deviation_note = ""
        else:
            plan_dt = ""
            forecast_dt = ""
            actual_dt = ""
            remark = ""
            source = "—"
            deviation_note = ""

        # Check ROS date breach on Milestone #11 (Receipt at Worksite)
        if key == "receipt_worksite" and ros_date:
            eff_deliv = actual_dt or forecast_dt or plan_dt
            ros_diff = _date_diff_days(ros_date, eff_deliv)
            if ros_diff is not None and ros_diff > 0:
                remark = f"CRITICAL: Delivery ETA ({eff_deliv}) is +{ros_diff}d AFTER Fixed ROS Date ({ros_date})!"
            elif ros_diff is not None:
                deviation_note = f"Fixed ROS: {ros_date} ({abs(ros_diff)}d buffer)"

        eff_compare = actual_dt or forecast_dt
        variance_days = _date_diff_days(plan_dt, eff_compare)

        if key == "receipt_worksite" and "AFTER Fixed ROS" in remark:
            status = "ROS Delayed!"
            tone = "danger"
        elif actual_dt:
            status = "Completed"
            tone = "success"
        elif remark and "WAITING" in remark.upper():
            status = "Action Pending"
            tone = "warning"
        elif variance_days is not None and variance_days > 0:
            status = f"Forecast +{variance_days}d"
            tone = "warning"
        elif plan_dt and plan_dt < "2026-10-16":
            status = "Overdue / Pending"
            tone = "danger"
        elif plan_dt:
            status = "Scheduled"
            tone = "info"
        else:
            status = "Not Scheduled"
            tone = "muted"

        reconciled.append({
            "no": num,
            "key": key,
            "description": label,
            "planned_date": plan_dt,
            "forecast_date": forecast_dt,
            "actual_date": actual_dt,
            "variance_days": variance_days,
            "status": status,
            "tone": tone,
            "remark": remark,
            "deviation_note": deviation_note,
            "source": source,
        })

    return reconciled


def _build_dynamic_package_progress_and_scurve(
    pkg_code: str,
    pkg_name: str,
    discipline: str,
    milestones: List[Dict[str, Any]],
    vdrl_status: Dict[str, Any],
    client_doc_status: Dict[str, Any],
    cut_off_date: str,
    ros_date: str,
    effective_delivery_date: str,
) -> Tuple[List[Dict[str, Any]], float, float, Dict[str, Any], List[str]]:
    """Dynamically compute package-specific Milestone Progress (Element 5) and S-Curve (Element 7)
    for packages that do not yet have a separate WPR file, using their actual PSR milestones,
    VDRL submission %, and Client Review approval % so switching packages updates Element 5 & 7 dynamically!
    """
    m_by_key = {m["key"]: m for m in milestones}

    def is_done(k: str) -> bool:
        return bool(m_by_key.get(k, {}).get("actual_date"))

    def is_planned_by_cutoff(k: str) -> bool:
        p_dt = m_by_key.get(k, {}).get("planned_date", "")
        return bool(p_dt and p_dt <= cut_off_date)

    # 1. Contract Award (2%) — PSR (po_issued)
    ca_act = 100.0 if is_done("po_issued") else 0.0
    ca_plan = 100.0 if is_planned_by_cutoff("po_issued") else 0.0

    # 2. Kick-off Meeting (2%) — PSR (kom) / KOM/
    kom_act = 100.0 if is_done("kom") else 0.0
    kom_plan = 100.0 if is_planned_by_cutoff("kom") else 0.0

    # 3. Engineering and document Receive (5%) — VDRL / D2 Log / PSR
    total_vdrl = int(vdrl_status.get("total_docs", 0) or 0)
    if total_vdrl > 0:
        due_vdrl = int(vdrl_status.get("submitted", 0) or 0) + int(vdrl_status.get("overdue", 0) or 0)
        eng_plan = round(min((due_vdrl / total_vdrl) * 100.0, 100.0), 1)
        eng_act = round(min(float(vdrl_status.get("pct_submitted", 0.0)), 100.0), 1)
    elif int(client_doc_status.get("total_submittals", 0) or 0) > 0:
        eng_plan = 100.0
        eng_act = 100.0
    else:
        eng_plan = 100.0 if is_planned_by_cutoff("vd_submission") else 0.0
        eng_act = 100.0 if is_done("vd_submission") else 0.0

    # 4. VDRL Submission (2%) — VDRL / PSR (vd_submission)
    vdr_act = 100.0 if (
        is_done("vd_submission")
        or int(vdrl_status.get("submitted", 0) or 0) > 0
        or int(client_doc_status.get("total_submittals", 0) or 0) > 0
    ) else 0.0
    vdr_plan = 100.0 if (is_planned_by_cutoff("vd_submission") or vdrl_status.get("available")) else 0.0

    # 5. Approval of Key Document (5%) — D2 Log (APPR + AWC) / VDRL / PSR (approval_key_docs)
    if is_done("approval_key_docs"):
        app_act = 100.0
    elif int(client_doc_status.get("total_submittals", 0) or 0) > 0:
        returned = int(client_doc_status.get("appr", 0) or 0) + int(client_doc_status.get("awc", 0) or 0)
        app_act = round(min((returned / client_doc_status["total_submittals"]) * 100.0, 100.0), 1)
    elif total_vdrl > 0:
        v_ret = int(vdrl_status.get("appr_count", 0) or 0) + int(vdrl_status.get("awc_count", 0) or 0)
        app_act = round(min((v_ret / total_vdrl) * 100.0, 100.0), 1)
    else:
        app_act = 0.0
    app_plan = 100.0 if (
        is_planned_by_cutoff("approval_key_docs")
        or int(client_doc_status.get("total_submittals", 0) or 0) > 0
    ) else 0.0

    # 6. Material and Procurement (10% + 4 x 5% sub-components = 30%) — PSR (delivery_materials)
    mat_done = is_done("delivery_materials")
    mat_plan_due = is_planned_by_cutoff("delivery_materials")
    mat_act = 100.0 if mat_done else 0.0
    mat_plan = 100.0 if mat_plan_due else 0.0

    # Discipline-specific sub-components for Section 6.1..6.4
    sub_labels = _get_subcomponent_labels(discipline, pkg_code)

    # 7..13 Downstream Fabrication, FAT, Shipment, Worksite, Final Docs — PSR
    deliv_act = 100.0 if mat_done else 0.0
    deliv_plan = 100.0 if mat_plan_due else 0.0
    prod_act = 100.0 if is_done("start_production") else 0.0
    prod_plan = 100.0 if is_planned_by_cutoff("start_production") else 0.0
    fat_act = 100.0 if is_done("fat_hardware") else 0.0
    fat_plan = 100.0 if is_planned_by_cutoff("fat_hardware") else 0.0
    punch_act = 100.0 if is_done("punch_clearance") else 0.0
    punch_plan = 100.0 if is_planned_by_cutoff("punch_clearance") else 0.0
    ship_act = 100.0 if is_done("ready_shipment") else 0.0
    ship_plan = 100.0 if is_planned_by_cutoff("ready_shipment") else 0.0
    rec_act = 100.0 if is_done("receipt_worksite") else 0.0
    rec_plan = 100.0 if is_planned_by_cutoff("receipt_worksite") else 0.0
    fdoc_act = 100.0 if is_done("final_documentation") else 0.0
    fdoc_plan = 100.0 if is_planned_by_cutoff("final_documentation") else 0.0

    rows = [
        {"id": "1", "description": "Contract Award", "target_week": 1, "weight_pct": 2.0, "planned_pct": ca_plan, "actual_pct": ca_act, "weighted_planned_pct": round(0.02 * ca_plan, 2), "weighted_actual_pct": round(0.02 * ca_act, 2), "is_sub": False},
        {"id": "2", "description": "Kick-off Meeting", "target_week": 2, "weight_pct": 2.0, "planned_pct": kom_plan, "actual_pct": kom_act, "weighted_planned_pct": round(0.02 * kom_plan, 2), "weighted_actual_pct": round(0.02 * kom_act, 2), "is_sub": False},
        {"id": "3", "description": "Engineering and document Receive", "target_week": 3, "weight_pct": 5.0, "planned_pct": eng_plan, "actual_pct": eng_act, "weighted_planned_pct": round(0.05 * eng_plan, 2), "weighted_actual_pct": round(0.05 * eng_act, 2), "is_sub": False},
        {"id": "4", "description": "VDRL Submission", "target_week": 4, "weight_pct": 2.0, "planned_pct": vdr_plan, "actual_pct": vdr_act, "weighted_planned_pct": round(0.02 * vdr_plan, 2), "weighted_actual_pct": round(0.02 * vdr_act, 2), "is_sub": False},
        {"id": "5", "description": "Approval of Key Document", "target_week": 6, "weight_pct": 5.0, "planned_pct": app_plan, "actual_pct": app_act, "weighted_planned_pct": round(0.05 * app_plan, 2), "weighted_actual_pct": round(0.05 * app_act, 2), "is_sub": False},
        {"id": "6", "description": "Material and Procurement", "target_week": 7, "weight_pct": 10.0, "planned_pct": mat_plan, "actual_pct": mat_act, "weighted_planned_pct": round(0.10 * mat_plan, 2), "weighted_actual_pct": round(0.10 * mat_act, 2), "is_sub": False},
        {"id": "6.1", "description": sub_labels[0], "target_week": 7, "weight_pct": 5.0, "planned_pct": mat_plan, "actual_pct": mat_act, "weighted_planned_pct": round(0.05 * mat_plan, 2), "weighted_actual_pct": round(0.05 * mat_act, 2), "is_sub": True},
        {"id": "6.2", "description": sub_labels[1], "target_week": 7, "weight_pct": 5.0, "planned_pct": mat_plan, "actual_pct": mat_act, "weighted_planned_pct": round(0.05 * mat_plan, 2), "weighted_actual_pct": round(0.05 * mat_act, 2), "is_sub": True},
        {"id": "6.3", "description": sub_labels[2], "target_week": 7, "weight_pct": 5.0, "planned_pct": mat_plan, "actual_pct": mat_act, "weighted_planned_pct": round(0.05 * mat_plan, 2), "weighted_actual_pct": round(0.05 * mat_act, 2), "is_sub": True},
        {"id": "6.4", "description": sub_labels[3], "target_week": 7, "weight_pct": 5.0, "planned_pct": mat_plan, "actual_pct": mat_act, "weighted_planned_pct": round(0.05 * mat_plan, 2), "weighted_actual_pct": round(0.05 * mat_act, 2), "is_sub": True},
        {"id": "7", "description": "Delivery of Major Equipment and Material", "target_week": 22, "weight_pct": 10.0, "planned_pct": deliv_plan, "actual_pct": deliv_act, "weighted_planned_pct": round(0.10 * deliv_plan, 2), "weighted_actual_pct": round(0.10 * deliv_act, 2), "is_sub": False},
        {"id": "8", "description": "Start Production/Fabrication", "target_week": 22, "weight_pct": 10.0, "planned_pct": prod_plan, "actual_pct": prod_act, "weighted_planned_pct": round(0.10 * prod_plan, 2), "weighted_actual_pct": round(0.10 * prod_act, 2), "is_sub": False},
        {"id": "9", "description": "Factory Acceptance Test", "target_week": 29, "weight_pct": 10.0, "planned_pct": fat_plan, "actual_pct": fat_act, "weighted_planned_pct": round(0.10 * fat_plan, 2), "weighted_actual_pct": round(0.10 * fat_act, 2), "is_sub": False},
        {"id": "10", "description": "Punchlist Clerance", "target_week": 29, "weight_pct": 10.0, "planned_pct": punch_plan, "actual_pct": punch_act, "weighted_planned_pct": round(0.10 * punch_plan, 2), "weighted_actual_pct": round(0.10 * punch_act, 2), "is_sub": False},
        {"id": "11", "description": "Ready for Shipment", "target_week": 30, "weight_pct": 2.0, "planned_pct": ship_plan, "actual_pct": ship_act, "weighted_planned_pct": round(0.02 * ship_plan, 2), "weighted_actual_pct": round(0.02 * ship_act, 2), "is_sub": False},
        {"id": "12", "description": "Receipt at Worksite", "target_week": 32, "weight_pct": 2.0, "planned_pct": rec_plan, "actual_pct": rec_act, "weighted_planned_pct": round(0.02 * rec_plan, 2), "weighted_actual_pct": round(0.02 * rec_act, 2), "is_sub": False},
        {"id": "13", "description": "Final Documentation", "target_week": 40, "weight_pct": 10.0, "planned_pct": fdoc_plan, "actual_pct": fdoc_act, "weighted_planned_pct": round(0.10 * fdoc_plan, 2), "weighted_actual_pct": round(0.10 * fdoc_act, 2), "is_sub": False},
    ]

    overall_planned = round(sum(r["weighted_planned_pct"] for r in rows) + 1e-9, 1)
    overall_actual = round(sum(r["weighted_actual_pct"] for r in rows) + 1e-9, 1)

    # Determine current reporting week based on PO/KOM start date vs cut_off_date
    start_iso = (
        m_by_key.get("po_issued", {}).get("actual_date")
        or m_by_key.get("kom", {}).get("actual_date")
        or m_by_key.get("po_issued", {}).get("planned_date")
        or "2026-08-01"
    )
    elapsed_days = _date_diff_days(start_iso, cut_off_date)
    current_week = max(1, min(16, int((elapsed_days or 28) // 7) + 1))

    # Build package-specific 40-week S-Curve anchored to this package's overall_planned & overall_actual at current_week
    scurve = _build_dynamic_scurve_points(
        start_iso=start_iso,
        cut_off_date=cut_off_date,
        current_week=current_week,
        planned_now=overall_planned,
        actual_now=overall_actual,
    )

    # Dynamic This Period Activities for this package
    this_period_acts = [
        f"Package {pkg_code} ({pkg_name}): Engineering & VDRL submission tracking ({vdrl_status.get('submitted', 0)}/{vdrl_status.get('total_docs', 0)} docs submitted, {overall_actual}% cumulative progress)",
    ]
    if client_doc_status.get("total_submittals", 0) > 0:
        this_period_acts.append(
            f"Contractual D2 Review: {client_doc_status['in_review']} In Review, {client_doc_status['appr']} APPR, {client_doc_status['awc']} AWC"
        )
    if effective_delivery_date and ros_date:
        this_period_acts.append(f"Delivery ETA ({effective_delivery_date}) monitored against Fixed ROS Date ({ros_date})")

    return rows, overall_planned, overall_actual, scurve, this_period_acts


def _get_subcomponent_labels(discipline: str, pkg_code: str) -> List[str]:
    if pkg_code == "PDC":
        return [
            "Major Component #1 — Pedestal Slewing Ring & Boom Structure",
            "Major Component #2 — Diesel-Hydraulic Power Pack Engine",
            "Major Component #3 — Main & Auxiliary Winch Hoist Drums",
            "Major Component #4 — Operator Cabin & Safe Load Indicator (SLI)",
        ]
    if discipline == "PV":
        return [
            "Major Component #1 — Vessel Shell Plates & Forged Nozzles",
            "Major Component #2 — Dished Heads & Skirt / Saddle Supports",
            "Major Component #3 — Internal Mist Extractor / Cyclones / Coals",
            "Major Component #4 — Skid Piping, Valves & Level Bridles",
        ]
    if discipline == "MEC":
        return [
            "Major Component #1 — Main Pump Casing & Impeller / Rotor",
            "Major Component #2 — Explosion-Proof Electric Motor / Driver",
            "Major Component #3 — Structural Skid Baseplate & Piping Manifold",
            "Major Component #4 — Local Control Panel & Skid Instrumentation",
        ]
    if discipline == "ELE":
        return [
            "Major Component #1 — Solar PV Modules & Mounting Frames",
            "Major Component #2 — AC/DC Distribution Boards & Enclosures",
            "Major Component #3 — External Battery Charger & Battery Bank",
            "Major Component #4 — Navigation Aid Lanterns & Control Panel",
        ]
    if discipline in ("INS", "PIP") and ("V" in pkg_code):
        return [
            "Major Component #1 — Forged / Cast Valve Body & Bonnet",
            "Major Component #2 — Valve Trim (Ball / Plug / Seat / Stem)",
            "Major Component #3 — Pneumatic / Hydraulic Actuator & Spring Cylinder",
            "Major Component #4 — Solenoid, Positioner & Limit Switch Accessories",
        ]
    return [
        "Major Component #1 — Enclosure / Cabinet Assembly",
        "Major Component #2 — Circuit Breakers & Protection Devices",
        "Major Component #3 — 3rd-Party Sub-Vendor Accessories",
        "Major Component #4 — Primary Control / Instrumentation Modules",
    ]


def _build_dynamic_scurve_points(
    start_iso: str,
    cut_off_date: str,
    current_week: int,
    planned_now: float,
    actual_now: float,
) -> Dict[str, Any]:
    """Generate a smooth, package-specific 40-week S-Curve passing through (current_week, planned_now) and (current_week, actual_now)."""
    try:
        start_dt = datetime.date.fromisoformat(start_iso[:10])
    except ValueError:
        start_dt = datetime.date(2026, 8, 15)

    points: List[Dict[str, Any]] = []
    cw = max(1, min(current_week, 36))

    for w in range(1, 41):
        w_date = (start_dt + datetime.timedelta(days=(w - 1) * 7)).strftime("%Y-%m-%d")
        if w <= cw:
            ratio = w / float(cw)
            # Smooth S-ramp from 0 to planned_now and actual_now
            p_val = round(planned_now * (ratio ** 1.15), 1)
            a_val: Optional[float] = round(actual_now * (ratio ** 1.15), 1)
            if w == cw:
                p_val = round(planned_now, 1)
                a_val = round(actual_now, 1)
        else:
            # S-curve trajectory from (cw, planned_now) to (40, 100.0)
            rem_ratio = (w - cw) / float(40 - cw)
            # Logistic-style ease-in-out curve
            s_factor = (3 * (rem_ratio ** 2)) - (2 * (rem_ratio ** 3))
            p_val = round(planned_now + (100.0 - planned_now) * s_factor, 1)
            a_val = None

        points.append({
            "week": f"W{w}",
            "week_num": w,
            "date": w_date,
            "planned_pct": p_val,
            "actual_pct": a_val,
        })

    return {
        "current_week": f"W{cw}",
        "cut_off_date": cut_off_date,
        "planned_to_date_pct": round(planned_now, 1),
        "actual_to_date_pct": round(actual_now, 1),
        "variance_pct": round(actual_now - planned_now, 1),
        "points": points,
    }


def _build_project_weighted_scurve(packages_list: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Comment #2: Calculate Overall Project S-Curve from the weighted progress across all packages."""
    if not packages_list:
        return {
            "current_week": "W4",
            "cut_off_date": "2026-10-16",
            "planned_to_date_pct": 0.0,
            "actual_to_date_pct": 0.0,
            "variance_pct": 0.0,
            "points": [],
            "package_weights_breakdown": [],
        }

    total_w = sum(float(p.get("package_weight_pct", 1.0)) for p in packages_list) or 1.0
    breakdown: List[Dict[str, Any]] = []
    weighted_plan_now = 0.0
    weighted_act_now = 0.0

    for p in packages_list:
        w_pct = float(p.get("package_weight_pct", 1.0))
        norm_w = w_pct / total_w
        p_plan = float(p["overall_progress"]["planned_pct"])
        p_act = float(p["overall_progress"]["actual_pct"])
        w_plan_contrib = round(norm_w * p_plan, 2)
        w_act_contrib = round(norm_w * p_act, 2)
        weighted_plan_now += norm_w * p_plan
        weighted_act_now += norm_w * p_act
        breakdown.append({
            "package_code": p["package_code"],
            "package_name": p["package_name"],
            "discipline": p["discipline"],
            "package_weight_pct": round(norm_w * 100.0, 2),
            "pkg_planned_pct": p_plan,
            "pkg_actual_pct": p_act,
            "weighted_planned_pct": w_plan_contrib,
            "weighted_actual_pct": w_act_contrib,
            "variance_pct": round(p_act - p_plan, 1),
        })

    # Aggregate 40-week curve across all packages
    agg_points: List[Dict[str, Any]] = []
    cut_off_week_num = 6
    for w_idx in range(40):
        w_num = w_idx + 1
        w_plan_sum = 0.0
        w_act_sum = 0.0
        for p in packages_list:
            w_pct = float(p.get("package_weight_pct", 1.0)) / total_w
            sc_raw = p.get("scurve") or {}
            pts = sc_raw.get("points", []) if isinstance(sc_raw, dict) else (sc_raw if isinstance(sc_raw, list) else [])
            if w_idx < len(pts):
                pt = pts[w_idx]
                w_plan_sum += w_pct * float(pt.get("planned_pct") or pt.get("plan_cum_pct") or 0.0)
                if pt.get("actual_pct") is not None or pt.get("actual_cum_pct") is not None:
                    act_val = pt.get("actual_pct") if pt.get("actual_pct") is not None else pt.get("actual_cum_pct")
                    w_act_sum += w_pct * float(act_val or 0.0)
                elif w_num <= cut_off_week_num:
                    w_act_sum += w_pct * float(p["overall_progress"]["actual_pct"])

        agg_points.append({
            "week": f"W{w_num}",
            "week_num": w_num,
            "date": _add_days("2026-09-04", (w_num - 1) * 7),
            "planned_pct": round(w_plan_sum, 1),
            "actual_pct": round(w_act_sum, 1) if w_num <= cut_off_week_num else None,
        })

    # Smoothly anchor cumulative S-curve through the exact weighted cut-off (no spike/drop at W6->W7)
    if len(agg_points) >= cut_off_week_num:
        cut_idx = cut_off_week_num - 1
        raw_cut_plan = agg_points[cut_idx]["planned_pct"] or 1.0
        raw_cut_act = agg_points[cut_idx]["actual_pct"] or 0.0
        target_plan = round(weighted_plan_now, 1)
        target_act = round(weighted_act_now, 1)
        end_plan = max(agg_points[-1]["planned_pct"] or 100.0, target_plan, 100.0)
        raw_span_after = max(1.0, (agg_points[-1]["planned_pct"] or 100.0) - raw_cut_plan)

        for idx, pt in enumerate(agg_points):
            if idx < cut_idx:
                pt["planned_pct"] = round(target_plan * ((pt["planned_pct"] or 0.0) / raw_cut_plan), 1)
                if raw_cut_act > 0:
                    pt["actual_pct"] = round(target_act * ((pt["actual_pct"] or 0.0) / raw_cut_act), 1)
                else:
                    pt["actual_pct"] = round(target_act * ((idx + 1) / cut_off_week_num), 1)
            elif idx == cut_idx:
                pt["planned_pct"] = target_plan
                pt["actual_pct"] = target_act
            else:
                frac = max(0.0, min(1.0, ((pt["planned_pct"] or 0.0) - raw_cut_plan) / raw_span_after))
                pt["planned_pct"] = round(target_plan + (end_plan - target_plan) * frac, 1)

        # Guarantee monotonic cumulative values
        for idx in range(1, len(agg_points)):
            if agg_points[idx]["planned_pct"] < agg_points[idx - 1]["planned_pct"]:
                agg_points[idx]["planned_pct"] = agg_points[idx - 1]["planned_pct"]
            if agg_points[idx]["actual_pct"] is not None and agg_points[idx - 1]["actual_pct"] is not None:
                if agg_points[idx]["actual_pct"] < agg_points[idx - 1]["actual_pct"]:
                    agg_points[idx]["actual_pct"] = agg_points[idx - 1]["actual_pct"]

    return {
        "current_week": f"W{cut_off_week_num}",
        "cut_off_date": "2026-10-16",
        "planned_to_date_pct": round(weighted_plan_now, 1),
        "actual_to_date_pct": round(weighted_act_now, 1),
        "variance_pct": round(weighted_act_now - weighted_plan_now, 1),
        "points": agg_points,
        "package_weights_breakdown": breakdown,
    }


def _build_package_highlights(
    pkg_code: str,
    pkg_name: str,
    milestones: List[Dict[str, Any]],
    cut_off_date: str,
    effective_delivery_date: str,
    delivery_source: str,
    ros_date: str,
    master_entry: Optional[Dict[str, Any]] = None,
    vdrl_status: Optional[Dict[str, Any]] = None,
    client_doc_status: Optional[Dict[str, Any]] = None,
    has_live_wpr: bool = False,
    live_wpr: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Comment #4.1, #6 & Master Schedule Priority:
    Compute:
    1) Current status based on key milestone (latest completed & active stage)
    2) Number of remaining days to next key milestone
    3) How many days to Delivery (ETA)
    4) Priority #1 Project Master Schedule (MM-ASK-1A-GEN01-PJM-SCD-0001.pdf) + Immutable PSR ROS Date compliance
    5) 4-Tier Health Classification (ALREADY_DELAYED, HIGH_RISK, LOW_RISK, HEALTHY) + Smart Issue & Timely Mitigation
    """
    completed_ms = [m for m in milestones if m.get("actual_date")]
    pending_ms = [m for m in milestones if not m.get("actual_date") and (m.get("planned_date") or m.get("forecast_date"))]

    latest_completed = completed_ms[-1] if completed_ms else None
    next_ms = pending_ms[0] if pending_ms else (milestones[-1] if milestones else None)

    next_target_dt = ""
    rem_days_next_ms: Optional[int] = None
    next_ms_badge = "—"
    next_ms_tone = "info"

    if next_ms:
        next_target_dt = next_ms.get("forecast_date") or next_ms.get("planned_date") or ""
        rem_days_next_ms = _date_diff_days(cut_off_date, next_target_dt)
        if rem_days_next_ms is not None:
            if rem_days_next_ms < 0:
                next_ms_badge = f"Overdue by {abs(rem_days_next_ms)}d"
                next_ms_tone = "danger"
            elif rem_days_next_ms <= 14:
                next_ms_badge = f"{rem_days_next_ms} Days Remaining (Due Soon)"
                next_ms_tone = "warning"
            else:
                next_ms_badge = f"{rem_days_next_ms} Days Remaining"
                next_ms_tone = "info"

    days_to_delivery = _date_diff_days(cut_off_date, effective_delivery_date)

    # Master Plan (MM-ASK-1A-GEN01-PJM-SCD-0001.pdf) Informational Metadata & Yard Erection
    m_entry = master_entry or {}
    master_eta_date = m_entry.get("master_eta_date") or ""
    yard_erection_date = m_entry.get("yard_erection_date") or ""
    yard_erection_activity = m_entry.get("yard_erection_activity") or "Yard Installation Window"
    master_activity_id = m_entry.get("activity_id") or ""
    master_key_doc_dt = m_entry.get("master_key_doc_approval_date") or ""
    is_critical_path = bool(m_entry.get("is_critical_path", False))

    # User Direct Instruction: Reference file is MM-ASK-1A-GEN01-PJM-SCD-0001 (Priority #1) together with Procurement Plan (Priority #2)
    governing_ros_date = master_eta_date or ros_date
    governing_ros_source = (
        "Master Schedule (MM-ASK-1A-GEN01-PJM-SCD-0001)"
        if master_eta_date
        else "Procurement Plan"
    )

    ros_delay_days = (
        _date_diff_days(governing_ros_date, effective_delivery_date)
        if (governing_ros_date and effective_delivery_date)
        else None
    )
    is_ros_delayed = bool(ros_delay_days is not None and ros_delay_days > 0)
    effective_delay_days = ros_delay_days if is_ros_delayed else 0
    ros_margin_days = (-ros_delay_days) if ros_delay_days is not None else None
    governing_target_date = governing_ros_date

    # Master Plan & Procurement Plan separate tracking
    master_delay_days = (
        _date_diff_days(master_eta_date, effective_delivery_date)
        if (master_eta_date and effective_delivery_date)
        else None
    )
    is_master_delayed = bool(master_delay_days is not None and master_delay_days > 0)
    master_margin_days = (-master_delay_days) if master_delay_days is not None else None

    psr_ros_delay_days = (
        _date_diff_days(ros_date, effective_delivery_date)
        if (ros_date and effective_delivery_date)
        else None
    )
    is_psr_ros_delayed = bool(psr_ros_delay_days is not None and psr_ros_delay_days > 0)
    is_erection_conflict = bool(
        yard_erection_date and effective_delivery_date and
        (_date_diff_days(yard_erection_date, effective_delivery_date) or 0) > 0
    )
    erection_delay_days = _date_diff_days(yard_erection_date, effective_delivery_date) or 0

    vdrl_ov = int((vdrl_status or {}).get("overdue", 0) or 0)
    key_ov = int((vdrl_status or {}).get("key_docs_overdue", 0) or 0)
    prod_ov = int((vdrl_status or {}).get("production_impact_overdue", 0) or 0)
    vdrl_sub = int((vdrl_status or {}).get("submitted", 0) or 0)
    vdrl_tot = int((vdrl_status or {}).get("total_docs", 0) or 0)
    d2_ov = int((client_doc_status or {}).get("overdue_reviews", 0) or 0)
    d2_awc = int((client_doc_status or {}).get("awc", 0) or 0)
    wpr_mat_holds = len((live_wpr or {}).get("supplier_material_status", [])) if has_live_wpr else 0

    # 4-Tier Health Classification (Governed by Master Schedule ROS & Procurement Plan ROS):
    # 1. ALREADY_DELAYED: Vendor ETA > Governing ROS Date
    # 2. HIGH_RISK: Critical Path with tight buffer (<=14d) OR 0-7d buffer vs ROS OR WPR sub-vendor hold OR >=4 Key VDRL overdue OR 0 VDRL submitted past Key Doc gate
    # 3. LOW_RISK: 8-14d buffer vs ROS OR 1-3 overdue VDRL docs OR overdue Client D2 reviews / AWC pending closure
    # 4. HEALTHY: >14d buffer ahead of ROS Date AND 0 overdue VDRL docs
    if is_ros_delayed:
        health_tier = "ALREADY_DELAYED"
        health_status_label = f"🔴 Already Delayed (+{effective_delay_days}d)"
        health_tone = "danger"
        tier_badge_text = f"+{effective_delay_days}d Late vs {governing_ros_source}"
    elif (
        (ros_margin_days is not None and 0 <= ros_margin_days <= 7 and (vdrl_ov > 0 or vdrl_tot > 0))
        or wpr_mat_holds > 0
        or key_ov >= 4
        or prod_ov >= 3
        or (is_critical_path and ros_margin_days is not None and ros_margin_days <= 14)
        or (vdrl_tot > 0 and vdrl_sub == 0 and ros_margin_days is not None and ros_margin_days <= 14)
    ):
        health_tier = "HIGH_RISK"
        health_status_label = "🟠 High Risk Delay"
        health_tone = "warning"
        if is_critical_path and ros_margin_days is not None:
            tier_badge_text = f"TF=0 Critical Path ({ros_margin_days}d Buf.)"
        elif ros_margin_days is not None and ros_margin_days <= 7:
            tier_badge_text = f"{ros_margin_days}d Buffer vs {governing_ros_source}"
        elif wpr_mat_holds > 0:
            tier_badge_text = f"Att.4 Mat Hold ({key_ov} Key VDRL Ov.)"
        elif vdrl_tot > 0 and vdrl_sub == 0:
            tier_badge_text = f"0/{vdrl_tot} VDRL Sub. (Key Doc Gate)"
        else:
            tier_badge_text = f"{key_ov} Key VDRL Overdue"
    elif (
        vdrl_ov > 0
        or d2_ov > 0
        or (ros_margin_days is not None and 0 <= ros_margin_days <= 14 and vdrl_tot > 0)
    ):
        health_tier = "LOW_RISK"
        health_status_label = "🟡 Low Risk"
        health_tone = "info"
        if vdrl_ov > 0:
            tier_badge_text = f"{ros_margin_days if ros_margin_days is not None else '—'}d Buf. • {vdrl_ov} VDRL Ov."
        elif ros_margin_days is not None:
            tier_badge_text = f"{ros_margin_days}d Buffer (Watch)"
        else:
            tier_badge_text = f"{d2_ov} D2 Review Overdue"
    else:
        health_tier = "HEALTHY"
        health_status_label = "🟢 Very Good / Healthy"
        health_tone = "success"
        tier_badge_text = (
            f"✓ {ros_margin_days}d Ahead of {governing_ros_source}"
            if ros_margin_days is not None and ros_margin_days > 0
            else f"✓ On Track vs {governing_ros_source}"
        )

    # Status label for Master Schedule & Procurement Plan ROS Card (Card 6)
    if ros_margin_days is None and not is_ros_delayed:
        ros_status_label = f"Governing ROS / ETA Pending ({governing_ros_source})"
        ros_tone = "muted"
    elif is_ros_delayed:
        ros_status_label = f"🚨 DELAYED: ETA ({effective_delivery_date}) is +{effective_delay_days}d AFTER {governing_ros_source} ({governing_ros_date})"
        ros_tone = "danger"
    elif ros_margin_days == 0:
        ros_status_label = f"⚠️ Zero Buffer: ETA Equals {governing_ros_source} ({governing_ros_date})"
        ros_tone = "warning"
    elif ros_margin_days is not None and ros_margin_days <= 7:
        ros_status_label = f"⚠️ Tight Buffer: {ros_margin_days}d Ahead of {governing_ros_source} ({governing_ros_date})"
        ros_tone = "warning"
    else:
        ros_status_label = f"✓ On Time: {ros_margin_days}d Buffer Ahead of {governing_ros_source} ({governing_ros_date})"
        ros_tone = "success"

    # Smart Issue Identification & On-Time Mitigation Measure
    smart_issue, smart_mitigation, mitigation_deadline, mitigation_owner = _build_smart_issue_and_mitigation(
        pkg_code=pkg_code,
        pkg_name=pkg_name,
        health_tier=health_tier,
        effective_delivery_date=effective_delivery_date,
        master_eta_date=master_eta_date,
        master_activity_id=master_activity_id,
        yard_erection_date=yard_erection_date,
        yard_erection_activity=yard_erection_activity,
        ros_date=ros_date,
        is_master_delayed=is_master_delayed,
        master_delay_days=master_delay_days or 0,
        is_psr_ros_delayed=is_psr_ros_delayed,
        psr_ros_delay_days=psr_ros_delay_days or 0,
        is_erection_conflict=is_erection_conflict,
        erection_delay_days=erection_delay_days or 0,
        ros_margin_days=ros_margin_days,
        is_critical_path=is_critical_path,
        master_key_doc_dt=master_key_doc_dt,
        vdrl_tot=vdrl_tot,
        vdrl_sub=vdrl_sub,
        vdrl_ov=vdrl_ov,
        key_ov=key_ov,
        prod_ov=prod_ov,
        d2_ov=d2_ov,
        d2_awc=d2_awc,
        wpr_mat_holds=wpr_mat_holds,
        cut_off_date=cut_off_date,
    )

    current_stage_summary = (
        f"Completed #{latest_completed['no']} {latest_completed['description']} ({latest_completed['actual_date']})"
        if latest_completed
        else "Awaiting Milestone #1 (Contract Award / KOM)"
    )
    active_stage_summary = (
        f"Active: #{next_ms['no']} {next_ms['description']}"
        if next_ms
        else "All Key Milestones Completed"
    )

    return {
        "reference_date": cut_off_date,
        "completed_milestones_count": len(completed_ms),
        "total_milestones_count": len(milestones),
        "current_stage_summary": current_stage_summary,
        "active_stage_summary": active_stage_summary,
        "latest_completed_milestone": latest_completed,
        "next_milestone_no": next_ms["no"] if next_ms else None,
        "next_milestone_name": next_ms["description"] if next_ms else "—",
        "next_milestone_date": next_target_dt,
        "remaining_days_to_next_milestone": rem_days_next_ms,
        "next_milestone_badge": next_ms_badge,
        "next_milestone_tone": next_ms_tone,
        "effective_delivery_date": effective_delivery_date,
        "delivery_source": delivery_source,
        "days_to_delivery": days_to_delivery,
        "ros_date": governing_target_date or ros_date,
        "governing_ros_date": governing_target_date or ros_date,
        "governing_ros_source": governing_ros_source,
        "psr_ros_date": ros_date,
        "master_eta_date": master_eta_date,
        "master_activity_id": master_activity_id,
        "master_key_doc_approval_date": master_key_doc_dt,
        "yard_erection_date": yard_erection_date,
        "yard_erection_activity": yard_erection_activity,
        "is_critical_path": is_critical_path,
        "is_master_delayed": is_master_delayed,
        "master_delay_days": master_delay_days if is_master_delayed else 0,
        "master_margin_days": master_margin_days,
        "is_psr_ros_delayed": is_psr_ros_delayed,
        "psr_ros_delay_days": psr_ros_delay_days if is_psr_ros_delayed else 0,
        "is_erection_conflict": is_erection_conflict,
        "erection_delay_days": erection_delay_days if is_erection_conflict else 0,
        "is_ros_delayed": is_ros_delayed,
        "ros_delay_days": effective_delay_days,
        "effective_delay_days": effective_delay_days,
        "ros_margin_days": ros_margin_days,
        "ros_status_label": ros_status_label,
        "ros_tone": ros_tone,
        "health_tier": health_tier,
        "health_status_label": health_status_label,
        "health_tone": health_tone,
        "tier_badge_text": tier_badge_text,
        "smart_issue": smart_issue,
        "smart_mitigation": smart_mitigation,
        "mitigation_deadline": mitigation_deadline,
        "mitigation_owner": mitigation_owner,
    }


def _build_smart_issue_and_mitigation(
    pkg_code: str,
    pkg_name: str,
    health_tier: str,
    effective_delivery_date: str,
    master_eta_date: str,
    master_activity_id: str,
    yard_erection_date: str,
    yard_erection_activity: str,
    ros_date: str,
    is_master_delayed: bool,
    master_delay_days: int,
    is_psr_ros_delayed: bool,
    psr_ros_delay_days: int,
    is_erection_conflict: bool,
    erection_delay_days: int,
    ros_margin_days: Optional[int],
    is_critical_path: bool,
    master_key_doc_dt: str,
    vdrl_tot: int,
    vdrl_sub: int,
    vdrl_ov: int,
    key_ov: int,
    prod_ov: int,
    d2_ov: int,
    d2_awc: int,
    wpr_mat_holds: int,
    cut_off_date: str,
) -> Tuple[str, str, str, str]:
    """Generate concise, smart root-cause issue diagnosis and timely mitigation measure with target deadline & owner."""
    cp_tag = " [Master Plan TF=0 Critical Path]" if is_critical_path else ""

    ref_name = f"Master Schedule ROS ({master_eta_date})" if master_eta_date else f"Procurement Plan ROS ({ros_date})"
    gov_target = master_eta_date or ros_date
    gov_delay = (
        _date_diff_days(gov_target, effective_delivery_date)
        if (gov_target and effective_delivery_date)
        else 0
    ) or 0

    if health_tier == "ALREADY_DELAYED":
        doc_note = f"; compounded by {vdrl_ov} overdue VDRL doc(s) ({key_ov} Key Docs)" if vdrl_ov > 0 else ""
        issue = (
            f"Vendor Delivery ETA ({effective_delivery_date}) is +{gov_delay}d AFTER {ref_name}{doc_note}{cp_tag}."
        )
        mitigation = (
            f"1) Issue formal Schedule Recovery Notice to Vendor to compress manufacturing lead-time by ≥{gov_delay + 7}d; "
            f"2) Fast-track Code-1/2 approval of {key_ov or 'pending'} Key Docs (GAD/Datasheet/BOM) within 5 days to release sub-orders; "
            f"3) Prioritize mill/stock allocation & express shipping to pull Receipt at Worksite ahead of {ref_name}."
        )
        deadline = _add_days(cut_off_date, 7)
        owner = "PMT Expediting & Package Lead / Vendor PM"
        return issue, mitigation, deadline, owner

    if health_tier == "HIGH_RISK":
        if pkg_code == "PLC" or wpr_mat_holds > 0:
            issue = (
                f"Only {ros_margin_days}d buffer vs {ref_name}; Live WPR Att. 4 shows Enclosure & Circuit Breaker POs on HOLD "
                f"pending GAD/Datasheet approval, with {key_ov} Key VDRL docs ({prod_ov} Production-Impact) overdue."
            )
            mitigation = (
                f"1) Convene urgent technical clarification with Rockwell & Client to close GAD/Datasheet comments and release Enclosure & Breaker sub-POs; "
                f"2) Clear {key_ov} overdue Key VDRL docs by {_add_days(cut_off_date, 7)} to protect Delivery ETA ({effective_delivery_date}) before {ref_name}."
            )
            deadline = _add_days(cut_off_date, 7)
            owner = "INS Package Engineer & Rockwell PM"
        elif vdrl_tot > 0 and vdrl_sub == 0:
            issue = (
                f"0/{vdrl_tot} VDRL documents submitted past Key Doc Approval gate, "
                f"leaving only {ros_margin_days if ros_margin_days is not None else 0}d buffer before {ref_name} & Yard Erection ({yard_erection_date or 'Plan'})."
            )
            mitigation = (
                f"1) Escalate with Vendor Management for immediate submission of Batch-1 Key VDRL documents (GAD, Datasheet, Schematic, ITP); "
                f"2) Pre-book sub-vendor long-lead material slots to prevent slippage past {ref_name}."
            )
            deadline = _add_days(cut_off_date, 5)
            owner = "Document Control & Package Lead"
        elif is_critical_path:
            issue = (
                f"Master Schedule Critical Path ({master_activity_id}, TF=0): ETA ({effective_delivery_date}) has only "
                f"{ros_margin_days if ros_margin_days is not None else 0}d buffer vs {ref_name} before {yard_erection_activity} ({yard_erection_date})"
                + (f", with {vdrl_ov} VDRL doc(s) overdue ({key_ov} Key)." if vdrl_ov > 0 else ".")
            )
            mitigation = (
                f"1) Enforce weekly sub-vendor material expediting (WPR Att. 4) & zero-slippage gate on Key Document Approval; "
                f"2) Clear {key_ov or vdrl_ov or 'all'} pending engineering submittals within 7 days so shop production starts on schedule."
            )
            deadline = _add_days(cut_off_date, 7)
            owner = "PMT Critical Path Coordinator & Package Lead"
        elif key_ov >= 4:
            issue = (
                f"High backlog of {key_ov} overdue Key VDRL documents ({vdrl_ov} total overdue, {prod_ov} Production-Impact) "
                f"threatening Phase 2B sub-vendor material ordering and {ref_name} delivery."
            )
            mitigation = (
                f"1) Hold dedicated engineering workshop with Vendor to clear {key_ov} overdue Key VDRL deliverables and {d2_awc} AWC client comments; "
                f"2) Authorize long-lead raw material mill order upon Datasheet/BOM approval."
            )
            deadline = _add_days(cut_off_date, 10)
            owner = "Discipline Lead Engineer & Expeditor"
        else:
            issue = (
                f"Tight {ros_margin_days}d delivery buffer between Vendor ETA ({effective_delivery_date}) and {ref_name}."
            )
            mitigation = (
                f"1) Lock in sub-supplier raw material delivery dates in WPR Attachment 4; "
                f"2) Combine shop ITP witness points with FAT to gain 5–7 days schedule cushion before {ref_name}."
            )
            deadline = _add_days(cut_off_date, 10)
            owner = "Package Expeditor & QA/QC Lead"
        return issue, mitigation, deadline, owner

    if health_tier == "LOW_RISK":
        issue = (
            f"Delivery ETA ({effective_delivery_date}) is {ros_margin_days if ros_margin_days is not None else '—'}d ahead of {ref_name}, "
            f"but {vdrl_ov} VDRL doc(s) ({key_ov} Key) and {d2_ov} Client review(s) ({d2_awc} AWC) require timely closure."
        )
        mitigation = (
            f"1) Expedite remaining {vdrl_ov} overdue VDRL submittals and close {d2_awc} AWC comment sheets in next bi-weekly cycle; "
            f"2) Maintain current {ros_margin_days if ros_margin_days is not None else 14}d buffer ahead of {ref_name}."
        )
        deadline = _add_days(cut_off_date, 14)
        owner = "Package Engineer"
        return issue, mitigation, deadline, owner

    # HEALTHY
    issue = (
        f"On track: Vendor ETA ({effective_delivery_date or 'Scheduled'}) maintains a healthy "
        f"{ros_margin_days if ros_margin_days is not None else 'compliant'}d buffer ahead of {ref_name} with 0 overdue VDRL docs."
    )
    mitigation = (
        f"Continue routine weekly/bi-weekly WPR & VDRL monitoring to preserve schedule buffer ahead of ROS Date ({ros_date}) & Yard Erection ({yard_erection_date or 'Plan'})."
    )
    deadline = "Routine Weekly"
    owner = "Package Engineer"
    return issue, mitigation, deadline, owner


def _build_attachments_3_and_4_summary(
    pkg_code: str,
    pkg_name: str,
    discipline: str,
    milestones: List[Dict[str, Any]],
    vdrl_status: Dict[str, Any],
    client_doc_status: Dict[str, Any],
    live_wpr: Optional[Dict[str, Any]],
    package_highlights: Dict[str, Any],
    cut_off_date: str,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Comment #5: Analyze WPR Attachment 3 (Overall Fabrication Detail Plan Gantt)
    and Attachment 4 (Expediting Report for Supplier's Material Status) to provide
    an executive situation analysis, critical path summary, and sub-vendor material expediting table for each package.
    """
    m_by_key = {m["key"]: m for m in milestones}

    def m_dt(k: str, fallback: str) -> str:
        m = m_by_key.get(k, {})
        return m.get("actual_date") or m.get("forecast_date") or m.get("planned_date") or fallback

    po_dt = m_dt("po_issued", "2026-08-15")
    kom_dt = m_dt("kom", "2026-09-02")
    vdr_dt = m_dt("vd_submission", "2026-09-24")
    app_dt = m_dt("approval_key_docs", "2026-11-12")
    mat_dt = m_dt("delivery_materials", "2027-01-20")
    prod_dt = m_dt("start_production", "2027-01-25")
    fat_dt = m_dt("fat_hardware", "2027-03-03")
    ship_dt = m_dt("ready_shipment", "2027-03-23")
    rec_dt = m_dt("receipt_worksite", "2027-04-08")
    fin_dt = m_dt("final_documentation", "2027-06-30")

    # Build Phase-level Gantt Summary + Task rows (from live WPR if present, else dynamic package schedule)
    raw_tasks = (live_wpr or {}).get("gantt_tasks", [])
    phases_summary = [
        {
            "phase_id": "Phase 1",
            "phase_name": "Project Initiation, PO & Kick-Off Meeting",
            "start_date": po_dt,
            "finish_date": vdr_dt,
            "duration_days": max(1, _date_diff_days(po_dt, vdr_dt) or 14),
            "status": "Completed" if m_by_key.get("kom", {}).get("actual_date") else "In Progress",
            "tone": "success" if m_by_key.get("kom", {}).get("actual_date") else "warning",
            "critical_note": "KOM & Initial VDRL baseline established" if m_by_key.get("kom", {}).get("actual_date") else "Waiting PO / KOM finalization",
        },
        {
            "phase_id": "Phase 2A",
            "phase_name": "Detailed Engineering, GAD & Key Document Approval Cycle",
            "start_date": kom_dt,
            "finish_date": app_dt,
            "duration_days": max(14, _date_diff_days(kom_dt, app_dt) or 40),
            "status": "Completed" if m_by_key.get("approval_key_docs", {}).get("actual_date") else ("Delayed / Overdue Docs" if vdrl_status.get("key_docs_overdue", 0) > 0 else "In Progress (Active Gate)"),
            "tone": "success" if m_by_key.get("approval_key_docs", {}).get("actual_date") else ("danger" if vdrl_status.get("key_docs_overdue", 0) > 0 else "info"),
            "critical_note": f"Gate for Sub-vendor Material PO placement ({vdrl_status.get('key_docs_overdue', 0)} Key VDRL overdue; {client_doc_status.get('in_review', 0)} in Client review)",
        },
        {
            "phase_id": "Phase 2B",
            "phase_name": "Sub-Vendor Materials & Long-Lead Component Procurement",
            "start_date": _add_days(app_dt, -10),
            "finish_date": mat_dt,
            "duration_days": max(30, _date_diff_days(_add_days(app_dt, -10), mat_dt) or 65),
            "status": "Completed" if m_by_key.get("delivery_materials", {}).get("actual_date") else "Hold on Key Doc Approval",
            "tone": "success" if m_by_key.get("delivery_materials", {}).get("actual_date") else "warning",
            "critical_note": "Sub-orders gated by GAD, BOM, & Datasheet Code-1/2 approval",
        },
        {
            "phase_id": "Phase 2C",
            "phase_name": "Factory Fabrication, Assembly & Internal Shop Testing",
            "start_date": prod_dt,
            "finish_date": fat_dt,
            "duration_days": max(20, _date_diff_days(prod_dt, fat_dt) or 40),
            "status": "Scheduled",
            "tone": "info",
            "critical_note": f"Shop assembly & ITP hold-point inspections prior to FAT ({fat_dt})",
        },
        {
            "phase_id": "Phase 2D",
            "phase_name": "FAT, Punch Clearance, Packing & Delivery to Worksite",
            "start_date": fat_dt,
            "finish_date": rec_dt,
            "duration_days": max(14, _date_diff_days(fat_dt, rec_dt) or 30),
            "status": "ROS Breach Risk!" if package_highlights["is_ros_delayed"] else "Scheduled",
            "tone": "danger" if package_highlights["is_ros_delayed"] else "info",
            "critical_note": (
                f"CRITICAL: Projected Worksite Delivery ({rec_dt}) exceeds Fixed ROS ({package_highlights['ros_date']}) by +{package_highlights['ros_delay_days']}d!"
                if package_highlights["is_ros_delayed"]
                else f"Target Delivery ({rec_dt}) meets Fixed ROS ({package_highlights['ros_date'] or '—'})"
            ),
        },
    ]

    if not raw_tasks:
        raw_tasks = [
            {"id": i + 1, "name": p["phase_name"], "is_summary": False, "duration": f"{p['duration_days']} days", "start_date": p["start_date"], "finish_date": p["finish_date"], "status": p["status"], "critical_note": p["critical_note"]}
            for i, p in enumerate(phases_summary)
        ]

    # Situation Analysis narrative for Attachment 3
    bottleneck_summary = (
        f"Critical Path Analysis ({pkg_code}): Active bottleneck is Phase 2A (Engineering & Key Document Approval target {app_dt}), "
        f"which directly gates Phase 2B Sub-Vendor Material Ordering ({mat_dt}). "
    )
    if package_highlights.get("has_milestone_discrepancy"):
        bottleneck_summary += f"{package_highlights.get('milestone_discrepancy_note', '')} "
    if package_highlights["is_ros_delayed"]:
        bottleneck_summary += (
            f"🚨 SCHEDULE RECOVERY REQUIRED: Current Delivery ETA ({package_highlights['effective_delivery_date']}) is "
            f"+{package_highlights['ros_delay_days']} days AFTER the immutable PSR ROS Date ({package_highlights['ros_date']})."
        )
    else:
        margin = package_highlights.get("ros_margin_days")
        bottleneck_summary += (
            f"Delivery ETA ({package_highlights['effective_delivery_date'] or '—'}) currently maintains a "
            f"{margin if margin is not None else '—'}-day buffer ahead of Fixed ROS Date ({package_highlights['ros_date'] or '—'})."
        )

    attachment3_summary = {
        "package_code": pkg_code,
        "total_phases": len(phases_summary),
        "total_tasks": len(raw_tasks),
        "critical_gate": f"Key Document Approval ({app_dt}) → Sub-Vendor Material Release",
        "situation_summary": bottleneck_summary,
        "phases": phases_summary,
        "tasks": raw_tasks,
    }

    # Attachment 4: Expediting Report for Supplier's Material Status Summary
    live_sub_orders = (live_wpr or {}).get("sub_orders", [])
    if live_sub_orders:
        expediting_items = []
        for idx_s, so in enumerate(live_sub_orders, 1):
            po_st = str(so.get("po_status") or "Available").strip()
            est_dt = str(so.get("est_delivery") or so.get("req_delivery") or mat_dt)
            po_dt_s = str(so.get("po_date") or po_dt)
            lt_d = max(14, _date_diff_days(po_dt_s, est_dt) or 45)
            expediting_items.append({
                "item_no": f"6.{idx_s}",
                "component": f"{so.get('sub_component', '')}" + (f" ({so.get('material_description')})" if so.get("material_description") else ""),
                "sub_vendor_origin": str(so.get("sub_vendor") or "Approved Sub-Supplier"),
                "prerequisite_gate": str(so.get("remarks") or "Sub-Vendor PO Issued & Expediting"),
                "po_status": f"{po_st} (PO: {po_dt_s})" if po_dt_s else po_st,
                "lead_time_days": lt_d,
                "planned_order_date": po_dt_s,
                "planned_eta": est_dt,
                "impact_on_production": "Critical Long-Lead Pump/Instrument" if idx_s in (1, 9) else "Skid Assembly Sub-Component",
                "risk_level": "Expedite (End of Dec Dispatch)" if idx_s == 1 else "Ordered / In Progress",
                "tone": "warning" if idx_s in (1, 9) else "success",
            })
        attachment4_summary = {
            "package_code": pkg_code,
            "total_components": len(expediting_items),
            "components_on_hold": sum(1 for x in expediting_items if "Hold" in x["po_status"]),
            "target_material_delivery_date": mat_dt,
            "expediting_narrative": (
                f"Attachment 4 Material Expediting Summary ({pkg_code} Live WPR): {len(expediting_items)} sub-supplier POs tracked "
                f"(including Milton Roy Gas Driven Metering Pump, Kenco Calibration Pot, Versa Solenoid Valve, Ashcroft Gauges). "
                f"Pump manufacturer confirmed dispatch around end of December 2026; MOG expediting Flow Switch & Pump arrival."
            ),
            "items": expediting_items,
        }
        return attachment3_summary, attachment4_summary

    sub_labels = _get_subcomponent_labels(discipline, pkg_code)
    mat_holds = (live_wpr or {}).get("supplier_material_status", [])
    hold_1 = mat_holds[0] if len(mat_holds) > 0 else "Waiting GAD / Datasheet Approval"
    hold_2 = mat_holds[1] if len(mat_holds) > 1 else "Waiting BOM / Material Spec Approval"
    hold_3 = mat_holds[2] if len(mat_holds) > 2 else ("Waiting Official PO Issue" if not m_by_key.get("po_issued", {}).get("actual_date") else "Sub-vendor Technical Clarification")

    expediting_items = [
        {
            "item_no": "6.1",
            "component": sub_labels[0].replace("Major Component #1 — ", ""),
            "sub_vendor_origin": "Approved Sub-Vendor / Mill",
            "prerequisite_gate": hold_1,
            "po_status": "Ordered / In Progress" if m_by_key.get("delivery_materials", {}).get("actual_date") else "Hold — Pending Key Doc Approval",
            "lead_time_days": 60,
            "planned_order_date": _add_days(app_dt, -7),
            "planned_eta": _add_days(mat_dt, -5),
            "impact_on_production": "Critical Path (Blocks Shop Assembly Start)",
            "risk_level": "High Risk" if vdrl_status.get("production_impact_overdue", 0) > 0 else "Medium / Watch",
            "tone": "danger" if vdrl_status.get("production_impact_overdue", 0) > 0 else "warning",
        },
        {
            "item_no": "6.2",
            "component": sub_labels[1].replace("Major Component #2 — ", ""),
            "sub_vendor_origin": "International Specialized Sub-Vendor",
            "prerequisite_gate": hold_2,
            "po_status": "Hold — Pending BOM / Calc Approval",
            "lead_time_days": 60,
            "planned_order_date": app_dt,
            "planned_eta": mat_dt,
            "impact_on_production": "High Impact (Long-Lead Imported Assembly)",
            "risk_level": "High Risk" if vdrl_status.get("fabrication_impact_overdue", 0) > 0 else "Medium / Watch",
            "tone": "danger" if vdrl_status.get("fabrication_impact_overdue", 0) > 0 else "warning",
        },
        {
            "item_no": "6.3",
            "component": sub_labels[2].replace("Major Component #3 — ", ""),
            "sub_vendor_origin": "Tier-1 Sub-Supplier (Air/Sea Freight)",
            "prerequisite_gate": hold_3,
            "po_status": "Sub-order Preparation",
            "lead_time_days": 45,
            "planned_order_date": _add_days(app_dt, 5),
            "planned_eta": _add_days(mat_dt, -10),
            "impact_on_production": "Moderate (Required Before Wiring / Piping Integration)",
            "risk_level": "Medium / Watch",
            "tone": "warning",
        },
        {
            "item_no": "6.4",
            "component": sub_labels[3].replace("Major Component #4 — ", ""),
            "sub_vendor_origin": "Primary OEM / In-House Factory",
            "prerequisite_gate": "Internal Engineering Release",
            "po_status": "Engineering & BOM Frozen" if m_by_key.get("kom", {}).get("actual_date") else "Pending KOM",
            "lead_time_days": 55,
            "planned_order_date": _add_days(kom_dt, 25),
            "planned_eta": _add_days(mat_dt, -14),
            "impact_on_production": "Primary Core Equipment",
            "risk_level": "On Track" if not package_highlights["is_ros_delayed"] else "High Risk (ROS Delay)",
            "tone": "success" if not package_highlights["is_ros_delayed"] else "danger",
        },
    ]

    attachment4_summary = {
        "package_code": pkg_code,
        "total_components": len(expediting_items),
        "components_on_hold": sum(1 for x in expediting_items if "Hold" in x["po_status"]),
        "target_material_delivery_date": mat_dt,
        "expediting_narrative": (
            f"Attachment 4 Material Expediting Summary ({pkg_code}): {sum(1 for x in expediting_items if 'Hold' in x['po_status'])} of 4 major component groups "
            f"are currently held pending Key Document approval ({vdrl_status.get('fabrication_impact_overdue', 0)} sub-order impact VDRL doc(s) overdue). "
            f"All long-lead sub-orders must be released by {app_dt} to achieve Material Delivery on {mat_dt} and protect the Fixed ROS Date ({package_highlights['ros_date'] or '—'})."
        ),
        "items": expediting_items,
    }

    return attachment3_summary, attachment4_summary


def _build_vdrl_status(vdrl_data: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not vdrl_data:
        return {
            "available": False,
            "source_file": "",
            "today_date": "",
            "total_docs": 0,
            "submitted": 0,
            "not_yet": 0,
            "overdue": 0,
            "closed": 0,
            "pct_submitted": 0.0,
            "pct_closed": 0.0,
            "key_docs_total": 0,
            "key_docs_submitted": 0,
            "key_docs_overdue": 0,
            "non_key_total": 0,
            "non_key_submitted": 0,
            "non_key_overdue": 0,
            "production_impact_total": 0,
            "production_impact_overdue": 0,
            "fabrication_impact_total": 0,
            "fabrication_impact_overdue": 0,
            "appr_count": 0,
            "awc_count": 0,
            "category_breakdown": [],
            "documents": [],
        }
    return {
        "available": True,
        **vdrl_data,
    }


def _build_client_doc_status(client_pkg_data: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not client_pkg_data:
        return {
            "available": False,
            "raw_package_code": "",
            "total_submittals": 0,
            "in_review": 0,
            "submitted": 0,
            "in_approval": 0,
            "appr": 0,
            "awc": 0,
            "appr_c": 0,
            "rec": 0,
            "pending_return": 0,
            "overdue_reviews": 0,
            "documents": [],
        }
    return {
        "available": True,
        **client_pkg_data,
    }


def _build_look_ahead_plan(
    pkg_code: str,
    pkg_name: str,
    live_wpr: Optional[Dict[str, Any]],
    vdrl_data: Optional[Dict[str, Any]],
    client_doc_status: Dict[str, Any],
    milestones: List[Dict[str, Any]],
    attachment3_summary: Dict[str, Any],
    attachment4_summary: Dict[str, Any],
    cut_off_date: str,
    is_sample: bool,
) -> Dict[str, Any]:
    """Element 8: Dynamic Package-Specific Look Ahead Plan (7–14 Days) (Comment #3)."""
    wpr_next_activities: List[Dict[str, Any]] = []
    if live_wpr:
        for act in live_wpr.get("next_period_activities", []):
            wpr_next_activities.append({
                "category": f"[{pkg_code}] WPR Next Period Activity",
                "title": act,
                "target_date": f"{cut_off_date} (+7–14d)",
                "source": "Live WPR Sec. 2",
            })
    else:
        # Synthesize package-specific next period activities from its actual VDRL, Client Log, and Attachment 4 gates
        overdue_key = [
            d for d in (vdrl_data or {}).get("documents", [])
            if "OVERDUE" in str(d.get("status", "")).upper() and d.get("is_key") == "Y"
        ]
        if overdue_key:
            top_docs = ", ".join(f"{d['doc_no']} ({d['doc_type']})" for d in overdue_key[:2])
            wpr_next_activities.append({
                "category": f"[{pkg_code}] Priority Key Document Expediting",
                "title": f"Expedite submission of {len(overdue_key)} overdue Key Document(s): {top_docs}",
                "target_date": f"{cut_off_date} (+7d)",
                "source": f"VDRL ({pkg_code})",
            })
        if client_doc_status.get("in_review", 0) > 0:
            wpr_next_activities.append({
                "category": f"[{pkg_code}] Client / PMT Review Follow-Up",
                "title": f"Follow up Client review return for {client_doc_status['in_review']} submittal(s) currently In Review",
                "target_date": f"{cut_off_date} (+7–14d)",
                "source": "Client Doc Log (D2)",
            })
        for item in attachment4_summary.get("items", [])[:2]:
            wpr_next_activities.append({
                "category": f"[{pkg_code}] Attachment 4 Sub-Order Gate",
                "title": f"Release prerequisite ({item['prerequisite_gate']}) for {item['component']}",
                "target_date": item["planned_order_date"] or f"{cut_off_date} (+14d)",
                "source": "WPR Att. 4 Analysis",
            })

    # Gantt tasks active or starting within 14 days
    gantt_window_items: List[Dict[str, Any]] = []
    for t in attachment3_summary.get("tasks", []):
        if t.get("is_summary"):
            continue
        if _task_active_in_window(t.get("start_date", ""), t.get("finish_date", ""), cut_off_date, days=14):
            gantt_window_items.append({
                "category": f"[{pkg_code}] Att. 3 Fabrication / Eng Schedule (7–14d)",
                "title": f"Task #{t['id']}: {t['name']} ({t['duration']})",
                "target_date": f"{t['start_date']} → {t['finish_date']}",
                "source": "WPR Att. 3 Gantt",
            })

    # Package-specific VDRL documents due in next 14 days (or top pending)
    vdrl_upcoming: List[Dict[str, Any]] = []
    ref_dt = (vdrl_data or {}).get("today_date") or cut_off_date
    for doc in (vdrl_data or {}).get("documents", []):
        if doc["status"] in ("SUBMITTED", "CLOSED"):
            continue
        target_dt = doc["forecast_date"] or doc["plan_date"]
        if _within_window(target_dt, ref_dt, days=14):
            vdrl_upcoming.append({
                "category": f"[{pkg_code}] Upcoming VDRL Submission (7–14d)",
                "title": f"{doc['doc_no']} — {doc['title']} ({'Key Doc' if doc['is_key'] == 'Y' else 'Non-Key'})",
                "target_date": target_dt,
                "source": "VDRL Register",
            })

    # Upcoming Key Milestones for this package
    upcoming_milestones: List[Dict[str, Any]] = []
    for m in milestones:
        if not m["actual_date"] and (m["planned_date"] or m["forecast_date"]):
            t_dt = m["forecast_date"] or m["planned_date"]
            upcoming_milestones.append({
                "category": f"[{pkg_code}] Next Key Milestone",
                "title": f"Milestone #{m['no']}: {m['description']}",
                "target_date": t_dt,
                "source": m["source"],
            })
            if len(upcoming_milestones) >= 2:
                break

    return {
        "window_start": cut_off_date,
        "window_days": 14,
        "wpr_next_activities": wpr_next_activities,
        "schedule_tasks_7_14d": gantt_window_items[:8],
        "vdrl_submissions_7_14d": vdrl_upcoming[:8],
        "upcoming_milestones": upcoming_milestones,
        "total_items": (
            len(wpr_next_activities)
            + len(gantt_window_items[:8])
            + len(vdrl_upcoming[:8])
            + len(upcoming_milestones)
        ),
    }


def _build_areas_of_concern(
    pkg_code: str,
    live_wpr: Optional[Dict[str, Any]],
    vdrl_status: Dict[str, Any],
    client_doc_status: Dict[str, Any],
    milestones: List[Dict[str, Any]],
    package_highlights: Dict[str, Any],
    attachment3_summary: Dict[str, Any],
    attachment4_summary: Dict[str, Any],
    is_sample: bool,
    mom_detail: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """Element 9: Package-Specific Area of Concern synthesizing:
    - Fixed PSR ROS Date Delay (Comment #6: ETA after ROS Date)
    - WPR Milestone vs. Gantt Chart ETA Discrepancy
    - MOM Key Concerns & Open Action Items
    - WPR Section 4 Outstanding Issues & Section 3 Material Bottlenecks
    - WPR Attachment 3 Gantt & Attachment 4 Sub-vendor Expediting Risks (Comment #5)
    - Overdue VDRL Key/Production-impact documents & Overdue Client Review submittals
    """
    concerns: List[Dict[str, Any]] = []
    noisy_headers = {"#", "description", "planned", "date", "actual", "receive", "remark", "for", "epc of wellhead"}

    # 1. Master Plan (MM-ASK-1A-GEN01-PJM-SCD-0001) & Fixed PSR ROS Date Breach (Priority #1!)
    if package_highlights.get("is_ros_delayed"):
        concerns.append({
            "severity": "high",
            "category": "MASTER PLAN / ROS DELIVERY DELAY",
            "title": (
                f"[{pkg_code}] Delivery ETA ({package_highlights['effective_delivery_date']}) is "
                f"+{package_highlights['ros_delay_days']} Days AFTER Master Plan / Fixed ROS ({package_highlights['ros_date']})!"
            ),
            "detail": (
                f"Issue: {package_highlights.get('smart_issue', '')} | "
                f"Mitigation (By {package_highlights.get('mitigation_deadline', 'Immediate')}): {package_highlights.get('smart_mitigation', '')}"
            ),
            "source": "Master Plan PJM-SCD-0001 / PSR",
        })
    elif package_highlights.get("health_tier") == "HIGH_RISK":
        concerns.append({
            "severity": "high" if package_highlights.get("is_critical_path") else "medium",
            "category": "High Risk Schedule / Critical Path Gate",
            "title": (
                f"[{pkg_code}] {package_highlights.get('tier_badge_text', 'High Risk Delay')} — "
                f"ETA ({package_highlights['effective_delivery_date']}) vs Master Plan ({package_highlights.get('master_eta_date') or package_highlights['ros_date']})"
            ),
            "detail": (
                f"Issue: {package_highlights.get('smart_issue', '')} | "
                f"Mitigation (By {package_highlights.get('mitigation_deadline', '7d')}): {package_highlights.get('smart_mitigation', '')}"
            ),
            "source": "Master Plan PJM-SCD-0001",
        })

    # 1B. WPR Milestone vs. Gantt Chart ETA Discrepancy Notification
    if package_highlights.get("has_milestone_discrepancy"):
        concerns.append({
            "severity": "high",
            "category": "WPR Milestone vs. Gantt ETA Discrepancy",
            "title": (
                f"[{pkg_code}] WPR Gantt ETA ({package_highlights.get('gantt_eta_date')}) vs. "
                f"Key Milestone #10 ({package_highlights.get('milestone_eta_date')}) — "
                f"+{package_highlights.get('milestone_discrepancy_days')} Days Sea Freight / Transit Difference"
            ),
            "detail": package_highlights.get("milestone_discrepancy_note", ""),
            "source": f"{(live_wpr or {}).get('document_no') or 'WPR'} (Sec. 5 vs Att. 3 Gantt)",
        })

    # 1C. MOM Key Concerns & Open Action Items (if MOM available)
    if mom_detail and mom_detail.get("available"):
        mom_ref = f"{mom_detail.get('document_no', 'MOM')} ({mom_detail.get('cut_off_date', '')})"
        for kc in mom_detail.get("key_concerns", []):
            concerns.append({
                "severity": kc.get("severity", "high"),
                "category": f"MOM {mom_detail.get('mom_no', 'Action')} ({kc.get('status', 'OPEN')})",
                "title": kc.get("title", ""),
                "detail": f"{kc.get('detail', '')} (Action By: {kc.get('action_by', 'Vendor')} | Target: {kc.get('target_date', '—')})",
                "source": mom_ref,
            })

    # 2. Live WPR Section 4 & Section 3 (if live WPR exists)
    if live_wpr:
        wpr_ref = f"{live_wpr.get('document_no') or 'WPR'} ({live_wpr.get('cut_off_date', '')})"
        for issue in live_wpr.get("outstanding_issues", []):
            clean_i = str(issue or "").strip()
            if len(clean_i) <= 4 or clean_i.lower() in noisy_headers:
                continue
            concerns.append({
                "severity": "high",
                "category": "WPR Outstanding Issue",
                "title": f"[{pkg_code}] {clean_i}",
                "detail": f"Flagged under Section 4 (Outstanding Issues) of Weekly Progress Report ({live_wpr.get('report_no', 'WPR')}).",
                "source": f"{wpr_ref} Sec. 4",
            })
        for mat in live_wpr.get("supplier_material_status", []):
            clean_m = str(mat or "").strip()
            if len(clean_m) <= 4 or clean_m.lower() in noisy_headers:
                continue
            concerns.append({
                "severity": "medium",
                "category": "Supplier Material Hold",
                "title": f"[{pkg_code}] {clean_m}",
                "detail": "Material ordering or sub-vendor release bottleneck reported in WPR Section 3 & Attachment 4.",
                "source": f"{wpr_ref} Sec. 3 / Att. 4",
            })

    # 3. Attachment 3 & 4 Critical Path & Sub-Vendor Material Bottlenecks (Comment #5)
    if vdrl_status.get("fabrication_impact_overdue", 0) > 0 or vdrl_status.get("production_impact_overdue", 0) > 0:
        concerns.append({
            "severity": "high",
            "category": "Att. 3 & 4 Fabrication / Sub-Order Gate",
            "title": (
                f"[{pkg_code}] Sub-Vendor Material Release Blocked by "
                f"{vdrl_status.get('production_impact_overdue', 0)} Production-Impact & "
                f"{vdrl_status.get('fabrication_impact_overdue', 0)} Sub-Order Key VDRL Overdue Doc(s)"
            ),
            "detail": attachment4_summary.get("expediting_narrative", ""),
            "source": "Att. 3 & 4 Analysis",
        })

    # 4. VDRL Overdue Documents
    if vdrl_status.get("overdue", 0) > 0:
        key_ov = vdrl_status.get("key_docs_overdue", 0)
        overdue_titles = [
            f"{d['doc_no']} ({d['title']})"
            for d in vdrl_status.get("documents", [])
            if "OVERDUE" in d.get("status", "").upper()
        ][:4]
        concerns.append({
            "severity": "high" if key_ov > 0 else "medium",
            "category": "VDRL Submission Overdue",
            "title": f"[{pkg_code}] {vdrl_status['overdue']} Vendor Document(s) Overdue ({key_ov} Key Docs)",
            "detail": "Overdue items: " + (", ".join(overdue_titles) if overdue_titles else "See VDRL table"),
            "source": f"VDRL ({pkg_code})",
        })

    # 5. Contractual Client Review Log Overdue Items
    if client_doc_status.get("overdue_reviews", 0) > 0:
        concerns.append({
            "severity": "medium",
            "category": "Client / PMT Review Turnaround",
            "title": f"[{pkg_code}] {client_doc_status['overdue_reviews']} Submittal(s) Pending Return Past Plan Finish Date",
            "detail": (
                f"Total Submittals: {client_doc_status['total_submittals']} | "
                f"In Review: {client_doc_status['in_review']} | AWC: {client_doc_status['awc']} | APPR: {client_doc_status['appr']}"
            ),
            "source": "Client Doc Log (D2)",
        })

    # 6. Next Milestone Overdue Check
    rem_ms = package_highlights.get("remaining_days_to_next_milestone")
    if rem_ms is not None and rem_ms < 0:
        concerns.append({
            "severity": "high",
            "category": "Key Milestone Overdue",
            "title": (
                f"[{pkg_code}] Milestone #{package_highlights['next_milestone_no']} "
                f"({package_highlights['next_milestone_name']}) is Overdue by {abs(rem_ms)} Days (Target: {package_highlights['next_milestone_date']})"
            ),
            "detail": package_highlights.get("current_stage_summary", ""),
            "source": "Element 4 Schedule",
        })

    return concerns
