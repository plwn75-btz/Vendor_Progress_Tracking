"""Unified Extraction and Synchronization CLI for Export Data File_ASK.xlsx and Export Data File_Z1F.xlsx.

Extracts all essential vendor tracking data from local PDF decks into separated project workbooks:
- ASK: Export Data File_ASK.xlsx
- Z1F: Export Data File_Z1F.xlsx

Usage:
    python export_dashboard_data.py             # Syncs BOTH projects (ASK & Z1F)
    python export_dashboard_data.py --project ASK
    python export_dashboard_data.py --project Z1F

Upload only the lightweight Excel file(s) to GitHub — heavy PDFs stay local.
"""

import argparse
import glob
import os
import re
import shutil
import sys
from typing import Any, Dict, List

from vendor_dashboard.extractors.excel_sync import (
    sync_lookahead_data_to_excel,
    sync_master_construction_to_excel,
    sync_mom_data_to_excel,
    sync_vdrl_summary_to_excel,
    sync_wpr_data_to_excel,
)
from vendor_dashboard.extractors.wpr_mom_extractor import extract_mom_pdf, extract_wpr_pdf
from vendor_dashboard.ingestors import (
    ASK_PACKAGE_METADATA,
    EXPORT_DATA_FILENAME,
    EXPORT_DATA_FILENAME_ASK,
    EXPORT_DATA_FILENAME_Z1F,
    Z1F_PACKAGE_METADATA,
    _file_freshness_key,
    _parse_kom_pdf_folder,
    export_kom_to_excel,
    get_export_data_filename,
    read_kom_export_file,
)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


def sync_project_export_data(workspace_dir: str, project: str = "ASK") -> Dict[str, Any]:
    """Extract and synchronize vendor tracking data into project-specific Export Data File."""
    proj = str(project or "ASK").strip().upper()
    export_filename = get_export_data_filename(proj)
    export_path = os.path.join(workspace_dir, export_filename)
    legacy_path = os.path.join(workspace_dir, EXPORT_DATA_FILENAME)

    # Initialize from legacy if project-specific file does not exist yet
    if not os.path.isfile(export_path) and os.path.isfile(legacy_path) and proj == "ASK":
        shutil.copy2(legacy_path, export_path)

    stats_summary: Dict[str, Any] = {
        "project": proj,
        "export_filename": export_filename,
        "export_path": export_path,
        "kom_count": 0,
        "wpr_count": 0,
        "mom_count": 0,
        "lookahead_7d_count": 0,
        "lookahead_14d_count": 0,
        "file_size_kb": 0.0,
    }

    print("-" * 55)
    print(f"📦 Synchronizing [{proj}] -> {export_filename}")
    print("-" * 55)

    catalog = ASK_PACKAGE_METADATA if proj == "ASK" else Z1F_PACKAGE_METADATA

    # 1. KOM Sync
    kom_dir = os.path.join(workspace_dir, "KOM")
    if os.path.exists(kom_dir):
        live_kom = _parse_kom_pdf_folder(kom_dir)
        filtered_kom = {k: v for k, v in live_kom.items() if k in catalog}
        merged = read_kom_export_file(export_path)
        merged.update(filtered_kom)
        if merged:
            export_kom_to_excel(merged, export_path, project=proj)
            stats_summary["kom_count"] = len(merged)
            print(f" [KOM] Synced {len(merged)} package baselines for {proj}.")
        else:
            print(f" [KOM] No KOM baselines found for {proj}.")
    else:
        print(" [KOM] KOM directory not found.")

    # 2. WPR Sync
    wpr_dir = os.path.join(workspace_dir, "WPR")
    if os.path.exists(wpr_dir):
        all_wpr_files = sorted(glob.glob(os.path.join(wpr_dir, "*.pdf")), key=_file_freshness_key)
        from vendor_dashboard.ingestors import fmt_date
        candidates_by_pkg: Dict[str, List[Dict[str, Any]]] = {}
        for wf in all_wpr_files:
            upper = os.path.basename(wf).upper()
            is_z1f = "ZTK" in upper or "Z1F" in upper
            if proj == "ASK" and is_z1f:
                continue
            if proj == "Z1F" and not is_z1f:
                continue
            m_pkg = re.search(r"MM-(?:ASK|ZTK)-[A-Z0-9]+-[A-Z0-9]+-([A-Z0-9_]+)-", upper)
            p_key = m_pkg.group(1).split("_")[0] if m_pkg else os.path.basename(wf)

            try:
                import pymupdf
                doc = pymupdf.open(wf)
                hdr_text = "\n".join([doc[i].get_text("text") for i in range(min(4, len(doc)))])
                doc.close()
            except Exception:
                hdr_text = os.path.basename(wf)

            date_m = (
                re.search(r"REPORT\s*DATE\s*:\s*([0-9A-Za-z\-]+)", hdr_text, re.I)
                or re.search(r"CUT\s*OFF\s*(?:PERIOD|DATE)\s*:\s*(?:[0-9A-Za-z\-]+\s*(?:to|-)\s*)?([0-9A-Za-z\-]+)", hdr_text, re.I)
                or re.search(r"Date\.?\s*:\s*([0-9A-Za-z\-]+)", hdr_text, re.I)
            )
            rep_date = fmt_date(date_m.group(1)) if date_m else ""
            rev_m = re.search(r"Rev(?:ision)?\.?\s*:\s*([0-9A-Z]+)", hdr_text, re.I)
            rev = rev_m.group(1) if rev_m else ""
            rep_m = re.search(r"REPORT\s*NO\.?\s*:\s*([0-9A-Za-z\-]+)", hdr_text, re.I)
            rep_no = rep_m.group(1) if rep_m else ""

            candidates_by_pkg.setdefault(p_key, []).append({
                "wf": wf,
                "cut_off_date": rep_date,
                "revision": rev,
                "report_no": rep_no,
                "_mtime": _file_freshness_key(wf),
            })

        wpr_records: List[Dict[str, Any]] = []
        for pkg, c_list in candidates_by_pkg.items():
            def _wpr_export_sort_key(item: Dict[str, Any]):
                cod = str(item.get("cut_off_date") or "").strip()[:10]
                rev_nums = re.findall(r"\d+", str(item.get("revision") or ""))
                rev_val = int(rev_nums[0]) if rev_nums else 0
                rep_nums = re.findall(r"\d+", str(item.get("report_no") or ""))
                rep_val = int(rep_nums[0]) if rep_nums else 0
                return (cod, rev_val, rep_val, item.get("_mtime", 0.0))

            c_list.sort(key=_wpr_export_sort_key)
            for cand in reversed(c_list):
                wf = cand["wf"]
                try:
                    rec = extract_wpr_pdf(wf)
                    summ = rec.get("summary", {})
                    if (
                        rec.get("milestones")
                        or rec.get("gantt_tasks")
                        or float(summ.get("overall_plan_pct") or 0.0) > 0.0
                        or float(summ.get("overall_actual_pct") or 0.0) > 0.0
                    ):
                        wpr_records.append(rec)
                        break
                except Exception as e:
                    print(f"⚠️ Error parsing WPR '{os.path.basename(wf)}': {e}")
        if wpr_records:
            wpr_stats = sync_wpr_data_to_excel(export_path, wpr_records)
            stats_summary["wpr_count"] = len(wpr_records)
            print(f" [WPR] Synced {len(wpr_records)} {proj} reports:")
            print(f"       • Summaries: {wpr_stats['summaries']}")
            print(f"       • S-Curve points: {wpr_stats['scurve']}")
            print(f"       • Gantt Tasks: {wpr_stats['gantt']}")
            print(f"       • Sub-Orders: {wpr_stats['sub_orders']}")
        else:
            print(f" [WPR] No {proj} WPR PDF files found.")
    else:
        print(" [WPR] WPR directory not found.")

    # 3. MOM Sync
    mom_dir = os.path.join(workspace_dir, "MOM")
    if os.path.exists(mom_dir):
        all_mom_files = sorted(glob.glob(os.path.join(mom_dir, "*.pdf")), key=_file_freshness_key)
        candidates_mom_by_pkg: Dict[str, List[Dict[str, Any]]] = {}
        for mf in all_mom_files:
            upper = os.path.basename(mf).upper()
            is_z1f = "ZTK" in upper or "Z1F" in upper
            if proj == "ASK" and is_z1f:
                continue
            if proj == "Z1F" and not is_z1f:
                continue
            try:
                rec = extract_mom_pdf(mf)
                meta = rec.get("metadata", {})
                pkg = meta.get("package")
                if pkg:
                    rec["_mf"] = mf
                    rec["_mtime"] = _file_freshness_key(mf)
                    candidates_mom_by_pkg.setdefault(pkg, []).append(rec)
            except Exception as e:
                print(f"⚠️ Error parsing MOM '{os.path.basename(mf)}': {e}")

        mom_records: List[Dict[str, Any]] = []
        for pkg, c_list in candidates_mom_by_pkg.items():
            def _mom_export_sort_key(item: Dict[str, Any]):
                meta = item.get("metadata", {})
                m_date = str(meta.get("meeting_date") or meta.get("cut_off_date") or "").strip()[:10]
                mom_nums = re.findall(r"\d+", str(meta.get("mom_no") or ""))
                mom_val = int(mom_nums[0]) if mom_nums else 0
                return (m_date, mom_val, item.get("_mtime", 0.0))

            c_list.sort(key=_mom_export_sort_key)
            mom_records.append(c_list[-1])
        if mom_records:
            mom_stats = sync_mom_data_to_excel(export_path, mom_records)
            stats_summary["mom_count"] = len(mom_records)
            print(f" [MOM] Synced {len(mom_records)} {proj} meetings:")
            print(f"       • Action items: {mom_stats['action_items']}")
            print(f"       • Procurement items: {mom_stats['procurement_status']}")
        else:
            print(f" [MOM] No {proj} MOM PDF files found.")
    else:
        print(" [MOM] MOM directory not found.")

    # 4. Look-Ahead Hub, VDRL Summary & Master Construction Sync
    try:
        from vendor_dashboard.ingestors import parse_all_vdrl_files
        from vendor_dashboard.reconciler import build_project_snapshot
        snap = build_project_snapshot(workspace_dir, project=proj, mode="live")
        la_hub = snap.get("lookahead_hub", {})
        l7 = la_hub.get("activities_7d", [])
        l14 = la_hub.get("activities_14d", [])
        if l7 or l14:
            la_stats = sync_lookahead_data_to_excel(export_path, l7, l14)
            stats_summary["lookahead_7d_count"] = la_stats["lookahead_7d"]
            stats_summary["lookahead_14d_count"] = la_stats["lookahead_14d"]
            print(f" [LookAhead] Synced {la_stats['lookahead_7d']} 7-day & {la_stats['lookahead_14d']} 14-day lookahead activities.")

        vdrl_dir = os.path.join(workspace_dir, "VDRL")
        vdrl_map = parse_all_vdrl_files(vdrl_dir, project_filter=proj)
        if vdrl_map:
            vdrl_cnt = sync_vdrl_summary_to_excel(export_path, vdrl_map)
            stats_summary["vdrl_summary_count"] = vdrl_cnt
            print(f" [VDRL] Synced {vdrl_cnt} package VDRL summaries.")

        cm = snap.get("construction_monitoring", {})
        mc_cnt = sync_master_construction_to_excel(
            export_path,
            cm.get("all_activities", []),
            cm.get("ti_activities", []),
        )
        if mc_cnt:
            stats_summary["master_construction_count"] = mc_cnt
            print(f" [Construction] Synced {mc_cnt} Master Schedule construction & T&I activities.")
    except Exception as e:
        print(f"⚠️ Error syncing Look-Ahead / Construction / VDRL activities for {proj}: {e}")

    if os.path.isfile(export_path):
        stats_summary["file_size_kb"] = round(os.path.getsize(export_path) / 1024, 1)

    # Mirror to legacy "Export Data File.xlsx" for backward compatibility if ASK
    if proj == "ASK" and os.path.isfile(export_path):
        try:
            shutil.copy2(export_path, legacy_path)
        except Exception:
            pass

    return stats_summary


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract and synchronize vendor data into project-specific Excel files.")
    parser.add_argument("--project", choices=["ASK", "Z1F", "ALL"], default="ALL", help="Target project (ASK, Z1F, or ALL)")
    args = parser.parse_args()

    print("=" * 60)
    print("🚀 Running Unified Vendor Data Extraction & Sync")
    print(f"Workspace: {ROOT_DIR}")
    print(f"Selected Scope: {args.project}")
    print("=" * 60)

    projects = ["ASK", "Z1F"] if args.project == "ALL" else [args.project]

    for p in projects:
        res = sync_project_export_data(ROOT_DIR, project=p)
        print(f"✅ [{p}] Complete: {res['export_filename']} ({res['file_size_kb']} KB)")

    print("=" * 60)
    print("All designated project export files are ready to commit to GitHub:")
    for p in projects:
        fn = get_export_data_filename(p)
        print(f"  • {fn}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())

