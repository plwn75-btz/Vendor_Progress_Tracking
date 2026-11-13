"""Excel Synchronization Module for Export Data File.xlsx.

Manages upsert (update-or-insert) logic across 6 relational sheets:
- KOM_Packages (preserved baseline)
- WPR_Weekly_Summary
- WPR_S_Curve
- WPR_Fabrication_Gantt
- WPR_Sub_Orders
- MOM_Action_Items
- MOM_Procurement_Status
"""

import os
from typing import Dict, Any, List, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
DATA_FONT = Font(name="Calibri", size=10)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
BORDER_THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def _apply_header_style(ws, headers: List[str]):
    ws.append(headers)
    ws.row_dimensions[1].height = 24
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = max(len(h) + 4, 14)


def _get_sheet_headers(sheet_name: str) -> List[str]:
    headers_map = {
        "WPR_Weekly_Summary": [
            "Package", "Report_No", "Cutoff_Date", "PO_Number", "Vendor_Name", "Target_Delivery",
            "Overall_Plan_Pct", "Overall_Actual_Pct", "Incremental_Plan_Pct", "Incremental_Actual_Pct",
            "Deviation_Pct", "Status_Tone", "Executive_Highlights", "Areas_of_Concern", "Critical_Path_Notes"
        ],
        "WPR_S_Curve": [
            "Package", "Report_No", "Cutoff_Date", "Week_No", "Week_Date",
            "Plan_Cum_Pct", "Actual_Cum_Pct", "Plan_Inc_Pct", "Actual_Inc_Pct"
        ],
        "WPR_Fabrication_Gantt": [
            "Package", "Report_No", "Cutoff_Date", "Task_ID", "Task_Name",
            "Duration", "Start_Date", "Finish_Date", "Predecessors", "Is_Milestone", "Phase_Category"
        ],
        "WPR_Sub_Orders": [
            "Package", "Report_No", "Cutoff_Date", "Sub_Component", "Material_Description",
            "Sub_Vendor", "PO_Status", "PO_Date", "Req_Delivery", "Est_Delivery", "Remarks"
        ],
        "MOM_Action_Items": [
            "Package", "MOM_No", "Meeting_Date", "Item_No", "Action_By", "Target_Date", "Status", "Description"
        ],
        "MOM_Procurement_Status": [
            "Package", "MOM_No", "Meeting_Date", "Item_No", "Component", "Plan_Date", "Actual_Date", "Status"
        ],
        "LookAhead_7Days": [
            "Scope", "Package_or_ID", "Package_Name", "Activity_ID", "Activity_Name",
            "Start_Date", "Finish_Date", "Duration_Days", "Critical_Path", "Category", "Source", "Status"
        ],
        "LookAhead_14Days": [
            "Scope", "Package_or_ID", "Package_Name", "Activity_ID", "Activity_Name",
            "Start_Date", "Finish_Date", "Duration_Days", "Critical_Path", "Category", "Source", "Status"
        ],
        "VDRL_Summary": [
            "Package_Code", "Package_Name", "Discipline", "Source_File", "Today_Date",
            "Total_Docs", "Submitted", "Not_Yet", "Overdue", "Closed", "Pct_Submitted", "Pct_Closed",
            "Key_Docs_Total", "Key_Docs_Submitted", "Key_Docs_Overdue",
            "Non_Key_Total", "Non_Key_Submitted", "Non_Key_Overdue",
            "Production_Impact_Total", "Production_Impact_Overdue",
            "Fabrication_Impact_Total", "Fabrication_Impact_Overdue",
            "APPR_Count", "AWC_Count"
        ],
        "Master_Construction": [
            "Activity_ID", "Activity_Name", "Category", "Start_Date", "Finish_Date",
            "Duration_Days", "Total_Float_Days", "Critical_Path", "Page", "Source"
        ],
    }
    return headers_map.get(sheet_name, [])


def sync_wpr_data_to_excel(excel_path: str, wpr_records: List[Dict[str, Any]]) -> Dict[str, int]:
    """Upsert WPR records (summaries, scurve, gantt, suborders) into Export Data File.xlsx."""
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    stats = {"summaries": 0, "scurve": 0, "gantt": 0, "sub_orders": 0}

    # 1. Sheet: WPR_Weekly_Summary
    s_name = "WPR_Weekly_Summary"
    if s_name not in wb.sheetnames:
        ws = wb.create_sheet(title=s_name)
        _apply_header_style(ws, _get_sheet_headers(s_name))
    else:
        ws = wb[s_name]

    headers = [cell.value for cell in ws[1]]
    key_col_pkg = headers.index("Package") + 1 if "Package" in headers else 1
    key_col_rpt = headers.index("Report_No") + 1 if "Report_No" in headers else 2

    # Map existing rows by (Package, Report_No)
    existing_summary_rows = {}
    for r in range(2, ws.max_row + 1):
        p_val = ws.cell(row=r, column=key_col_pkg).value
        r_val = ws.cell(row=r, column=key_col_rpt).value
        if p_val and r_val:
            existing_summary_rows[(str(p_val).strip(), str(r_val).strip())] = r

    for wpr in wpr_records:
        s = wpr["summary"]
        key = (str(s["package"]).strip(), str(s["report_no"]).strip())
        row_vals = [
            s.get("package", ""), s.get("report_no", ""), s.get("cutoff_date", ""),
            s.get("po_number", ""), s.get("vendor_name", ""), s.get("target_delivery", ""),
            s.get("overall_plan_pct", 0.0), s.get("overall_actual_pct", 0.0),
            s.get("incremental_plan_pct", 0.0), s.get("incremental_actual_pct", 0.0),
            s.get("deviation_pct", 0.0), s.get("status_tone", "info"),
            s.get("executive_highlights", ""), s.get("areas_of_concern", ""), s.get("critical_path_notes", "")
        ]

        if key in existing_summary_rows:
            target_row = existing_summary_rows[key]
            for c_idx, val in enumerate(row_vals, 1):
                ws.cell(row=target_row, column=c_idx, value=val)
        else:
            ws.append(row_vals)
            existing_summary_rows[key] = ws.max_row
        stats["summaries"] += 1

    # 2. Sheet: WPR_S_Curve
    s_name = "WPR_S_Curve"
    if s_name not in wb.sheetnames:
        ws = wb.create_sheet(title=s_name)
        _apply_header_style(ws, _get_sheet_headers(s_name))
    else:
        ws = wb[s_name]

    # Map existing rows by (Package, Report_No, Week_No)
    headers = [cell.value for cell in ws[1]]
    key_col_pkg = headers.index("Package") + 1 if "Package" in headers else 1
    key_col_rpt = headers.index("Report_No") + 1 if "Report_No" in headers else 2
    key_col_wk = headers.index("Week_No") + 1 if "Week_No" in headers else 4

    existing_scurve_rows = {}
    for r in range(2, ws.max_row + 1):
        p_val = ws.cell(row=r, column=key_col_pkg).value
        r_val = ws.cell(row=r, column=key_col_rpt).value
        w_val = ws.cell(row=r, column=key_col_wk).value
        if p_val and r_val and w_val is not None:
            existing_scurve_rows[(str(p_val).strip(), str(r_val).strip(), int(w_val))] = r

    for wpr in wpr_records:
        for pt in wpr.get("scurve", []):
            key = (str(pt["package"]).strip(), str(pt["report_no"]).strip(), int(pt["week_no"]))
            row_vals = [
                pt["package"], pt["report_no"], pt["cutoff_date"], pt["week_no"], pt["week_date"],
                pt["plan_cum_pct"], pt["actual_cum_pct"], pt["plan_inc_pct"], pt["actual_inc_pct"]
            ]
            if key in existing_scurve_rows:
                target_row = existing_scurve_rows[key]
                for c_idx, val in enumerate(row_vals, 1):
                    ws.cell(row=target_row, column=c_idx, value=val)
            else:
                ws.append(row_vals)
                existing_scurve_rows[key] = ws.max_row
            stats["scurve"] += 1

    # 3. Sheet: WPR_Fabrication_Gantt
    s_name = "WPR_Fabrication_Gantt"
    if s_name not in wb.sheetnames:
        ws = wb.create_sheet(title=s_name)
        _apply_header_style(ws, _get_sheet_headers(s_name))
    else:
        ws = wb[s_name]

    headers = [cell.value for cell in ws[1]]
    key_col_pkg = headers.index("Package") + 1 if "Package" in headers else 1
    key_col_rpt = headers.index("Report_No") + 1 if "Report_No" in headers else 2
    key_col_tid = headers.index("Task_ID") + 1 if "Task_ID" in headers else 4

    existing_gantt_rows = {}
    for r in range(2, ws.max_row + 1):
        p_val = ws.cell(row=r, column=key_col_pkg).value
        r_val = ws.cell(row=r, column=key_col_rpt).value
        t_val = ws.cell(row=r, column=key_col_tid).value
        if p_val and r_val and t_val is not None:
            existing_gantt_rows[(str(p_val).strip(), str(r_val).strip(), int(t_val))] = r

    for wpr in wpr_records:
        for t in wpr.get("gantt_tasks", []):
            key = (str(t["package"]).strip(), str(t["report_no"]).strip(), int(t["task_id"]))
            row_vals = [
                t["package"], t["report_no"], t["cutoff_date"], t["task_id"], t["task_name"],
                t["duration"], t["start_date"], t["finish_date"], t["predecessors"], t["is_milestone"], t["phase_category"]
            ]
            if key in existing_gantt_rows:
                target_row = existing_gantt_rows[key]
                for c_idx, val in enumerate(row_vals, 1):
                    ws.cell(row=target_row, column=c_idx, value=val)
            else:
                ws.append(row_vals)
                existing_gantt_rows[key] = ws.max_row
            stats["gantt"] += 1

    # 4. Sheet: WPR_Sub_Orders
    s_name = "WPR_Sub_Orders"
    if s_name not in wb.sheetnames:
        ws = wb.create_sheet(title=s_name)
        _apply_header_style(ws, _get_sheet_headers(s_name))
    else:
        ws = wb[s_name]

    headers = [cell.value for cell in ws[1]]
    key_col_pkg = headers.index("Package") + 1 if "Package" in headers else 1
    key_col_rpt = headers.index("Report_No") + 1 if "Report_No" in headers else 2
    key_col_comp = headers.index("Sub_Component") + 1 if "Sub_Component" in headers else 4

    existing_sub_rows = {}
    for r in range(2, ws.max_row + 1):
        p_val = ws.cell(row=r, column=key_col_pkg).value
        r_val = ws.cell(row=r, column=key_col_rpt).value
        c_val = ws.cell(row=r, column=key_col_comp).value
        if p_val and r_val and c_val:
            existing_sub_rows[(str(p_val).strip(), str(r_val).strip(), str(c_val).strip())] = r

    for wpr in wpr_records:
        for so in wpr.get("sub_orders", []):
            key = (str(so["package"]).strip(), str(so["report_no"]).strip(), str(so["sub_component"]).strip())
            row_vals = [
                so["package"], so["report_no"], so["cutoff_date"], so["sub_component"], so["material_description"],
                so["sub_vendor"], so["po_status"], so["po_date"], so["req_delivery"], so["est_delivery"], so.get("remarks", "")
            ]
            if key in existing_sub_rows:
                target_row = existing_sub_rows[key]
                for c_idx, val in enumerate(row_vals, 1):
                    ws.cell(row=target_row, column=c_idx, value=val)
            else:
                ws.append(row_vals)
                existing_sub_rows[key] = ws.max_row
            stats["sub_orders"] += 1

    wb.save(excel_path)
    return stats


def sync_mom_data_to_excel(excel_path: str, mom_records: List[Dict[str, Any]]) -> Dict[str, int]:
    """Upsert MOM records (action items and procurement status) into Export Data File.xlsx."""
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    stats = {"action_items": 0, "procurement_status": 0}

    # 1. Sheet: MOM_Action_Items
    s_name = "MOM_Action_Items"
    if s_name not in wb.sheetnames:
        ws = wb.create_sheet(title=s_name)
        _apply_header_style(ws, _get_sheet_headers(s_name))
    else:
        ws = wb[s_name]

    headers = [cell.value for cell in ws[1]]
    key_col_pkg = headers.index("Package") + 1 if "Package" in headers else 1
    key_col_mom = headers.index("MOM_No") + 1 if "MOM_No" in headers else 2
    key_col_it = headers.index("Item_No") + 1 if "Item_No" in headers else 4

    existing_action_rows = {}
    for r in range(2, ws.max_row + 1):
        p_val = ws.cell(row=r, column=key_col_pkg).value
        m_val = ws.cell(row=r, column=key_col_mom).value
        i_val = ws.cell(row=r, column=key_col_it).value
        if p_val and m_val and i_val:
            existing_action_rows[(str(p_val).strip(), str(m_val).strip(), str(i_val).strip())] = r

    for mom in mom_records:
        for ai in mom.get("action_items", []):
            key = (str(ai["package"]).strip(), str(ai["mom_no"]).strip(), str(ai["item_no"]).strip())
            row_vals = [
                ai["package"], ai["mom_no"], ai["meeting_date"], ai["item_no"],
                ai.get("action_by", ""), ai.get("target_date", ""), ai.get("status", "OPEN"), ai.get("description", "")
            ]
            if key in existing_action_rows:
                target_row = existing_action_rows[key]
                for c_idx, val in enumerate(row_vals, 1):
                    ws.cell(row=target_row, column=c_idx, value=val)
            else:
                ws.append(row_vals)
                existing_action_rows[key] = ws.max_row
            stats["action_items"] += 1

    # 2. Sheet: MOM_Procurement_Status
    s_name = "MOM_Procurement_Status"
    if s_name not in wb.sheetnames:
        ws = wb.create_sheet(title=s_name)
        _apply_header_style(ws, _get_sheet_headers(s_name))
    else:
        ws = wb[s_name]

    headers = [cell.value for cell in ws[1]]
    key_col_pkg = headers.index("Package") + 1 if "Package" in headers else 1
    key_col_mom = headers.index("MOM_No") + 1 if "MOM_No" in headers else 2
    key_col_it = headers.index("Item_No") + 1 if "Item_No" in headers else 4

    existing_proc_rows = {}
    for r in range(2, ws.max_row + 1):
        p_val = ws.cell(row=r, column=key_col_pkg).value
        m_val = ws.cell(row=r, column=key_col_mom).value
        i_val = ws.cell(row=r, column=key_col_it).value
        if p_val and m_val and i_val is not None:
            existing_proc_rows[(str(p_val).strip(), str(m_val).strip(), int(i_val))] = r

    for mom in mom_records:
        for ps in mom.get("procurement_status", []):
            key = (str(ps["package"]).strip(), str(ps["mom_no"]).strip(), int(ps["item_no"]))
            row_vals = [
                ps["package"], ps["mom_no"], ps["meeting_date"], ps["item_no"],
                ps.get("component", ""), ps.get("plan_date", ""), ps.get("actual_date", ""), ps.get("status", "In Progress")
            ]
            if key in existing_proc_rows:
                target_row = existing_proc_rows[key]
                for c_idx, val in enumerate(row_vals, 1):
                    ws.cell(row=target_row, column=c_idx, value=val)
            else:
                ws.append(row_vals)
                existing_proc_rows[key] = ws.max_row
            stats["procurement_status"] += 1

    wb.save(excel_path)
    return stats


def sync_lookahead_data_to_excel(
    excel_path: str,
    lookahead_7d: List[Dict[str, Any]],
    lookahead_14d: List[Dict[str, Any]],
) -> Dict[str, int]:
    """Synchronize 7-Day and 14-Day Look-Ahead activities into Export Data File.xlsx.
    Creates/updates sheets 'LookAhead_7Days' and 'LookAhead_14Days'.
    """
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    stats = {"lookahead_7d": len(lookahead_7d), "lookahead_14d": len(lookahead_14d)}

    sheets_data = [
        ("LookAhead_7Days", lookahead_7d),
        ("LookAhead_14Days", lookahead_14d),
    ]

    for s_name, data_list in sheets_data:
        if s_name in wb.sheetnames:
            wb.remove(wb[s_name])
        ws = wb.create_sheet(title=s_name)
        headers = _get_sheet_headers(s_name)
        _apply_header_style(ws, headers)

        for item in data_list:
            row_vals = [
                item.get("scope", ""),
                item.get("package_or_id", ""),
                item.get("package_name", ""),
                item.get("activity_id", ""),
                item.get("activity_name", ""),
                item.get("start_date", ""),
                item.get("finish_date", ""),
                item.get("duration_days"),
                item.get("critical_path", "NO"),
                item.get("category", ""),
                item.get("source", ""),
                item.get("status", ""),
            ]
            ws.append(row_vals)
            curr_row = ws.max_row
            for col_idx in range(1, len(row_vals) + 1):
                c = ws.cell(row=curr_row, column=col_idx)
                c.font = DATA_FONT
                c.border = BORDER_THIN
                if col_idx in (1, 2, 6, 7, 8, 9, 12):
                    c.alignment = ALIGN_CENTER
                else:
                    c.alignment = ALIGN_LEFT

    wb.save(excel_path)
    return stats


def sync_vdrl_summary_to_excel(
    excel_path: str,
    vdrl_map: Dict[str, Dict[str, Any]],
) -> int:
    """Synchronize per-package VDRL summary metrics into Export Data File_*.xlsx ('VDRL_Summary' sheet)."""
    if not vdrl_map:
        return 0
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    s_name = "VDRL_Summary"
    if s_name in wb.sheetnames:
        wb.remove(wb[s_name])
    ws = wb.create_sheet(title=s_name)
    _apply_header_style(ws, _get_sheet_headers(s_name))

    count = 0
    for code in sorted(vdrl_map.keys()):
        item = vdrl_map[code]
        row_vals = [
            item.get("package_code", code),
            item.get("package_name", ""),
            item.get("discipline", ""),
            item.get("source_file", ""),
            item.get("today_date", ""),
            int(item.get("total_docs", 0) or 0),
            int(item.get("submitted", 0) or 0),
            int(item.get("not_yet", 0) or 0),
            int(item.get("overdue", 0) or 0),
            int(item.get("closed", 0) or 0),
            float(item.get("pct_submitted", 0.0) or 0.0),
            float(item.get("pct_closed", 0.0) or 0.0),
            int(item.get("key_docs_total", 0) or 0),
            int(item.get("key_docs_submitted", 0) or 0),
            int(item.get("key_docs_overdue", 0) or 0),
            int(item.get("non_key_total", 0) or 0),
            int(item.get("non_key_submitted", 0) or 0),
            int(item.get("non_key_overdue", 0) or 0),
            int(item.get("production_impact_total", 0) or 0),
            int(item.get("production_impact_overdue", 0) or 0),
            int(item.get("fabrication_impact_total", 0) or 0),
            int(item.get("fabrication_impact_overdue", 0) or 0),
            int(item.get("appr_count", 0) or 0),
            int(item.get("awc_count", 0) or 0),
        ]
        ws.append(row_vals)
        curr_row = ws.max_row
        for col_idx in range(1, len(row_vals) + 1):
            c = ws.cell(row=curr_row, column=col_idx)
            c.font = DATA_FONT
            c.border = BORDER_THIN
            c.alignment = ALIGN_LEFT if col_idx in (2, 4) else ALIGN_CENTER
        count += 1

    wb.save(excel_path)
    return count


def sync_master_construction_to_excel(
    excel_path: str,
    construction_acts: List[Dict[str, Any]],
    ti_acts: Optional[List[Dict[str, Any]]] = None,
) -> int:
    """Synchronize deduplicated Master Schedule construction & T&I activities into 'Master_Construction' sheet."""
    merged: Dict[tuple, Dict[str, Any]] = {}
    for act in list(construction_acts or []) + list(ti_acts or []):
        aid = str(act.get("activity_id") or "").strip()
        s = str(act.get("start_date") or "")[:10]
        f = str(act.get("finish_date") or "")[:10]
        if not aid or not (s or f):
            continue
        k = (aid, s, f)
        if k not in merged:
            merged[k] = act

    if not merged:
        return 0

    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

    s_name = "Master_Construction"
    if s_name in wb.sheetnames:
        wb.remove(wb[s_name])
    ws = wb.create_sheet(title=s_name)
    _apply_header_style(ws, _get_sheet_headers(s_name))

    rows = sorted(
        merged.values(),
        key=lambda x: (str(x.get("start_date") or ""), str(x.get("finish_date") or ""), str(x.get("activity_id") or "")),
    )
    for act in rows:
        row_vals = [
            act.get("activity_id", ""),
            act.get("activity_name", ""),
            act.get("category", "Fabrication & Erection"),
            act.get("start_date", ""),
            act.get("finish_date", ""),
            act.get("duration_days"),
            act.get("total_float_days"),
            "YES" if act.get("is_critical_path") else "NO",
            act.get("page"),
            act.get("source", ""),
        ]
        ws.append(row_vals)
        curr_row = ws.max_row
        for col_idx in range(1, len(row_vals) + 1):
            c = ws.cell(row=curr_row, column=col_idx)
            c.font = DATA_FONT
            c.border = BORDER_THIN
            c.alignment = ALIGN_LEFT if col_idx in (2, 3, 10) else ALIGN_CENTER

    wb.save(excel_path)
    return len(rows)



