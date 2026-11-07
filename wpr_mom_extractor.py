"""WPR and MOM PDF Ingestion and Data Extraction Module.

Extracts structured engineering, progress, Gantt tasks, S-curve points,
procurement sub-orders, and action items from:
- Weekly Progress Reports (WPR)
- Minutes of Meeting (MOM)
Supports both file paths and byte streams.
"""

import os
import re
from typing import Dict, Any, List, Optional, Union
import pymupdf


def _clean_str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _parse_date(text: str, mdy_first: bool = False) -> Optional[str]:
    """Extract and normalize date to YYYY-MM-DD or standard format."""
    if not text:
        return None
    # Matches dd-mm-yy, dd-mm-yyyy, dd/mm/yy, dd/mm/yyyy (or mm/dd/yy when mdy_first=True)
    m1 = re.search(r'\b(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})\b', text)
    if m1:
        p1, p2, y = m1.groups()
        if len(y) == 2:
            y = f"20{y}"
        v1, v2 = int(p1), int(p2)
        if mdy_first:
            if v1 <= 12 and v2 <= 31:
                m_val, d_val = v1, v2
            else:
                d_val, m_val = v1, v2
        else:
            if v2 <= 12:
                d_val, m_val = v1, v2
            elif v1 <= 12:
                m_val, d_val = v1, v2
            else:
                d_val, m_val = v1, v2
        return f"{y}-{m_val:02d}-{d_val:02d}"
    # Matches dd/Mon/yy or dd Mon yy, e.g. 31/Jul/26 or 8 Feb 27
    m2 = re.search(r'\b(\d{1,2})[-/\s]+([A-Za-z]{3})[a-z]*[-/\s]+(\d{2,4})\b', text)
    if m2:
        d, mon_str, y = m2.groups()
        months = {
            "jan": "01", "feb": "02", "mar": "03", "apr": "04",
            "may": "05", "jun": "06", "jul": "07", "aug": "08",
            "sep": "09", "oct": "10", "nov": "11", "dec": "12"
        }
        m_num = months.get(mon_str.lower()[:3], "01")
        if len(y) == 2:
            y = f"20{y}"
        return f"{y}-{m_num}-{int(d):02d}"
    return text.strip()


def extract_wpr_pdf(source: Union[str, bytes], filename: str = "") -> Dict[str, Any]:
    """Parse a WPR PDF and extract:
    - Metadata (Package, Report_No, Cutoff_Date, PO, Vendor)
    - Highlights & Areas of Concern
    - Overall Progress (Section 6)
    - Key Milestones (Section 5)
    - Progress S-Curve (Attachment 2)
    - Microsoft Project Gantt tasks (Attachment 3)
    - Sub-Order Status Report / Material Expediting (Attachment 4)
    """
    if isinstance(source, bytes):
        doc = pymupdf.open(stream=source, filetype="pdf")
    else:
        doc = pymupdf.open(source)
        if not filename:
            filename = os.path.basename(source)

    # 1. Detect Package Code and Report No
    pkg_match = re.search(r'MM-(?:ASK-1A|ZTK-1F)-GEN\w*-([A-Z0-9]+)-', filename, re.IGNORECASE)
    if not pkg_match:
        pkg_match = re.search(r'MM-[A-Z0-9]+-[A-Z0-9]+-(?:GEN\w*-)?([A-Z0-9]+)-', filename, re.IGNORECASE)
    pkg_code = pkg_match.group(1).upper() if pkg_match else ""

    rpt_match = re.search(r'WPR[-_]?(\d+)', filename, re.IGNORECASE)
    report_no = f"WPR-{int(rpt_match.group(1)):04d}" if rpt_match else ""

    full_text_pages = [page.get_text() for page in doc]
    doc.close()
    full_doc_text = "\n".join(full_text_pages)

    if not pkg_code:
        for code in ["CIP", "PLC", "WHCP", "IGS", "ABV", "BDV", "ESD", "PIG", "TEG", "FLARE"]:
            if code in full_doc_text.upper():
                pkg_code = code
                break
        if not pkg_code:
            pkg_code = "GEN"

    if not report_no:
        rpt_doc_m = re.search(r'REPORT\s*NO\.?\s*:\s*WPR[-_]?(\d+)', full_doc_text, re.IGNORECASE)
        if rpt_doc_m:
            report_no = f"WPR-{int(rpt_doc_m.group(1)):04d}"
        else:
            report_no = "WPR-0001"

    # 2. Extract Metadata from Header / Cover
    cutoff_date = ""
    po_no = ""
    vendor_name = ""
    target_delivery = ""

    po_m = re.search(r'P/?O\s*NO\.?\s*:\s*([^\s\n\r]+)', full_doc_text, re.IGNORECASE)
    if po_m:
        po_no = po_m.group(1).strip()

    vendor_m = re.search(r'Vendor\s*:\s*([^\n\r]+)', full_doc_text, re.IGNORECASE)
    if vendor_m:
        vendor_name = vendor_m.group(1).strip()

    deliv_m = re.search(r'CONTRACTUAL\s*DELIVERY\s*DATE\s*:\s*([^\n\r]+)', full_doc_text, re.IGNORECASE)
    if deliv_m:
        target_delivery = _parse_date(deliv_m.group(1)) or deliv_m.group(1).strip()

    rep_date_m = (
        re.search(r'REPORT\s*DATE\s*:\s*([0-9A-Za-z\-/]+)', full_doc_text, re.IGNORECASE)
        or re.search(r'CUT\s*OFF\s*(?:PERIOD|DATE)\s*:\s*(?:[0-9A-Za-z\-/]+\s*(?:to|-)\s*)?([0-9A-Za-z\-/]+)', full_doc_text, re.IGNORECASE)
        or re.search(r'Date\s*:\s*([0-9A-Za-z\-/]+)', full_doc_text, re.IGNORECASE)
    )
    if rep_date_m:
        raw_d = rep_date_m.group(1).strip()
        parsed_d = _parse_date(raw_d)
        if parsed_d and len(parsed_d) == 10 and parsed_d[0].isdigit():
            cutoff_date = parsed_d
        elif raw_d:
            cutoff_date = raw_d

    # 3. Extract Section 1, 2, 4 Narrative (Activities, Highlights, Concerns)
    this_period_act = []
    concerns = []

    def _is_boilerplate(txt: str) -> bool:
        t = txt.upper().strip()
        if len(t) <= 4 or t.startswith(":"):
            return True
        if re.match(r'^DATE\s*:', t) or re.match(r'^REV', t):
            return True
        return any(b in t for b in [
            "PROGRESS REPORT", "EPC OF WELLHEAD", "DEVELOPMENT",
            "DOCUMENT NO", "PROJECT PHASE", "REVISION", "DATE :",
            "CHEMICAL INJECTION", "ATTACHMENT", "PAGE "
        ])

    m_act1 = re.search(r'1\.\s*THIS\s*PERIOD\s*ACTIVITIES(.*?)(?=2\.\s*NEXT\s*PERIOD)', full_doc_text, re.DOTALL | re.IGNORECASE)
    if m_act1:
        lines = [l.strip("- ").strip() for l in m_act1.group(1).split("\n") if l.strip("- ").strip()]
        this_period_act = [l for l in lines if len(l) > 3 and not _is_boilerplate(l)][:6]

    m_concerns = re.search(r'4\.\s*OUTSTANDING\s*ISSUES(.*?)(?=5\.\s*KEY\s*MILESTONES)', full_doc_text, re.DOTALL | re.IGNORECASE)
    if m_concerns:
        lines = [l.strip("- ").strip() for l in m_concerns.group(1).split("\n") if l.strip("- ").strip()]
        concerns = [l for l in lines if len(l) > 3 and not _is_boilerplate(l)]

    highlights_summary = "; ".join(this_period_act) if this_period_act else "Engineering and procurement activities proceeding as planned."
    concerns_summary = "; ".join(concerns) if concerns else "No critical technical showstoppers reported."

    # 4. Extract Section 5: Key Milestones
    key_milestones = []
    m_miles = re.search(r'5\.\s*KEY\s*MILESTONES(.*?)(?=6\.\s*OVERALL\s*PROGRESS)', full_doc_text, re.DOTALL | re.IGNORECASE)
    if m_miles:
        mlines = [l.strip() for l in m_miles.group(1).split("\n") if l.strip()]
        curr_m: Dict[str, Any] = {}
        for line in mlines:
            if line.lower().startswith("note:") or line.lower().startswith("milestone 6"):
                break
            num_m = re.match(r'^(\d+)\.?$', line)
            if num_m:
                if curr_m and curr_m.get("description"):
                    key_milestones.append(curr_m)
                curr_m = {"item_no": num_m.group(1), "description": "", "planned_date": "", "actual_date": "", "remark": ""}
                continue
            if _is_boilerplate(line):
                continue
            if curr_m:
                d_m = re.search(r'\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b', line)
                if d_m:
                    parsed_d = _parse_date(d_m.group(0)) or d_m.group(0)
                    if not curr_m["planned_date"]:
                        curr_m["planned_date"] = parsed_d
                    elif not curr_m["actual_date"]:
                        curr_m["actual_date"] = parsed_d
                elif not curr_m["description"] and line.lower() not in ("#", "description", "planned", "date", "actual", "remark"):
                    curr_m["description"] = line
                elif curr_m["planned_date"] and not curr_m["remark"]:
                    curr_m["remark"] = line
        if curr_m and curr_m.get("description"):
            key_milestones.append(curr_m)

    # 5. Extract Section 6: Overall Progress (Page 5)
    overall_plan_pct = 0.0
    overall_actual_pct = 0.0
    incremental_plan_pct = 0.0
    incremental_actual_pct = 0.0
    deviation_pct = 0.0

    # 6. Locate Target Pages (starting search from page index 5 to avoid table-of-contents false matches)
    scurve_page_idx = None
    gantt_page_idx = None
    suborder_page_idx = None

    for i in range(5, len(full_text_pages)):
        txt = full_text_pages[i]
        if scurve_page_idx is None and len(re.findall(r'\b\d{1,2}/[A-Za-z]{3}/\d{2}\b', txt)) >= 10:
            scurve_page_idx = i
        if gantt_page_idx is None and "PREDECESSORS" in txt.upper() and "TASK NAME" in txt.upper() and ("MON " in txt.upper() or "TUE " in txt.upper()):
            gantt_page_idx = i
        if suborder_page_idx is None and "SUB ORDER STATUS REPORT" in txt.upper() and ("SUB-SUPPLIER" in txt.upper() or "CHK'D DATE" in txt.upper()):
            suborder_page_idx = i

    # S-Curve Parsing
    scurve_points = []
    if scurve_page_idx is not None:
        p_txt = full_text_pages[scurve_page_idx]
        dates = re.findall(r'\b\d{1,2}/[A-Za-z]{3}/\d{2}\b', p_txt)
        pct_vals = []
        for l in p_txt.split("\n"):
            m_pct = re.match(r'^(-?\d+(?:\.\d+)?)\s*%$', l.strip())
            if m_pct:
                pct_vals.append(float(m_pct.group(1)))

        n_dates = len(dates)
        if n_dates >= 10 and len(pct_vals) >= n_dates * 2:
            inc_p = pct_vals[0:n_dates]
            cum_p = pct_vals[n_dates:n_dates * 2]
            rest = pct_vals[n_dates * 2:]
            half_rest = len(rest) // 3 if len(rest) >= 3 else len(rest) // 2
            inc_a = rest[0:half_rest] if half_rest > 0 else []
            cum_a = rest[half_rest:half_rest * 2] if half_rest > 0 else []

            for w_idx, dt in enumerate(dates):
                p_cum = cum_p[w_idx] if w_idx < len(cum_p) else 0.0
                p_inc = inc_p[w_idx] if w_idx < len(inc_p) else 0.0
                norm_date = _parse_date(dt) or dt

                # Actuals only exist up to the cutoff date / current week
                is_future = bool(cutoff_date and norm_date > cutoff_date)
                a_cum = None if is_future else (cum_a[w_idx] if w_idx < len(cum_a) else None)
                a_inc = None if is_future else (inc_a[w_idx] if w_idx < len(inc_a) else None)

                scurve_points.append({
                    "package": pkg_code,
                    "report_no": report_no,
                    "cutoff_date": cutoff_date,
                    "week_no": w_idx,
                    "week_date": norm_date,
                    "plan_cum_pct": p_cum,
                    "actual_cum_pct": a_cum,
                    "plan_inc_pct": p_inc,
                    "actual_inc_pct": a_inc,
                })

            # Determine current cutoff progress from S-curve matching cutoff date
            past_or_current = [pt for pt in scurve_points if pt["actual_cum_pct"] is not None and (not cutoff_date or pt["week_date"] <= cutoff_date)]
            if past_or_current:
                current_pt = past_or_current[-1]
                overall_actual_pct = current_pt["actual_cum_pct"] or 0.0
                overall_plan_pct = current_pt["plan_cum_pct"] or 0.0
                incremental_actual_pct = current_pt["actual_inc_pct"] or 0.0
                incremental_plan_pct = current_pt["plan_inc_pct"] or 0.0
                deviation_pct = round(overall_actual_pct - overall_plan_pct, 2)

    if overall_plan_pct <= 0.0 and overall_actual_pct <= 0.0:
        m_tot = re.search(r'Total Must\s*=\s*100%\s*100%\s*([0-9.]+)%\s*([0-9.]+)%', full_doc_text, re.IGNORECASE)
        if m_tot:
            try:
                overall_plan_pct = float(m_tot.group(1))
                overall_actual_pct = float(m_tot.group(2))
                deviation_pct = round(overall_actual_pct - overall_plan_pct, 2)
            except (ValueError, TypeError):
                pass

    # Gantt Schedule Parsing
    gantt_tasks = []
    if gantt_page_idx is not None:
        p_txt = full_text_pages[gantt_page_idx]
        lines = [l.strip() for l in p_txt.split("\n") if l.strip()]
        i = 0
        while i < len(lines):
            if lines[i].isdigit() and int(lines[i]) == 1:
                break
            i += 1

        date_regex = re.compile(r'(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?\s*(\d{1,2}/\d{1,2}/\d{2,4})')
        dur_regex = re.compile(r'(\d+(?:\.\d+)?)\s*days?', re.IGNORECASE)

        curr_task: Optional[Dict[str, Any]] = None
        while i < len(lines):
            line = lines[i]
            if line.isdigit() and (curr_task is None or int(line) >= curr_task["task_id"] + 1):
                if curr_task:
                    gantt_tasks.append(curr_task)
                curr_task = {
                    "package": pkg_code,
                    "report_no": report_no,
                    "cutoff_date": cutoff_date,
                    "task_id": int(line),
                    "name_parts": [],
                    "duration": "",
                    "start_date": "",
                    "finish_date": "",
                    "predecessors": "",
                    "is_milestone": "No",
                    "phase_category": "Execution"
                }
                i += 1
                continue

            if curr_task:
                dur_m = dur_regex.search(line)
                if not dur_m and re.match(r'^\d+(\.\d+)?$', line) and i + 1 < len(lines) and "day" in lines[i + 1].lower():
                    line = f"{line} {lines[i + 1]}"
                    i += 1
                    dur_m = dur_regex.search(line)

                if dur_m and not curr_task["duration"]:
                    curr_task["duration"] = dur_m.group(0)
                    if "0 day" in dur_m.group(0).lower():
                        curr_task["is_milestone"] = "Yes"
                elif date_regex.search(line):
                    dates = date_regex.findall(line)
                    pred_m = re.search(r'\b(\d+)\b$', line)
                    if pred_m and not curr_task["predecessors"] and pred_m.group(1) != dates[0].split('/')[-1]:
                        curr_task["predecessors"] = pred_m.group(1)
                    for d in dates:
                        norm_d = _parse_date(d, mdy_first=True) or d
                        if not curr_task["start_date"]:
                            curr_task["start_date"] = norm_d
                        elif not curr_task["finish_date"]:
                            curr_task["finish_date"] = norm_d
                elif not curr_task["duration"]:
                    curr_task["name_parts"].append(line)
                else:
                    if line.isdigit() and not curr_task["predecessors"]:
                        curr_task["predecessors"] = line
            i += 1

        if curr_task:
            gantt_tasks.append(curr_task)

        for t in gantt_tasks:
            t["task_name"] = " ".join(t["name_parts"]).strip()
            del t["name_parts"]
            name_u = t["task_name"].upper()
            if "ENGINEERING" in name_u:
                t["phase_category"] = "Engineering"
            elif "PROCUREMENT" in name_u or "EQUIPMENT" in name_u or "VALVE" in name_u or "PUMP" in name_u:
                t["phase_category"] = "Procurement"
            elif "FABRICATION" in name_u or "SKID" in name_u or "WELD" in name_u or "ASSEMBLY" in name_u:
                t["phase_category"] = "Fabrication"
            elif "FAT" in name_u or "TEST" in name_u or "INSPECTION" in name_u:
                t["phase_category"] = "Testing"
            elif any(kw in name_u for kw in ("DELIVERY", "SHIPMENT", "SEA FREIGHT", "TRANSIT", "CIF", "WORKSITE")):
                t["phase_category"] = "Delivery"
            elif t["is_milestone"] == "Yes":
                t["phase_category"] = "Milestone"

    # Sub-Orders Parsing
    sub_orders = []
    if suborder_page_idx is not None:
        p_txt = full_text_pages[suborder_page_idx]
        lines = [l.strip() for l in p_txt.split("\n") if l.strip()]
        start_idx = 0
        for idx, l in enumerate(lines):
            if "备注" in l or "ACTUAL" in l.upper():
                start_idx = idx + 1

        items = lines[start_idx:]
        known = [
            "Pump", "Calibration Pot", "Solenoid Valve", "Strainer", "Dampener",
            "Pressure Regulator", "Pressure Indicator", "Flow Switch", "Junction Box", "Pressure Relief Valv"
        ]

        curr_so: Optional[Dict[str, Any]] = None
        date_pattern = re.compile(r'\b\d{1,2}[-/][A-Za-z0-9]{1,3}[-/]\d{2,4}\b')

        for l in items:
            matched_name = None
            for k in known:
                if l.lower() == k.lower() or l.lower().startswith(k.lower()):
                    matched_name = k
                    break
            if matched_name:
                if curr_so:
                    sub_orders.append(curr_so)
                curr_so = {
                    "package": pkg_code,
                    "report_no": report_no,
                    "cutoff_date": cutoff_date,
                    "sub_component": matched_name.replace("Valv", "Valve"),
                    "material_description": "",
                    "po_date": "",
                    "sub_vendor": "",
                    "po_status": "Ordered / In Progress",
                    "req_delivery": "",
                    "est_delivery": "",
                    "actual_delivery": "",
                    "remarks": ""
                }
                continue

            if curr_so:
                if l.lower() in ["milton roy", "kenco", "versa", "barton firtop", "fox", "fisher", "wika", "fci", "crouse-hinds", "swagelok"]:
                    curr_so["sub_vendor"] = l.title()
                elif l.lower() in ["available", "ordered", "open", "in progress"]:
                    curr_so["po_status"] = l.title()
                elif date_pattern.search(l):
                    raw_d = date_pattern.search(l).group(0)
                    norm_d = _parse_date(raw_d) or raw_d
                    if not curr_so["po_date"]:
                        curr_so["po_date"] = norm_d
                    elif not curr_so["req_delivery"]:
                        curr_so["req_delivery"] = norm_d
                    elif not curr_so["est_delivery"]:
                        curr_so["est_delivery"] = norm_d
                elif not curr_so["material_description"] and len(l) > 2 and "SUB ORDER" not in l.upper():
                    curr_so["material_description"] = l

        if curr_so:
            sub_orders.append(curr_so)

    status_tone = "success" if deviation_pct >= 0 else ("danger" if deviation_pct <= -5.0 else "warning")

    weekly_summary = {
        "package": pkg_code,
        "report_no": report_no,
        "cutoff_date": cutoff_date,
        "po_number": po_no,
        "vendor_name": vendor_name,
        "target_delivery": target_delivery,
        "overall_plan_pct": overall_plan_pct,
        "overall_actual_pct": overall_actual_pct,
        "incremental_plan_pct": incremental_plan_pct,
        "incremental_actual_pct": incremental_actual_pct,
        "deviation_pct": deviation_pct,
        "status_tone": status_tone,
        "executive_highlights": highlights_summary,
        "areas_of_concern": concerns_summary,
        "critical_path_notes": f"Gate: {key_milestones[2]['description'] if len(key_milestones) > 2 else 'Document Approval'} / Delivery {target_delivery}",
    }

    return {
        "summary": weekly_summary,
        "milestones": key_milestones,
        "scurve": scurve_points,
        "gantt_tasks": gantt_tasks,
        "sub_orders": sub_orders,
    }


def extract_mom_pdf(source: Union[str, bytes], filename: str = "") -> Dict[str, Any]:
    """Parse a MOM PDF and extract:
    - Metadata (Package, MOM_No, Meeting_Date, Subject, PO, Vendor)
    - Action Items across all sections (Item_No, Description, Action_By, Target_Date, Status)
    - Section 7.1 Major Components Procurement Status
    """
    if isinstance(source, bytes):
        doc = pymupdf.open(stream=source, filetype="pdf")
    else:
        doc = pymupdf.open(source)
        if not filename:
            filename = os.path.basename(source)

    pkg_match = re.search(r'MM-(?:ASK-1A|ZTK-1F)-GEN\w*-([A-Z0-9]+)-', filename, re.IGNORECASE)
    if not pkg_match:
        pkg_match = re.search(r'MM-[A-Z0-9]+-[A-Z0-9]+-(?:GEN\w*-)?([A-Z0-9]+)-', filename, re.IGNORECASE)
    pkg_code = pkg_match.group(1).upper() if pkg_match else ""

    mom_match = re.search(r'MOM[-_]?(\d+)', filename, re.IGNORECASE)
    mom_no = f"MOM-{int(mom_match.group(1)):04d}" if mom_match else ""

    full_text_pages = [page.get_text() for page in doc]
    doc.close()
    full_doc_text = "\n".join(full_text_pages)

    if not pkg_code:
        for code in ["CIP", "PLC", "WHCP", "IGS", "ABV", "BDV", "ESD", "PIG", "TEG", "FLARE"]:
            if code in full_doc_text.upper():
                pkg_code = code
                break
        if not pkg_code:
            pkg_code = "GEN"

    if not mom_no:
        mom_doc_m = re.search(r'MOM[-_]?(\d+)', full_doc_text, re.IGNORECASE)
        if mom_doc_m:
            mom_no = f"MOM-{int(mom_doc_m.group(1)):04d}"
        else:
            mom_no = "MOM-0001"

    meeting_date = ""
    date_m = re.search(r'Date\s*:\s*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})', full_doc_text)
    if date_m:
        meeting_date = _parse_date(date_m.group(1)) or date_m.group(1)

    # Extract Action Items from table (Pages 3 to end)
    action_items = []
    item_pattern = re.compile(r'^(\d+(?:\.\d+)?)\.?$')

    for page_idx, page_txt in enumerate(full_text_pages[2:], start=3):
        lines = [l.strip() for l in page_txt.split("\n") if l.strip()]
        curr_item: Optional[Dict[str, Any]] = None

        li = 0
        while li < len(lines):
            line = lines[li]
            m_it = item_pattern.match(line)
            if m_it:
                if curr_item and curr_item.get("desc_parts"):
                    action_items.append(curr_item)
                curr_item = {
                    "package": pkg_code,
                    "mom_no": mom_no,
                    "meeting_date": meeting_date,
                    "item_no": m_it.group(1),
                    "desc_parts": [],
                    "action_by": "",
                    "target_date": "",
                    "status": "OPEN"
                }
                li += 1
                continue

            if curr_item:
                u_line = line.upper()
                if u_line in ["CLOSED", "OPEN", "INFO", "INFO.", "PENDING"]:
                    curr_item["status"] = u_line.rstrip(".")
                elif u_line in ["MOG", "TJE", "GCME", "CLIENT", "CONTRACTOR", "VENDOR"]:
                    curr_item["action_by"] = line
                elif re.match(r'^\d{1,2}-[A-Za-z]{3}-\d{2,4}$', line) or re.match(r'^\d{1,2}/\d{1,2}/\d{2,4}$', line):
                    curr_item["target_date"] = _parse_date(line) or line
                elif "ITEM" not in u_line and "DESCRIPTION" not in u_line and "ACTION BY" not in u_line and "PAGE " not in u_line and "MM-ASK" not in u_line:
                    curr_item["desc_parts"].append(line)

            li += 1

        if curr_item and curr_item.get("desc_parts"):
            action_items.append(curr_item)

    for ai in action_items:
        ai["description"] = " ".join(ai.get("desc_parts", [])).strip()
        if "desc_parts" in ai:
            del ai["desc_parts"]

    # Extract Section 7.1 Major Components Procurement Status table if present
    procurement_status = []
    m_p_sec = re.search(r'7\.1\s*(.*?)7\.2', full_doc_text, re.DOTALL | re.IGNORECASE)
    if m_p_sec:
        sec_lines = [l.strip() for l in m_p_sec.group(1).split("\n") if l.strip()]
        curr_p: Optional[Dict[str, Any]] = None
        for sline in sec_lines:
            if sline.isdigit() and int(sline) in range(1, 20):
                if curr_p:
                    procurement_status.append(curr_p)
                curr_p = {
                    "package": pkg_code,
                    "mom_no": mom_no,
                    "meeting_date": meeting_date,
                    "item_no": int(sline),
                    "component": "",
                    "plan_date": "",
                    "actual_date": "",
                    "status": "In Progress"
                }
                continue
            if curr_p:
                d_m = re.search(r'\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b', sline)
                if d_m:
                    parsed_d = _parse_date(d_m.group(0)) or d_m.group(0)
                    if not curr_p["plan_date"]:
                        curr_p["plan_date"] = parsed_d
                    elif not curr_p["actual_date"]:
                        curr_p["actual_date"] = parsed_d
                elif not curr_p["component"] and len(sline) > 2 and sline.lower() != "no":
                    curr_p["component"] = sline
        if curr_p:
            procurement_status.append(curr_p)

    return {
        "metadata": {
            "package": pkg_code,
            "mom_no": mom_no,
            "meeting_date": meeting_date,
        },
        "action_items": action_items,
        "procurement_status": procurement_status,
    }
