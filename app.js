/* ==========================================================================
   ASK & Z1F Weekly Project Management & Vendor Package Monitoring Dashboard
   Frontend Application Controller (app.js) — Updated for Feedback Items 1–7
   ========================================================================== */

const APP_STATE = {
    project: "ASK",             // 'ASK' | 'Z1F' (Separated & Switchable, starts on ASK)
    mode: "live",               // 'live' (Strict WPR #1 > SCD/PSR/VDRL #2) | 'sample' (PLC WPR Dev Sample)
    activeTab: "package",       // 'package' | 'portfolio' | 'doccontrol' | 'lookahead'
    selectedPkgCode: "CIP",     // Starts on CIP (Live WPR-0010 Cut-off 2026-10-27 + MOM-0010)
    scurveView: "package",      // 'package' (Selected Pkg S-Curve) | 'project' (Weighted All Pkgs S-Curve)
    showWeightsTable: true,
    showIncorporatedDetails: true,
    lastUploadBannerMsg: "",
    execScope: "core",          // 'core' (31 Core Packages) | 'all' (All 108 Master Plan / PSR Equipment)
    execMitigationFilter: "ACTIONABLE", // 'ACTIONABLE' | 'ALREADY_DELAYED' | 'HIGH_RISK' | 'CRITICAL_PATH' | 'LOW_RISK' | 'ALL'
    execMitigationSearch: "",
    showMitigationMatrix: true,
    lookaheadWindow: 7,         // 7 | 14
    theme: "dark",              // 'dark' | 'light' (Daylight Mode)
    snapshot: null,
    sortState: {
        scurveWeights: { col: "package_weight_pct", dir: "desc" },
        att3Phases: { col: "phase_id", dir: "asc" },
        att4Materials: { col: "item_no", dir: "asc" },
        keyMilestones: { col: "no", dir: "asc" },
        milestoneProgress: { col: "id", dir: "asc" },
        pkgVdrlDocs: { col: "plan_date", dir: "asc" },
        pkgClientDocs: { col: "submittal_id", dir: "desc" },
        portfolio: { col: "package_code", dir: "asc" },
        hubClientPkg: { col: "total_submittals", dir: "desc" },
        hubSubmittals: { col: "submittal_id", dir: "desc" },
        lookahead: { col: "start_date", dir: "asc" },
    },
};

function esc(val) {
    if (val === null || val === undefined) return "";
    return String(val)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function fmtBlank(val, placeholder = "—") {
    if (val === null || val === undefined) return `<span class="blank-val">${placeholder}</span>`;
    const s = String(val).trim();
    if (!s || s === "—" || s.toUpperCase() === "NONE" || s.toUpperCase() === "NAN") {
        return `<span class="blank-val">${placeholder}</span>`;
    }
    return esc(s);
}

function badgeClassForTone(tone) {
    switch (tone) {
        case "success": return "badge-success";
        case "warning": return "badge-warning";
        case "danger": return "badge-danger";
        case "info": return "badge-info";
        case "purple": return "badge-purple";
        default: return "badge-muted";
    }
}

/* ==========================================================================
   COMMENT #1: UNIVERSAL ASCENDING / DESCENDING TABLE SORTING ENGINE
   ========================================================================== */
function compareValues(a, b, dir = "asc") {
    const mult = dir === "asc" ? 1 : -1;
    if (a === null || a === undefined || a === "" || a === "—") return 1;
    if (b === null || b === undefined || b === "" || b === "—") return -1;

    const numA = typeof a === "number" ? a : Number(String(a).replace(/[%+,]/g, "").trim());
    const numB = typeof b === "number" ? b : Number(String(b).replace(/[%+,]/g, "").trim());
    const isIsoDate = (v) => /^\d{4}-\d{2}-\d{2}/.test(String(v).trim());

    if (!isIsoDate(a) && !isIsoDate(b) && !Number.isNaN(numA) && !Number.isNaN(numB) && String(a).trim() !== "" && String(b).trim() !== "") {
        if (numA === numB) return 0;
        return (numA < numB ? -1 : 1) * mult;
    }

    const strA = String(a).trim().toUpperCase();
    const strB = String(b).trim().toUpperCase();
    return strA.localeCompare(strB, undefined, { numeric: true, sensitivity: "base" }) * mult;
}

function sortRows(rows, tableKey, extractorMap = {}) {
    const st = APP_STATE.sortState[tableKey];
    if (!st || !st.col) return [...rows];
    const col = st.col;
    const dir = st.dir || "asc";
    const getter = extractorMap[col] || ((row) => row[col]);
    return [...rows].sort((r1, r2) => compareValues(getter(r1), getter(r2), dir));
}

function updateSortHeadersUI(tableKey) {
    const st = APP_STATE.sortState[tableKey] || {};
    const prefix = `sort_${tableKey}_`;
    document.querySelectorAll(`[id^="${prefix}"]`).forEach(span => {
        const col = span.id.slice(prefix.length);
        const th = span.closest("th");
        if (col === st.col) {
            span.textContent = st.dir === "asc" ? "▲" : "▼";
            if (th) th.classList.add("sort-active");
        } else {
            span.textContent = "⇅";
            if (th) th.classList.remove("sort-active");
        }
    });
}

function toggleSort(tableKey, colKey) {
    const cur = APP_STATE.sortState[tableKey] || { col: colKey, dir: "asc" };
    if (cur.col === colKey) {
        cur.dir = cur.dir === "asc" ? "desc" : "asc";
    } else {
        cur.col = colKey;
        cur.dir = "asc";
    }
    APP_STATE.sortState[tableKey] = cur;
    updateSortHeadersUI(tableKey);

    switch (tableKey) {
        case "scurveWeights":
            renderScurveWeightsTable();
            break;
        case "att3Phases":
        case "att4Materials":
            renderAttachments3And4();
            break;
        case "keyMilestones":
            renderKeyMilestonesTable();
            break;
        case "milestoneProgress":
            renderMilestoneProgressTable();
            break;
        case "pkgVdrlDocs":
        case "pkgClientDocs":
            renderPackageDocuments();
            break;
        case "portfolio":
            renderPortfolioTable();
            break;
        case "hubClientPkg":
            renderDocControlHub();
            break;
        case "hubSubmittals":
            renderHubSubmittals();
            break;
        case "lookahead":
            renderLookAheadTable();
            break;
    }
}

/* ==========================================================================
   COMMENT #7: DAYLIGHT MODE / DARK MODE TOGGLE
   ========================================================================== */
function initTheme() {
    try {
        const saved = localStorage.getItem("vendor_dashboard_theme");
        if (saved === "light" || saved === "dark") {
            APP_STATE.theme = saved;
        }
    } catch (_) {}
    applyTheme();
}

function toggleTheme() {
    APP_STATE.theme = APP_STATE.theme === "dark" ? "light" : "dark";
    try {
        localStorage.setItem("vendor_dashboard_theme", APP_STATE.theme);
    } catch (_) {}
    applyTheme();
    if (APP_STATE.snapshot) {
        renderSelectedPackageWorkspace();
    }
}

function applyTheme() {
    document.documentElement.setAttribute("data-theme", APP_STATE.theme);
    const btn = document.getElementById("btnThemeToggle");
    if (btn) {
        btn.innerHTML = APP_STATE.theme === "light"
            ? "🌙 Dark Mode"
            : "☀️ Daylight Mode";
    }
}

/* ==========================================================================
   INITIALIZATION & DATA FETCHING
   ========================================================================== */
document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    setupDragAndDrop();
    refreshDashboard(false);
});

function getSelectedPackage() {
    const pkgs = (APP_STATE.snapshot && APP_STATE.snapshot.packages) || [];
    return pkgs.find(p => p.package_code === APP_STATE.selectedPkgCode) || pkgs[0] || null;
}

/* ==========================================================================
   CLIENT-SIDE NORMALIZER & FALLBACK CALCULATOR
   Ensures Key Milestone #11 ("Receipt at Worksite") = ETA and
   ROS Date from "Attachment 2 - Procurement Plan - ASK" are ALWAYS populated
   for PLC and ALL packages, even if server.py has not been restarted yet.
   ========================================================================== */
const ASK_PSR_ROS_MAP = {
    PLC:  { ros: "2027-04-28", psr_eta: "2027-04-23", psr_plan_rec: "2027-04-15", weight: 5.58 },
    ABV:  { ros: "2027-04-08", psr_eta: "2027-03-24", psr_plan_rec: "2027-05-12", weight: 3.72 },
    ACBV: { ros: "2027-04-08", psr_eta: "2027-03-24", psr_plan_rec: "2027-05-14", weight: 3.72 },
    ACV:  { ros: "2027-04-12", psr_eta: "2027-04-08", psr_plan_rec: "2027-05-27", weight: 3.72 },
    CV:   { ros: "2027-04-08", psr_eta: "2027-03-29", psr_plan_rec: "2027-03-18", weight: 3.72 },
    SRV:  { ros: "2027-04-08", psr_eta: "2027-03-29", psr_plan_rec: "2027-04-01", weight: 3.72 },
    WHCP: { ros: "2027-03-22", psr_eta: "2027-03-20", psr_plan_rec: "2027-03-24", weight: 5.58 },
    PDC:  { ros: "2027-06-02", psr_eta: "2027-05-26", psr_plan_rec: "2027-06-02", weight: 5.58 },
    CIP:  { ros: "2027-03-05", psr_eta: "2027-03-30", psr_plan_rec: "2027-03-05", weight: 5.58 },
    CDP:  { ros: "2027-03-28", psr_eta: "2027-02-16", psr_plan_rec: "2027-03-09", weight: 3.72 },
    DTP:  { ros: "2027-04-06", psr_eta: "2027-02-16", psr_plan_rec: "2027-04-02", weight: 3.72 },
    OPDP: { ros: "2027-04-17", psr_eta: "2027-02-16", psr_plan_rec: "2027-03-30", weight: 3.72 },
    PV:   { ros: "2027-03-02", psr_eta: "2027-02-15", psr_plan_rec: "2027-02-22", weight: 4.96 },
    TSP:  { ros: "2027-03-02", psr_eta: "2027-02-15", psr_plan_rec: "2027-02-22", weight: 4.96 },
    PSP:  { ros: "2027-03-02", psr_eta: "2027-02-15", psr_plan_rec: "2027-02-22", weight: 4.96 },
    SKT:  { ros: "2027-03-23", psr_eta: "2027-02-28", psr_plan_rec: "2027-03-01", weight: 4.34 },
    CDS:  { ros: "2027-02-12", psr_eta: "2027-01-19", psr_plan_rec: "2027-01-26", weight: 3.72 },
    DGS:  { ros: "2027-03-24", psr_eta: "2027-03-10", psr_plan_rec: "2027-03-13", weight: 3.72 },
    DOL:  { ros: "2027-03-23", psr_eta: "2027-03-15", psr_plan_rec: "2027-03-18", weight: 4.34 },
    IGF:  { ros: "2027-03-23", psr_eta: "2027-03-15", psr_plan_rec: "2027-03-18", weight: 4.34 },
    IGFIL:{ ros: "2027-02-26", psr_eta: "2027-01-30", psr_plan_rec: "2027-02-15", weight: 3.10 },
    TANK: { ros: "2027-03-24", psr_eta: "2027-02-26", psr_plan_rec: "2027-02-04", weight: 3.10 },
    SOLR: { ros: "2027-06-23", psr_eta: "2027-03-19", psr_plan_rec: "2027-04-08", weight: 4.34 },
    BV:   { ros: "2027-01-28", psr_eta: "2027-01-25", psr_plan_rec: "2027-02-09", weight: 1.86 },
    GLV:  { ros: "2027-01-28", psr_eta: "2027-01-11", psr_plan_rec: "2027-02-16", weight: 1.86 },
    PLS:  { ros: "2027-04-07", psr_eta: "2027-02-24", psr_plan_rec: "2027-02-24", weight: 1.24 },
    PSV:  { ros: "2027-04-27", psr_eta: "2027-04-21", psr_plan_rec: "2027-04-21", weight: 1.86 },
    PG:   { ros: "2027-03-23", psr_eta: "2027-02-23", psr_plan_rec: "2027-02-24", weight: 1.24 },
    PLT:  { ros: "2026-10-12", psr_eta: "2026-09-28", psr_plan_rec: "2026-09-28", weight: 1.86 },
    BEAM: { ros: "2026-10-20", psr_eta: "2026-10-06", psr_plan_rec: "2026-10-06", weight: 1.86 },
    SMT:  { ros: "2026-10-20", psr_eta: "2026-10-06", psr_plan_rec: "2026-10-06", weight: 1.86 },
};

function dateDiffDays(d1Str, d2Str) {
    if (!d1Str || !d2Str) return null;
    const t1 = Date.parse(String(d1Str).slice(0, 10));
    const t2 = Date.parse(String(d2Str).slice(0, 10));
    if (Number.isNaN(t1) || Number.isNaN(t2)) return null;
    return Math.round((t2 - t1) / 86400000);
}

function normalizeSnapshotAndPackages(data) {
    if (!data || !Array.isArray(data.packages)) return data;
    const pkgs = data.packages;
    const nPkgs = pkgs.length || 1;

    pkgs.forEach((pkg, idx) => {
        const code = pkg.package_code;
        const psrRef = ASK_PSR_ROS_MAP[code] || {};
        const cutOff = pkg.cut_off_date || "2026-10-16";
        const ms = pkg.key_milestones || [];

        // 1. Ensure ROS Date from Attachment 2 Procurement Plan is always populated for all packages
        if (!pkg.ros_date && psrRef.ros) {
            pkg.ros_date = psrRef.ros;
        }
        const rosDate = pkg.ros_date || psrRef.ros || "";

        // 2. Track ETA with Gantt Chart Priority #1 (if available) > Key Milestone Receipt at Worksite > PSR
        const m11 = ms.find(m => m.key === "receipt_worksite" || String(m.description || "").toUpperCase().includes("RECEIPT AT WORKSITE") || m.no === 11) || ms[10] || {};
        let etaDate = "";
        let deliverySource = "Key Milestone (Receipt at Worksite)";
        if (pkg.has_live_wpr && pkg.gantt_eta_date) {
            etaDate = pkg.gantt_eta_date;
            deliverySource = (pkg.package_highlights && pkg.package_highlights.delivery_source) || `WPR Gantt Chart (${pkg.gantt_eta_task_name || "CIF Arrival"})`;
        } else if (pkg.has_live_wpr) {
            etaDate = m11.actual_date || m11.planned_date || m11.forecast_date || pkg.eta_date || psrRef.psr_eta || "";
            deliverySource = "WPR Receipt at Worksite (ETA)";
            if (m11 && m11.planned_date && m11.forecast_date && m11.forecast_date !== m11.planned_date) {
                m11.deviation_note = `PSR Fcst: ${m11.forecast_date}`;
                m11.forecast_date = "";
            }
        } else {
            etaDate = m11.actual_date || m11.forecast_date || pkg.eta_date || psrRef.psr_eta || m11.planned_date || "";
            deliverySource = "SCD / PSR Receipt at Worksite (ETA)";
        }
        pkg.eta_date = etaDate;

        // 3. Ensure package_weight_pct
        if (!pkg.package_weight_pct) {
            pkg.package_weight_pct = psrRef.weight || Number((100 / nPkgs).toFixed(2));
        }

        // 4. Compute / Enrich package_highlights so 4.1A, 4.1B, 4.1C, and 4.1D (#6) are NEVER blank
        const completedMs = ms.filter(m => m.actual_date);
        const pendingMs = ms.filter(m => !m.actual_date && (m.planned_date || m.forecast_date));
        const latestCompleted = completedMs.length ? completedMs[completedMs.length - 1] : null;
        const nextMs = pendingMs.length ? pendingMs[0] : (ms.length ? ms[ms.length - 1] : null);
        const nextTargetDt = nextMs ? (nextMs.forecast_date || nextMs.planned_date || "") : "";
        const remDaysNext = dateDiffDays(cutOff, nextTargetDt);
        const daysToDelivery = dateDiffDays(cutOff, etaDate);
        const rosDelayDays = (rosDate && etaDate) ? dateDiffDays(rosDate, etaDate) : null;
        const isRosDelayed = Boolean(rosDelayDays !== null && rosDelayDays > 0);
        const rosMarginDays = rosDelayDays !== null ? -rosDelayDays : null;

        let rosStatusLabel = "ROS Date / ETA Pending";
        let rosTone = "muted";
        if (rosDelayDays !== null) {
            if (rosDelayDays > 0) {
                rosStatusLabel = `🚨 DELAYED: ETA is +${rosDelayDays}d AFTER ROS Date`;
                rosTone = "danger";
            } else if (rosDelayDays === 0) {
                rosStatusLabel = "⚠️ Critical: ETA Equals ROS Date (0d Buffer)";
                rosTone = "warning";
            } else if (Math.abs(rosDelayDays) <= 14) {
                rosStatusLabel = `⚠️ Tight Buffer: ${Math.abs(rosDelayDays)}d Before ROS Date`;
                rosTone = "warning";
            } else {
                rosStatusLabel = `✓ On Time: ${Math.abs(rosDelayDays)}d Buffer Before ROS`;
                rosTone = "success";
            }
        }

        pkg.package_highlights = {
            reference_date: cutOff,
            completed_milestones_count: completedMs.length,
            total_milestones_count: ms.length,
            current_stage_summary: latestCompleted
                ? `Completed #${latestCompleted.no} ${latestCompleted.description} (${latestCompleted.actual_date})`
                : "Awaiting Milestone #1 (Contract Award / KOM)",
            active_stage_summary: nextMs
                ? `Active: #${nextMs.no} ${nextMs.description}`
                : "All Key Milestones Completed",
            next_milestone_no: nextMs ? nextMs.no : null,
            next_milestone_name: nextMs ? `#${nextMs.no} ${nextMs.description}` : "—",
            next_milestone_date: nextTargetDt,
            remaining_days_to_next_milestone: remDaysNext,
            effective_delivery_date: etaDate,
            delivery_source: deliverySource,
            days_to_delivery: daysToDelivery,
            ros_date: rosDate,
            is_ros_delayed: isRosDelayed,
            ros_delay_days: isRosDelayed ? rosDelayDays : 0,
            ros_margin_days: rosMarginDays,
            ros_status_label: rosStatusLabel,
            ros_tone: rosTone,
            ...(pkg.package_highlights || {}),
        };

        // 5. Fallback dynamic progress & S-curve ONLY in Dev Sample mode if running against an older server
        if (data.mode === "sample" && !pkg.has_live_wpr && pkg.overall_progress && pkg.overall_progress.planned_pct === 12.0 && pkg.overall_progress.actual_pct === 8.5 && code !== "PLC") {
            const vdrlPct = Number((pkg.vdrl_status || {}).pct_submitted || 0);
            const hasKom = Boolean((pkg.manufacturer_detail || {}).available || (ms[1] && ms[1].actual_date));
            const dynPlan = Number(Math.min(28, 6.5 + (idx % 7) * 1.8 + (hasKom ? 4.0 : 0)).toFixed(1));
            const dynAct = Number(Math.min(25, (hasKom ? 4.0 : 1.5) + (vdrlPct * 0.12)).toFixed(1));
            pkg.overall_progress.planned_pct = dynPlan;
            pkg.overall_progress.actual_pct = dynAct;
            pkg.overall_progress.variance_pct = Number((dynAct - dynPlan).toFixed(1));
            pkg.overall_progress.this_period_activities = [
                `[${code}] Package engineering & VDRL tracking active (${(pkg.vdrl_status || {}).submitted || 0}/${(pkg.vdrl_status || {}).total_docs || 0} docs submitted, ${vdrlPct.toFixed(1)}%)`,
                `[${code}] Receipt at Worksite (ETA): ${etaDate || "—"} vs Fixed PSR ROS Date: ${rosDate || "—"} (${rosStatusLabel})`,
            ];
            if (pkg.scurve && Array.isArray(pkg.scurve.points)) {
                pkg.scurve.planned_to_date_pct = dynPlan;
                pkg.scurve.actual_to_date_pct = dynAct;
                pkg.scurve.variance_pct = Number((dynAct - dynPlan).toFixed(1));
                pkg.scurve.points = pkg.scurve.points.map((pt, wi) => ({
                    ...pt,
                    planned_pct: Number(Math.min(100, pt.planned_pct * (0.88 + (idx % 5) * 0.04)).toFixed(2)),
                    actual_pct: pt.actual_pct !== null && pt.actual_pct !== undefined
                        ? Number(((wi + 1) / 4 * dynAct).toFixed(2))
                        : null,
                }));
            }
        }

        // 6. Fallback Attachment 3 & Attachment 4 summaries if not yet in server response
        if (!pkg.attachment3_summary || !pkg.attachment4_summary) {
            const m4 = ms[3] || {};
            const m5 = ms[4] || {};
            const m7 = ms[6] || {};
            const appDt = m4.forecast_date || m4.planned_date || "2026-11-16";
            const matDt = m5.forecast_date || m5.planned_date || "2027-01-18";
            const fatDt = m7.forecast_date || m7.planned_date || "2027-02-23";
            const vdrlOverdue = (pkg.vdrl_status || {}).key_docs_overdue || 0;
            pkg.attachment3_summary = {
                package_code: code,
                total_phases: 5,
                total_tasks: (pkg.gantt_tasks || []).length || 61,
                critical_gate: `Key Document Approval (${appDt}) → Sub-Vendor Material Release (${matDt})`,
                situation_summary: `Critical Path Analysis (${code}): Active gate is Key Document Approval (${appDt}) driving Material Delivery (${matDt}). Receipt at Worksite (ETA): ${etaDate || "—"} vs Fixed PSR ROS Date: ${rosDate || "—"} (${rosStatusLabel}).`,
                phases: [
                    { phase_id: "Phase 1", phase_name: "Project Initiation, PO & Kick-Off Meeting", duration_days: 14, start_date: (ms[0] || {}).planned_date || "2026-09-18", finish_date: (ms[2] || {}).planned_date || "2026-10-08", status: "Completed", tone: "success", critical_note: "PO & KOM baseline established" },
                    { phase_id: "Phase 2A", phase_name: "Detailed Engineering, GAD & Key Document Approval", duration_days: 40, start_date: (ms[1] || {}).planned_date || "2026-09-25", finish_date: appDt, status: vdrlOverdue > 0 ? "Delayed / Overdue Docs" : "In Progress (Active Gate)", tone: vdrlOverdue > 0 ? "danger" : "info", critical_note: `${vdrlOverdue} Key VDRL doc(s) overdue; gates sub-vendor PO placement` },
                    { phase_id: "Phase 2B", phase_name: "Sub-Vendor Materials & Long-Lead Procurement", duration_days: 60, start_date: appDt, finish_date: matDt, status: "Hold on Key Doc Approval", tone: "warning", critical_note: "Sub-orders gated by GAD, BOM & Datasheet approval" },
                    { phase_id: "Phase 2C", phase_name: "Shop Fabrication, Assembly & Internal Testing", duration_days: 40, start_date: (ms[5] || {}).planned_date || appDt, finish_date: fatDt, status: "Scheduled", tone: "info", critical_note: `Shop assembly prior to FAT (${fatDt})` },
                    { phase_id: "Phase 2D", phase_name: "FAT, Packing & Receipt at Worksite (ETA)", duration_days: 30, start_date: fatDt, finish_date: etaDate, status: isRosDelayed ? "ROS Breach Risk!" : "Scheduled", tone: isRosDelayed ? "danger" : "success", critical_note: `Receipt at Worksite (ETA: ${etaDate || "—"}) vs Fixed PSR ROS (${rosDate || "—"})` },
                ],
            };
            pkg.attachment4_summary = {
                package_code: code,
                total_components: 4,
                components_on_hold: 3,
                target_material_delivery_date: matDt,
                expediting_narrative: `Attachment 4 Material Expediting Summary (${code}): 3 of 4 major sub-vendor material packages are held pending Key Document & BOM approval. Sub-orders must be released by ${appDt} to meet Material Delivery (${matDt}) and Receipt at Worksite ETA (${etaDate || "—"}) before Fixed ROS (${rosDate || "—"}).`,
                items: [
                    { item_no: "6.1", component: `${pkg.package_name} — Primary Enclosure / Skid / Body`, sub_vendor_origin: "Approved Sub-Vendor / Mill", prerequisite_gate: "Waiting GAD / Datasheet Approval", po_status: "Hold — Pending Key Doc Approval", lead_time_days: 60, planned_order_date: appDt, planned_eta: matDt, risk_level: vdrlOverdue > 0 ? "High Risk" : "Medium / Watch", tone: vdrlOverdue > 0 ? "danger" : "warning", impact_on_production: "Critical Path (Blocks Shop Assembly Start)" },
                    { item_no: "6.2", component: `${pkg.package_name} — Actuation / Trim / Breakers`, sub_vendor_origin: "International Specialized Sub-Vendor", prerequisite_gate: "Waiting BOM / Sizing Approval", po_status: "Hold — Pending BOM Approval", lead_time_days: 60, planned_order_date: appDt, planned_eta: matDt, risk_level: "Medium / Watch", tone: "warning", impact_on_production: "High Impact (Long-Lead Imported Assembly)" },
                    { item_no: "6.3", component: `${pkg.package_name} — Instrumentation & Bulk Accessories`, sub_vendor_origin: "Tier-1 Sub-Supplier", prerequisite_gate: "Sub-vendor Technical Clarification", po_status: "Hold — Pending Spec Freeze", lead_time_days: 45, planned_order_date: appDt, planned_eta: matDt, risk_level: "Medium / Watch", tone: "warning", impact_on_production: "Moderate (Required Before Integration)" },
                    { item_no: "6.4", component: `${pkg.package_name} — Core OEM Assembly`, sub_vendor_origin: "Primary OEM Factory", prerequisite_gate: "Internal Engineering Release", po_status: "Engineering Active", lead_time_days: 55, planned_order_date: appDt, planned_eta: matDt, risk_level: isRosDelayed ? "High Risk (ROS Delay)" : "On Track", tone: isRosDelayed ? "danger" : "success", impact_on_production: "Primary Core Equipment" },
                ],
            };
        }
    });

    // 7. Ensure project_overall_scurve is populated across all packages
    if (!data.project_overall_scurve) {
        let weightedPlan = 0;
        let weightedAct = 0;
        const breakdown = pkgs.map(p => {
            const w = Number(p.package_weight_pct || 0);
            const pl = Number((p.overall_progress || {}).planned_pct || 0);
            const ac = Number((p.overall_progress || {}).actual_pct || 0);
            const wPl = Number(((w * pl) / 100).toFixed(2));
            const wAc = Number(((w * ac) / 100).toFixed(2));
            weightedPlan += wPl;
            weightedAct += wAc;
            return {
                package_code: p.package_code,
                package_name: p.package_name,
                discipline: p.discipline,
                package_weight_pct: w,
                pkg_planned_pct: pl,
                pkg_actual_pct: ac,
                weighted_planned_pct: wPl,
                weighted_actual_pct: wAc,
            };
        });
        const basePts = (pkgs[0] && pkgs[0].scurve && pkgs[0].scurve.points) || [];
        data.project_overall_scurve = {
            current_week: "W4",
            cut_off_date: "2026-10-16",
            planned_to_date_pct: Number(weightedPlan.toFixed(2)),
            actual_to_date_pct: Number(weightedAct.toFixed(2)),
            variance_pct: Number((weightedAct - weightedPlan).toFixed(2)),
            calculation_basis: `Weighted cumulative progress across all ${pkgs.length} ${data.project} packages (Total Weight = 100.0%)`,
            package_weights_breakdown: breakdown,
            points: basePts,
        };
    }

    if (data.portfolio_kpis) {
        const delayedList = pkgs.filter(p => (p.package_highlights || {}).is_ros_delayed).map(p => p.package_name || p.package_code);
        data.portfolio_kpis.ros_delayed_packages_count = delayedList.length;
        data.portfolio_kpis.ros_delayed_packages_list = delayedList;
        data.portfolio_kpis.overall_planned_pct = data.project_overall_scurve.planned_to_date_pct;
        data.portfolio_kpis.overall_actual_pct = data.project_overall_scurve.actual_to_date_pct;
        data.portfolio_kpis.overall_variance_pct = data.project_overall_scurve.variance_pct;
    }

    if (data.client_doc_log_overview && Array.isArray(data.client_doc_log_overview.package_breakdown)) {
        // Point 1: Ensure PLT (Plate) is never duplicated as PLT (PLC)
        data.client_doc_log_overview.package_breakdown = data.client_doc_log_overview.package_breakdown.filter(
            r => !(r.raw_package_code === "PLT" && r.package_code === "PLC")
        );
    }

    return data;
}

async function refreshDashboard(force = false) {
    const statusEl = document.getElementById("lastGeneratedText");
    if (statusEl) statusEl.textContent = `Loading ${APP_STATE.project} snapshot...`;

    try {
        const url = `/api/dashboard?project=${encodeURIComponent(APP_STATE.project)}&mode=${encodeURIComponent(APP_STATE.mode)}${force ? "&refresh=1" : ""}&_t=${Date.now()}`;
        const resp = await fetch(url, { cache: "no-store" });
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const rawData = await resp.json();
        const data = normalizeSnapshotAndPackages(rawData);
        APP_STATE.snapshot = data;

        const pkgs = data.packages || [];
        if (pkgs.length > 0) {
            const exists = pkgs.some(p => p.package_code === APP_STATE.selectedPkgCode);
            if (!exists) {
                APP_STATE.selectedPkgCode = pkgs[0].package_code;
            }
        }

        renderAll();
    } catch (err) {
        console.error("Failed loading dashboard snapshot:", err);
        if (statusEl) statusEl.textContent = `Error loading snapshot: ${err.message}`;
    }
}

function switchProject(projectCode) {
    if (APP_STATE.project === projectCode) return;
    APP_STATE.project = projectCode;
    APP_STATE.selectedPkgCode = projectCode === "ASK" ? "CIP" : "CCV";

    document.getElementById("btnProjASK").classList.toggle("active", projectCode === "ASK");
    document.getElementById("btnProjZ1F").classList.toggle("active", projectCode === "Z1F");
    document.getElementById("brandBadge").textContent = projectCode;

    // Update Export File download button for active project
    const dlBtn = document.getElementById("btnDownloadExportFile");
    if (dlBtn) {
        dlBtn.href = `/api/download-export-file?project=${projectCode}`;
        dlBtn.setAttribute("download", `Export Data File_${projectCode}.xlsx`);
        dlBtn.innerHTML = `💾 Download Export Data File (${projectCode})`;
        dlBtn.setAttribute("title", `Download current ${projectCode} Export Data File (< 1 MB)`);
    }

    // Immediately align Look-Ahead Scope dropdown and subtitle with selected project
    const isZ1F = projectCode === "Z1F";
    const defaultDoc = isZ1F ? "MM-ZTK-1F-GEN-PJM-SCD-0002" : "MM-ASK-1A-GEN01-PJM-SCD-0001";
    const defaultRev = isZ1F ? "C1" : "B1";
    const defaultCutoff = isZ1F ? "2026-09-16" : "2026-10-16";

    const scopeSel = document.getElementById("lookAheadScopeFilter");
    if (scopeSel) {
        const optMaster = scopeSel.querySelector('option[value="MASTER"]');
        if (optMaster) {
            optMaster.textContent = `Master Schedule Only (${defaultDoc})`;
        }
    }
    const subPre = document.getElementById("lookAheadSubtitle");
    if (subPre) {
        subPre.textContent = `Cut-off: ${defaultCutoff} • Integrated from ${defaultDoc} (Rev. ${defaultRev}) & Package Schedules`;
    }

    refreshDashboard(false);
}

function switchMode(mode) {
    if (APP_STATE.mode === mode) return;
    APP_STATE.mode = mode;

    document.getElementById("btnModeSample").classList.toggle("active", mode === "sample");
    document.getElementById("btnModeLive").classList.toggle("active", mode === "live");

    refreshDashboard(false);
}

function switchTab(tabName) {
    APP_STATE.activeTab = tabName;
    document.getElementById("tabBtnPackage").classList.toggle("active", tabName === "package");
    document.getElementById("tabBtnPortfolio").classList.toggle("active", tabName === "portfolio");
    document.getElementById("tabBtnDocControl").classList.toggle("active", tabName === "doccontrol");
    const btnLookAhead = document.getElementById("tabBtnLookAhead");
    if (btnLookAhead) btnLookAhead.classList.toggle("active", tabName === "lookahead");

    document.getElementById("panelPackage").classList.toggle("active", tabName === "package");
    document.getElementById("panelPortfolio").classList.toggle("active", tabName === "portfolio");
    document.getElementById("panelDocControl").classList.toggle("active", tabName === "doccontrol");
    const panelLookAhead = document.getElementById("panelLookAhead");
    if (panelLookAhead) panelLookAhead.classList.toggle("active", tabName === "lookahead");

    // Only show quick package select dropdown on Tab 1 to prevent awkward layout wrapping
    const qf = document.getElementById("quickPackageFilterBar");
    if (qf) {
        qf.style.display = tabName === "package" ? "flex" : "none";
    }

    if (tabName === "lookahead") {
        renderLookAheadHub();
    }
}

function selectPackage(pkgCode, openWorkspace = false) {
    APP_STATE.selectedPkgCode = pkgCode;
    const selectEl = document.getElementById("quickPackageSelect");
    if (selectEl && selectEl.value !== pkgCode) {
        selectEl.value = pkgCode;
    }
    if (openWorkspace && APP_STATE.activeTab !== "package") {
        switchTab("package");
    }
    renderPackageSelectorControls();
    renderSelectedPackageWorkspace();
    renderPortfolioTable();
}

function switchScurveView(viewMode) {
    APP_STATE.scurveView = viewMode;
    document.getElementById("btnScurvePkg").classList.toggle("active", viewMode === "package");
    document.getElementById("btnScurveProject").classList.toggle("active", viewMode === "project");
    const pkg = getSelectedPackage();
    if (pkg) renderPackageSCurve(pkg);
}

function toggleWeightsBreakdownTable() {
    APP_STATE.showWeightsTable = !APP_STATE.showWeightsTable;
    const wrap = document.getElementById("scurveWeightsTableWrapper");
    if (wrap) {
        wrap.style.display = APP_STATE.showWeightsTable ? "block" : "none";
    }
}

/* ==========================================================================
   MASTER RENDER ORCHESTRATOR
   ========================================================================== */
function renderAll() {
    const snap = APP_STATE.snapshot;
    if (!snap) return;

    document.getElementById("brandBadge").textContent = snap.project;
    document.getElementById("projectSubtitle").textContent = snap.project_full_name;
    document.getElementById("lastGeneratedText").textContent =
        `Synced: ${snap.generated_at} • Mode: ${snap.mode === "sample" ? "Dev Sample (Dynamic Package + PLC Template)" : "Live Execution (WPR > PSR)"}`;

    const src = snap.sources_summary || {};
    const isZ1F = (snap.project || APP_STATE.project) === "Z1F";
    const defaultMasterFile = isZ1F ? "MM-ZTK-1F-GEN-PJM-SCD-0002.pdf" : "MM-ASK-1A-GEN01-PJM-SCD-0001.pdf";
    const defaultMasterRev = isZ1F ? "C1" : "B1";
    document.getElementById("sourcesRibbonPills").innerHTML = `
        <span class="source-pill" style="border-color:rgba(124,58,237,0.5);background:rgba(124,58,237,0.12);"><strong>1. Master Schedule (Governing ROS)</strong> ${esc(src.master_schedule_file || defaultMasterFile)} (Rev. ${esc(src.master_schedule_rev || defaultMasterRev)})</span>
        <span class="source-pill" style="border-color:rgba(16,185,129,0.5);background:rgba(16,185,129,0.12);"><strong>2. Procurement Plan (ROS Companion)</strong> ${esc(src.psr_file || (isZ1F ? "Attachment 1 - Procurement Plan-Z1F..." : "Attachment 2-Procurement Plan-ASK..."))}</span>
        <span class="source-pill"><strong>3. WPR/</strong> ${src.wpr_files_count || 0} Live File(s) (.pdf/.docx)</span>
        <span class="source-pill"><strong>4. MOM/</strong> ${src.mom_files_count || 0} Live Meeting(s) (.pdf)</span>
        <span class="source-pill"><strong>5. VDRL/</strong> ${src.vdrl_packages_count || 0} Pkg Registers (.xlsx)</span>
        <span class="source-pill"><strong>6. Client Document Log</strong> ${esc(src.client_doc_log_file || "None")} ${src.client_doc_log_updated ? `(${esc(src.client_doc_log_updated)})` : ""}</span>
        <span class="source-pill"><strong>7. KOM/</strong> ${src.kom_packages_count || 0} Pkg Manufacturer Profiles</span>
    `;

    renderIncorporatedFilesBanner();

    Object.keys(APP_STATE.sortState).forEach(k => updateSortHeadersUI(k));

    renderExecutiveKpiStrip(snap.portfolio_kpis || {});
    renderTab1ExecutiveSummary();
    renderPackageSelectorControls(false);
    renderSelectedPackageWorkspace();
    renderPortfolioTable();
    renderDocControlHub();
    renderLookAheadHub();
}

function toggleIncorporatedFilesDetails() {
    APP_STATE.showIncorporatedDetails = !APP_STATE.showIncorporatedDetails;
    renderIncorporatedFilesBanner();
}

function renderIncorporatedFilesBanner() {
    const bannerEl = document.getElementById("incorporatedFilesBanner");
    const snap = APP_STATE.snapshot;
    if (!bannerEl || !snap) return;

    const src = snap.sources_summary || {};
    const incList = Array.isArray(src.incorporated_files) ? src.incorporated_files : [];
    const wprItems = incList.filter(f => f.category === "WPR");
    const momItems = incList.filter(f => f.category === "MOM");
    const scdItem = incList.find(f => f.category === "SCD");
    const psrItem = incList.find(f => f.category === "PSR");

    const modeBadge = snap.mode === "live"
        ? `<span class="badge badge-success">✅ Execution Mode Active (Live WPR + MOM &gt; SCD + VDRL)</span>`
        : `<span class="badge badge-info">🧪 Dev Sample Mode Active</span>`;

    const uploadAlertHtml = APP_STATE.lastUploadBannerMsg
        ? `<div class="alert-item sev-lookahead" style="margin-top:4px;padding:8px 12px;">
               <div class="alert-main">
                   <div class="alert-title">🎉 Latest File Upload &amp; Sync Incorporated into Dashboard</div>
                   <div class="alert-detail">${APP_STATE.lastUploadBannerMsg}</div>
               </div>
               <button type="button" class="btn-action" style="padding:3px 8px;font-size:11px;" onclick="APP_STATE.lastUploadBannerMsg='';renderIncorporatedFilesBanner();">Dismiss</button>
           </div>`
        : "";

    const wprPills = wprItems.map(f => `
        <button type="button" class="inc-file-pill inc-wpr" onclick="selectPackage('${esc(f.package)}', true)" title="Click to open ${esc(f.package)} workspace (${esc(f.filename)})">
            <span>📊 <strong>WPR [${esc(f.package)}]</strong>: ${esc(f.report_no || "WPR")}</span>
            <span class="badge badge-success" style="font-size:10px;padding:1px 6px;">Cut-off: ${esc(f.cut_off_date || "—")}</span>
            <span style="font-size:11px;color:var(--text-secondary);">Plan ${Number(f.planned_pct || 0).toFixed(2)}% / Act ${Number(f.actual_pct || 0).toFixed(2)}% • ${esc(f.filename)}</span>
        </button>
    `).join("");

    const momPills = momItems.map(f => `
        <button type="button" class="inc-file-pill inc-mom" onclick="selectPackage('${esc(f.package)}', true)" title="Click to open ${esc(f.package)} MOM summary (${esc(f.filename)})">
            <span>📋 <strong>MOM [${esc(f.package)}]</strong>: ${esc(f.report_no || "MOM")}</span>
            <span class="badge badge-info" style="font-size:10px;padding:1px 6px;">Date: ${esc(f.cut_off_date || "—")}</span>
            <span style="font-size:11px;color:var(--text-secondary);">${esc(f.detail || "")}</span>
        </button>
    `).join("");

    const basePills = `
        ${scdItem ? `<span class="inc-file-pill inc-scd">📅 <strong>Master Schedule (SCD):</strong> ${esc(scdItem.filename)} (${esc(scdItem.report_no || "")} • Cut-off: ${esc(scdItem.cut_off_date || "—")})</span>` : ""}
        ${psrItem ? `<span class="inc-file-pill inc-scd">📦 <strong>Procurement Plan:</strong> ${esc(psrItem.filename)}</span>` : ""}
    `;

    bannerEl.innerHTML = `
        <div class="incorporated-banner-head">
            <div class="incorporated-banner-title">
                <span>🟢 Uploaded Files Incorporated &amp; Updated in Dashboard</span>
                ${modeBadge}
                <span style="font-size:11.5px;font-weight:500;color:var(--text-secondary);">
                    • Last Synced: <strong>${esc(snap.generated_at)}</strong> • ${wprItems.length} Live WPR(s), ${momItems.length} Live MOM(s), ${src.vdrl_packages_count || 0} VDRL(s) Active
                </span>
            </div>
            <div style="display:flex;gap:8px;align-items:center;">
                <span style="font-size:11px;color:var(--text-muted);">Packages without WPR/MOM show SCD &amp; VDRL data only</span>
                <button type="button" class="btn-action" style="padding:4px 10px;font-size:11px;" onclick="toggleIncorporatedFilesDetails()">
                    ${APP_STATE.showIncorporatedDetails ? "▲ Hide Incorporated Files" : "▼ Show Incorporated Files"}
                </button>
            </div>
        </div>
        ${uploadAlertHtml}
        <div class="incorporated-pills-row" style="display:${APP_STATE.showIncorporatedDetails ? "flex" : "none"};">
            ${wprPills}
            ${momPills}
            ${basePills}
        </div>
    `;
}

/* ==========================================================================
   EXECUTIVE 6-CARD SINGLE-ROW KPI STRIP (Includes Weighted S-Curve & Master Plan / ROS Delays)
   ========================================================================== */
function renderExecutiveKpiStrip(kpis) {
    const variance = Number(kpis.overall_variance_pct || 0);
    const varSign = variance > 0 ? `+${variance.toFixed(2)}%` : `${variance.toFixed(2)}%`;
    const varBadge = variance >= 0 ? "badge-success" : "badge-warning";
    const rosDelayedCount = Number(kpis.ros_delayed_packages_count ?? kpis.ros_delayed_packages ?? 0);
    const rosDelayedList = (kpis.ros_delayed_packages_list || []).join(", ");
    const highRiskCount = Number(kpis.high_risk_count || 0);

    document.getElementById("executiveKpiStrip").innerHTML = `
        <div class="kpi-card">
            <div class="kpi-label">1. Project &amp; Active Packages</div>
            <div class="kpi-value">
                <span>${esc(APP_STATE.project)} — ${kpis.total_packages || 0} Pkgs</span>
            </div>
            <div class="kpi-sub">${kpis.core_packages_count || 31} Core Pkgs • ${kpis.live_wpr_packages || 0} Live WPR • ${kpis.kom_populated_packages || 0} KOM</div>
        </div>

        <div class="kpi-card tone-emerald">
            <div class="kpi-label">2. Weighted Project S-Curve (All Pkgs)</div>
            <div class="kpi-value">
                <span>${Number(kpis.overall_actual_pct || 0).toFixed(2)}%</span>
                <span class="badge ${varBadge}">${varSign} vs ${Number(kpis.overall_planned_pct || 0).toFixed(2)}% Plan</span>
            </div>
            <div class="kpi-sub">Weighted across all ${kpis.total_packages || 0} packages (Total Weight = 100%)</div>
        </div>

        <div class="kpi-card ${rosDelayedCount > 0 ? "tone-rose" : "tone-indigo"}">
            <div class="kpi-label">3. 🎯 Procurement Plan ROS Delivery Check</div>
            <div class="kpi-value">
                <span>${rosDelayedCount} Delayed</span>
                <span class="badge ${rosDelayedCount > 0 ? "badge-danger" : "badge-success"}">${rosDelayedCount > 0 ? `ETA > ROS (${esc(rosDelayedList)})` : "All ETA ≤ ROS Date"}</span>
            </div>
            <div class="kpi-sub">Governed by Procurement Plan ROS Date • Rule: Package Delivery ETA ≤ Fixed ROS Date</div>
        </div>

        <div class="kpi-card tone-amber">
            <div class="kpi-label">4. VDRL Vendor Submission</div>
            <div class="kpi-value">
                <span>${kpis.vdrl_submitted_docs || 0} / ${kpis.vdrl_total_docs || 0}</span>
                <span class="badge badge-info">${Number(kpis.vdrl_submitted_pct || 0).toFixed(1)}%</span>
            </div>
            <div class="kpi-sub">${kpis.vdrl_overdue_docs || 0} Overdue vendor submission(s) across VDRL/</div>
        </div>

        <div class="kpi-card tone-purple">
            <div class="kpi-label">5. Client Review &amp; Return (D2 Log)</div>
            <div class="kpi-value">
                <span>${kpis.client_log_total_submittals || 0}</span>
                <span class="badge badge-purple">${kpis.client_log_in_review || 0} In Review</span>
            </div>
            <div class="kpi-sub">APPR: ${kpis.client_log_appr || 0} • AWC: ${kpis.client_log_awc || 0} • Overdue Review: ${kpis.client_log_overdue || 0}</div>
        </div>

        <div class="kpi-card tone-rose">
            <div class="kpi-label">6. Active Areas of Concern</div>
            <div class="kpi-value">
                <span>${kpis.total_areas_of_concern || 0}</span>
                <span class="badge badge-danger">Action Items</span>
            </div>
            <div class="kpi-sub">Master Plan delays, Att.3/4 holds, overdue VDRLs &amp; client reviews</div>
        </div>
    `;
}

/* ==========================================================================
   ONE-PICTURE EXECUTIVE SUMMARY (4-TIER PACKAGE HEALTH + SMART MITIGATION MATRIX)
   Tracked Against Project Master Plan (MM-ASK-1A-GEN01-PJM-SCD-0001) in Priority
   ========================================================================== */
function switchExecScope(scopeMode) {
    APP_STATE.execScope = scopeMode;
    renderTab1ExecutiveSummary();
}

function setExecMitigationFilter(filterVal) {
    APP_STATE.execMitigationFilter = filterVal;
    renderTab1ExecutiveSummary();
}

function onExecMitigationSearch(val) {
    APP_STATE.execMitigationSearch = val || "";
    renderExecMitigationTableRows();
}

function toggleMitigationMatrix() {
    APP_STATE.showMitigationMatrix = !APP_STATE.showMitigationMatrix;
    renderTab1ExecutiveSummary();
}

function renderTab1ExecutiveSummary() {
    const snap = APP_STATE.snapshot;
    const cardEl = document.getElementById("tab1ExecutiveSummaryCard");
    if (!snap || !cardEl) return;

    const kpis = snap.portfolio_kpis || {};
    const src = snap.sources_summary || {};
    const exec = snap.executive_findings_summary || {};
    const pkg = getSelectedPackage() || {};
    const pkgHl = pkg.package_highlights || {};

    const isZ1F = (snap.project || APP_STATE.project) === "Z1F";
    const defaultMasterFile = isZ1F ? "MM-ZTK-1F-GEN-PJM-SCD-0002.pdf" : "MM-ASK-1A-GEN01-PJM-SCD-0001.pdf";
    const defaultMasterRev = isZ1F ? "C1" : "B1";
    const masterFile = exec.active_master_schedule_file || src.master_schedule_file || defaultMasterFile;
    const masterRev = exec.active_master_schedule_rev || src.master_schedule_rev || defaultMasterRev;
    const kd = exec.project_key_dates || {};

    const isCoreScope = APP_STATE.execScope === "core";
    const tierGroup = (isCoreScope ? exec.tiers_core : exec.tiers_all) || exec.tiers_all || {
        total: 0,
        already_delayed: { count: 0, pct: 0, packages: [] },
        high_risk: { count: 0, pct: 0, packages: [] },
        low_risk: { count: 0, pct: 0, packages: [] },
        healthy: { count: 0, pct: 0, packages: [] },
    };

    const tDelayed = tierGroup.already_delayed || { count: 0, pct: 0, packages: [] };
    const tHigh = tierGroup.high_risk || { count: 0, pct: 0, packages: [] };
    const tLow = tierGroup.low_risk || { count: 0, pct: 0, packages: [] };
    const tHealthy = tierGroup.healthy || { count: 0, pct: 0, packages: [] };
    const totalInScope = tierGroup.total || 1;

    const renderPkgPills = (pkgList) => {
        if (!pkgList || !pkgList.length) {
            return `<span class="blank-val" style="font-size:11px;">None in this category</span>`;
        }
        return pkgList.map(item => {
            const isSel = item.package_code === APP_STATE.selectedPkgCode;
            const cpIcon = item.is_critical_path ? `<span title="Master Schedule Critical Path (TF=0)" style="color:#f43f5e;font-weight:800;">⚡TF=0</span>` : "";
            return `
                <button type="button" class="tier-pkg-pill ${isSel ? "active-pkg" : ""}" onclick="selectPackage('${esc(item.package_code)}')" title="${esc(item.package_name)} — ETA: ${esc(item.eta_date || '—')} vs Master Plan: ${esc(item.master_eta_date || item.ros_date || '—')}\nIssue: ${esc(item.smart_issue || '')}">
                    <span>${esc(item.package_name)}</span>
                    ${cpIcon}
                    <span class="pill-meta">${esc(item.badge_text || "")}</span>
                </button>
            `;
        }).join("");
    };

    cardEl.innerHTML = `
        <div class="exec-summary-header">
            <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                <span class="element-num-badge">One-Picture Executive Summary</span>
                <div>
                    <div class="card-title" style="font-size:14.5px;">
                        ${esc(snap.project)} One-Picture Package Health &amp; Smart Mitigation Summary — Tracked vs. Procurement Plan ROS (<strong>${esc(src.psr_file || "Procurement Plan")}</strong>) in Priority
                    </div>
                    <div class="card-subtitle">
                        <strong>Data Reference:</strong> Procurement Plan ROS Date is governing • Master Schedule (<strong>${esc(masterFile)}</strong> Rev. ${esc(masterRev)}) informs Critical Path (TF=0) &amp; Yard Erection alignment •
                        Active Procurement Plan: <strong>${esc(src.psr_file || "—")}</strong>
                    </div>
                </div>
            </div>

            <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <div class="segmented-control" title="Switch One-Picture View between Core Monitored Packages (31) and All Master Plan / PSR Equipment (${kpis.total_packages || 108})">
                    <button type="button" class="seg-btn ${isCoreScope ? "active" : ""}" onclick="switchExecScope('core')">
                        <span>Core Packages (${kpis.core_packages_count || 31})</span>
                    </button>
                    <button type="button" class="seg-btn ${!isCoreScope ? "active" : ""}" onclick="switchExecScope('all')">
                        <span>All Equipment (${kpis.total_packages || 108})</span>
                    </button>
                </div>
                <button type="button" class="btn-action" onclick="toggleMitigationMatrix()" style="padding:5px 10px;font-size:11.5px;">
                    ${APP_STATE.showMitigationMatrix ? "▲ Hide Issue & Mitigation Table" : "▼ Show Issue & Mitigation Table"}
                </button>
            </div>
        </div>

        <!-- Visual Proportional 4-Tier Health Bar ("One Picture" Across All Packages) -->
        <div class="one-picture-bar-wrap">
            <div class="one-picture-bar-header">
                <span>📊 ONE-PICTURE PORTFOLIO HEALTH DISTRIBUTION (${isCoreScope ? `${totalInScope} Core Monitored Packages` : `All ${totalInScope} Master Plan & PSR Equipment Items`}) — Click any package below to inspect its 9-Point Workspace</span>
                <div style="display:flex;gap:12px;flex-wrap:wrap;font-size:11.5px;">
                    <span style="color:var(--accent-rose);">🔴 Already Delayed: <strong>${tDelayed.count}</strong> (${tDelayed.pct}%)</span>
                    <span style="color:var(--accent-amber);">🟠 High Risk Delay: <strong>${tHigh.count}</strong> (${tHigh.pct}%)</span>
                    <span style="color:var(--accent-blue);">🟡 Low Risk: <strong>${tLow.count}</strong> (${tLow.pct}%)</span>
                    <span style="color:var(--accent-emerald);">🟢 Very Good / Healthy: <strong>${tHealthy.count}</strong> (${tHealthy.pct}%)</span>
                </div>
            </div>
            <div class="one-picture-bar">
                <div class="one-picture-seg seg-delayed" style="width:${Math.max(tDelayed.count > 0 ? 6 : 0, tDelayed.pct)}%;" title="Already Delayed: ${tDelayed.count} (${tDelayed.pct}%)">
                    ${tDelayed.pct >= 6 ? `${tDelayed.count} Delayed (${tDelayed.pct}%)` : ""}
                </div>
                <div class="one-picture-seg seg-high-risk" style="width:${Math.max(tHigh.count > 0 ? 6 : 0, tHigh.pct)}%;" title="High Risk Delay: ${tHigh.count} (${tHigh.pct}%)">
                    ${tHigh.pct >= 6 ? `${tHigh.count} High Risk (${tHigh.pct}%)` : ""}
                </div>
                <div class="one-picture-seg seg-low-risk" style="width:${Math.max(tLow.count > 0 ? 6 : 0, tLow.pct)}%;" title="Low Risk: ${tLow.count} (${tLow.pct}%)">
                    ${tLow.pct >= 6 ? `${tLow.count} Low Risk (${tLow.pct}%)` : ""}
                </div>
                <div class="one-picture-seg seg-healthy" style="width:${Math.max(tHealthy.count > 0 ? 6 : 0, tHealthy.pct)}%;" title="Very Good / Healthy: ${tHealthy.count} (${tHealthy.pct}%)">
                    ${tHealthy.pct >= 6 ? `${tHealthy.count} Healthy (${tHealthy.pct}%)` : ""}
                </div>
            </div>
        </div>

        <!-- 4-Column One-Picture Health Tier Cards -->
        <div class="tier-4col-grid">
            <!-- Tier 1: Already Delayed -->
            <div class="tier-card tier-delayed">
                <div class="tier-card-head">
                    <div>
                        <div class="tier-card-title">🔴 1. ALREADY DELAYED</div>
                        <div class="tier-card-rule">Vendor ETA &gt; Procurement Plan ROS Date</div>
                    </div>
                    <div class="tier-card-count" style="color:var(--accent-rose);">${tDelayed.count}</div>
                </div>
                <div class="tier-pills-box">
                    ${renderPkgPills(tDelayed.packages)}
                </div>
            </div>

            <!-- Tier 2: High Risk Delay -->
            <div class="tier-card tier-high-risk">
                <div class="tier-card-head">
                    <div>
                        <div class="tier-card-title">🟠 2. HIGH RISK DELAY</div>
                        <div class="tier-card-rule">0–7d Buffer vs Procurement Plan ROS, Master Schedule TF=0 Critical Path, or WPR Material Hold</div>
                    </div>
                    <div class="tier-card-count" style="color:var(--accent-amber);">${tHigh.count}</div>
                </div>
                <div class="tier-pills-box">
                    ${renderPkgPills(tHigh.packages)}
                </div>
            </div>

            <!-- Tier 3: Low Risk -->
            <div class="tier-card tier-low-risk">
                <div class="tier-card-head">
                    <div>
                        <div class="tier-card-title">🟡 3. LOW RISK (WATCH)</div>
                        <div class="tier-card-rule">8–14d Buffer Ahead of Procurement Plan ROS; Minor VDRL or Client D2 Review Follow-up</div>
                    </div>
                    <div class="tier-card-count" style="color:var(--accent-blue);">${tLow.count}</div>
                </div>
                <div class="tier-pills-box">
                    ${renderPkgPills(tLow.packages)}
                </div>
            </div>

            <!-- Tier 4: Very Good / Healthy -->
            <div class="tier-card tier-healthy">
                <div class="tier-card-head">
                    <div>
                        <div class="tier-card-title">🟢 4. VERY GOOD / HEALTHY</div>
                        <div class="tier-card-rule">Vendor ETA Ahead of Procurement Plan ROS Date (&gt;14d) • 0 Overdue VDRL Docs</div>
                    </div>
                    <div class="tier-card-count" style="color:var(--accent-emerald);">${tHealthy.count}</div>
                </div>
                <div class="tier-pills-box">
                    ${renderPkgPills(tHealthy.packages)}
                </div>
            </div>
        </div>

        <!-- Smart Issue Identification & Timely Mitigation Action Matrix -->
        <div class="mitigation-section" style="display:${APP_STATE.showMitigationMatrix ? "block" : "none"};">
            <div class="mitigation-toolbar">
                <div>
                    <div class="card-title" style="font-size:13px;">
                        ⚡ Smart Issue Identification &amp; Timely Mitigation Action Matrix (Tracked vs. Procurement Plan ROS)
                    </div>
                    <div class="card-subtitle">
                        Pinpoints root-cause bottlenecks (Procurement Plan ROS breach risk, Yard Topside erection clash, WPR Att. 4 sub-supplier hold, Key VDRL gate) and prescribes on-time mitigation measures
                    </div>
                </div>
                <div class="filter-bar">
                    <select class="select-control" onchange="setExecMitigationFilter(this.value)" title="Filter Mitigation Matrix by Risk Tier">
                        <option value="ACTIONABLE" ${APP_STATE.execMitigationFilter === "ACTIONABLE" ? "selected" : ""}>🔴+🟠 Delayed &amp; High Risk (${tDelayed.count + tHigh.count})</option>
                        <option value="ALREADY_DELAYED" ${APP_STATE.execMitigationFilter === "ALREADY_DELAYED" ? "selected" : ""}>🔴 Already Delayed Only (${tDelayed.count})</option>
                        <option value="HIGH_RISK" ${APP_STATE.execMitigationFilter === "HIGH_RISK" ? "selected" : ""}>🟠 High Risk Delay Only (${tHigh.count})</option>
                        <option value="CRITICAL_PATH" ${APP_STATE.execMitigationFilter === "CRITICAL_PATH" ? "selected" : ""}>⚡ Master Plan TF=0 Critical Path</option>
                        <option value="LOW_RISK" ${APP_STATE.execMitigationFilter === "LOW_RISK" ? "selected" : ""}>🟡 Low Risk Watch Items (${tLow.count})</option>
                        <option value="ALL" ${APP_STATE.execMitigationFilter === "ALL" ? "selected" : ""}>All Packages (${totalInScope})</option>
                    </select>
                    <input type="text" class="input-control" placeholder="Search package, issue, mitigation..." value="${esc(APP_STATE.execMitigationSearch)}" oninput="onExecMitigationSearch(this.value)" style="min-width:190px;" />
                </div>
            </div>

            <div class="mitigation-table-wrap">
                <table class="data-table mitigation-matrix-table">
                    <thead>
                        <tr>
                            <th style="width:16%;">PACKAGE NAME</th>
                            <th style="width:11%;" class="center">HEALTH TIER</th>
                            <th style="width:17%;">VENDOR ETA vs. PROCUREMENT PLAN ROS &amp; MASTER PLAN</th>
                            <th style="width:27%;">SMART ISSUE IDENTIFICATION (ROOT CAUSE)</th>
                            <th style="width:29%;">ON-TIME MITIGATION MEASURE &amp; DEADLINE</th>
                        </tr>
                    </thead>
                    <tbody id="execMitigationTableBody"></tbody>
                </table>
            </div>
        </div>
    `;

    renderExecMitigationTableRows();
}

function renderExecMitigationTableRows() {
    const snap = APP_STATE.snapshot;
    const tbody = document.getElementById("execMitigationTableBody");
    if (!snap || !tbody) return;

    const exec = snap.executive_findings_summary || {};
    const isCoreScope = APP_STATE.execScope === "core";
    const tierGroup = (isCoreScope ? exec.tiers_core : exec.tiers_all) || exec.tiers_all || {};

    const allRows = [
        ...((tierGroup.already_delayed && tierGroup.already_delayed.packages) || []),
        ...((tierGroup.high_risk && tierGroup.high_risk.packages) || []),
        ...((tierGroup.low_risk && tierGroup.low_risk.packages) || []),
        ...((tierGroup.healthy && tierGroup.healthy.packages) || []),
    ];

    const fMode = APP_STATE.execMitigationFilter || "ACTIONABLE";
    const q = (APP_STATE.execMitigationSearch || "").trim().toUpperCase();

    const filtered = allRows.filter(item => {
        const isDelayed = ((tierGroup.already_delayed && tierGroup.already_delayed.packages) || []).includes(item);
        const isHigh = ((tierGroup.high_risk && tierGroup.high_risk.packages) || []).includes(item);
        const isLow = ((tierGroup.low_risk && tierGroup.low_risk.packages) || []).includes(item);

        if (fMode === "ACTIONABLE" && !(isDelayed || isHigh)) return false;
        if (fMode === "ALREADY_DELAYED" && !isDelayed) return false;
        if (fMode === "HIGH_RISK" && !isHigh) return false;
        if (fMode === "CRITICAL_PATH" && !item.is_critical_path) return false;
        if (fMode === "LOW_RISK" && !isLow) return false;

        if (q) {
            const hay = `${item.package_code} ${item.package_name} ${item.discipline} ${item.smart_issue} ${item.smart_mitigation} ${item.activity_id}`.toUpperCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });

    if (!filtered.length) {
        tbody.innerHTML = `<tr><td colspan="5" class="center" style="padding:16px;color:var(--text-muted);">No packages match the current filter.</td></tr>`;
        return;
    }

    tbody.innerHTML = filtered.map(item => {
        const isDelayed = ((tierGroup.already_delayed && tierGroup.already_delayed.packages) || []).includes(item);
        const isHigh = ((tierGroup.high_risk && tierGroup.high_risk.packages) || []).includes(item);
        const isLow = ((tierGroup.low_risk && tierGroup.low_risk.packages) || []).includes(item);
        const isSel = item.package_code === APP_STATE.selectedPkgCode;

        let tierBadge = `<span class="badge badge-success">🟢 Healthy</span>`;
        if (isDelayed) {
            tierBadge = `<span class="badge badge-danger">🔴 Delayed (+${item.effective_delay_days}d)</span>`;
        } else if (isHigh) {
            tierBadge = `<span class="badge badge-warning">🟠 High Risk</span>`;
        } else if (isLow) {
            tierBadge = `<span class="badge badge-info">🟡 Low Risk</span>`;
        }

        const cpBadge = item.is_critical_path
            ? `<div style="margin-top:3px;"><span class="badge badge-danger" style="font-size:9.5px;padding:1px 6px;">⚡ Master Plan TF=0 Critical Path</span></div>`
            : "";

        return `
            <tr class="clickable-row ${isSel ? "selected-row" : ""}" onclick="selectPackage('${esc(item.package_code)}')">
                <td>
                    <strong>${esc(item.package_name)}</strong>
                    <div style="font-size:10.5px;color:var(--text-muted);margin-top:2px;">
                        <span class="badge badge-info" style="font-size:9.5px;padding:1px 5px;">${esc(item.discipline)}</span>
                        ${item.activity_id ? `• ${esc(item.activity_id)}` : ""}
                    </div>
                    ${cpBadge}
                </td>
                <td class="center">
                    ${tierBadge}
                    <div style="font-size:10px;color:var(--text-muted);margin-top:3px;">${esc(item.badge_text || "")}</div>
                </td>
                <td>
                    <div><strong>Vendor ETA:</strong> <span style="color:${isDelayed ? "var(--accent-rose)" : "var(--text-primary)"};font-weight:700;">${fmtBlank(item.eta_date)}</span></div>
                    <div><strong>Procurement Plan ROS:</strong> <strong>${fmtBlank(item.ros_date)}</strong></div>
                    ${item.master_eta_date ? `<div style="font-size:10.5px;color:var(--text-muted);"><strong>Master Plan:</strong> ${esc(item.master_eta_date)}</div>` : ""}
                    ${item.yard_erection_date ? `<div style="font-size:10.5px;color:var(--text-muted);"><strong>Yard Erection:</strong> ${esc(item.yard_erection_date)}</div>` : ""}
                </td>
                <td>
                    <div style="color:var(--text-primary);">${esc(item.smart_issue || "—")}</div>
                </td>
                <td>
                    <div style="color:var(--text-primary);">${esc(item.smart_mitigation || "—")}</div>
                    <div style="margin-top:4px;display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
                        <span class="badge ${isDelayed ? "badge-danger" : (isHigh ? "badge-warning" : "badge-info")}" style="font-size:9.5px;padding:1px 6px;">Target: ${esc(item.mitigation_deadline || "Immediate")}</span>
                        <span style="font-size:10px;color:var(--text-muted);">Owner: ${esc(item.mitigation_owner || "PMT / Vendor")}</span>
                    </div>
                </td>
            </tr>
        `;
    }).join("");
}

/* ==========================================================================
   POINT 4: EQUIPMENT & PACKAGE DROPDOWN SELECTOR (ALL PSR EQUIPMENT INCLUDED)
   ========================================================================== */
function getFilteredTab1Packages() {
    const pkgs = (APP_STATE.snapshot && APP_STATE.snapshot.packages) || [];
    const cov = document.getElementById("tab1CoverageFilter")?.value || "ALL";
    const disc = document.getElementById("tab1DiscFilter")?.value || "ALL";
    const q = (document.getElementById("tab1PkgSearchInput")?.value || "").trim().toUpperCase();

    return pkgs.filter(p => {
        const hl = p.package_highlights || {};
        const isCore = Boolean(p.is_core_package !== undefined ? p.is_core_package : ((p.vdrl_status || {}).available || (p.manufacturer_detail || {}).available));
        if (cov === "CORE" && !isCore) return false;
        if (cov === "PSR_ONLY" && isCore) return false;
        if (cov === "ROS_DELAY" && !hl.is_ros_delayed) return false;
        if (disc !== "ALL" && p.discipline !== disc) return false;
        if (q) {
            const hay = `${p.package_code} ${p.package_name} ${p.discipline} ${p.rfq_no || ""}`.toUpperCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });
}

function stepPackageSelection(delta) {
    const list = getFilteredTab1Packages();
    if (!list.length) return;
    const curIdx = list.findIndex(p => p.package_code === APP_STATE.selectedPkgCode);
    const nextIdx = curIdx < 0 ? 0 : (curIdx + delta + list.length) % list.length;
    selectPackage(list[nextIdx].package_code);
}

function renderPackageSelectorControls(autoSelectFirstOnFilter = false) {
    const allPkgs = (APP_STATE.snapshot && APP_STATE.snapshot.packages) || [];
    const filtered = getFilteredTab1Packages();

    if (autoSelectFirstOnFilter && filtered.length > 0 && !filtered.some(p => p.package_code === APP_STATE.selectedPkgCode)) {
        APP_STATE.selectedPkgCode = filtered[0].package_code;
        renderSelectedPackageWorkspace();
        renderPortfolioTable();
    }

    const coreList = filtered.filter(p => Boolean(p.is_core_package !== undefined ? p.is_core_package : ((p.vdrl_status || {}).available || (p.manufacturer_detail || {}).available)));
    const psrOnlyList = filtered.filter(p => !coreList.includes(p));

    const countLabel = document.getElementById("pkgDropdownCountLabel");
    if (countLabel) {
        const selPkg = getSelectedPackage();
        const selMeta = selPkg
            ? ` • Active: [${selPkg.package_code}] ${selPkg.has_live_wpr ? `Live ${selPkg.report_no} (Cut-off: ${selPkg.cut_off_date})` : "SCD + VDRL Baseline"}${selPkg.has_live_mom ? ` + ${(selPkg.mom_summary || {}).mom_no || "MOM"}` : ""}`
            : "";
        countLabel.textContent = `Showing ${filtered.length} of ${allPkgs.length} Equipment / Packages (${coreList.length} Core VDRL/KOM Pkgs + ${psrOnlyList.length} PSR Equipment Items)${selMeta}`;
    }

    const formatOption = (p) => {
        return `<option value="${esc(p.package_code)}" ${p.package_code === APP_STATE.selectedPkgCode ? "selected" : ""}>${esc(p.package_name)}</option>`;
    };

    let optionsHtml = "";
    if (coreList.length > 0) {
        optionsHtml += `<optgroup label="Core Monitored Packages (With VDRL / KOM / Client D2 Log — ${coreList.length})">${coreList.map(formatOption).join("")}</optgroup>`;
    }
    if (psrOnlyList.length > 0) {
        optionsHtml += `<optgroup label="PSR Equipment Catalog (Awaiting VDRL: Barred Tee, Corrosion Coupon, Sight Glass, etc. — ${psrOnlyList.length})">${psrOnlyList.map(formatOption).join("")}</optgroup>`;
    }
    if (!optionsHtml) {
        optionsHtml = `<option value="">No matching equipment found</option>`;
    }

    const tab1Drop = document.getElementById("tab1EquipmentDropdown");
    if (tab1Drop) {
        tab1Drop.innerHTML = optionsHtml;
    }

    const selectEl = document.getElementById("quickPackageSelect");
    if (selectEl) {
        selectEl.innerHTML = allPkgs.map(formatOption).join("");
    }
}

/* ==========================================================================
   COMMENT #4.1 & #6: PACKAGE KEY HIGHLIGHT INFORMATION BANNER
   ========================================================================== */
function renderPackageHighlightBanner(pkg) {
    const hl = pkg.package_highlights || {};
    const bannerEl = document.getElementById("packageHighlightBanner");
    if (!bannerEl) return;

    const remDays = hl.remaining_days_to_next_milestone;
    const remDaysText = remDays === null || remDays === undefined
        ? "—"
        : (remDays < 0 ? `${Math.abs(remDays)} Days Overdue` : `${remDays} Days Remaining`);
    const remTone = remDays !== null && remDays < 0 ? "hl-rose" : (remDays !== null && remDays <= 14 ? "hl-amber" : "hl-emerald");
    const remBadge = remDays !== null && remDays < 0 ? "badge-danger" : (remDays !== null && remDays <= 14 ? "badge-warning" : "badge-info");

    const delDays = hl.days_to_delivery;
    const delDaysText = delDays === null || delDays === undefined
        ? "—"
        : (delDays < 0 ? `${Math.abs(delDays)} Days Past ETA` : `${delDays} Days to Delivery`);

    const isRosDelayed = Boolean(hl.is_ros_delayed);
    const rosTone = isRosDelayed ? "hl-rose" : (hl.ros_margin_days !== null && hl.ros_margin_days <= 5 ? "hl-amber" : "hl-emerald");
    const rosBadge = isRosDelayed ? "badge-danger" : (hl.ros_margin_days !== null && hl.ros_margin_days <= 5 ? "badge-warning" : "badge-success");

    const hasDiscrepancy = Boolean(pkg.has_milestone_discrepancy || hl.has_milestone_discrepancy);
    const discDays = pkg.milestone_discrepancy_days ?? hl.milestone_discrepancy_days ?? 0;
    const ganttEta = pkg.gantt_eta_date || hl.gantt_eta_date || "";
    const msEta = pkg.milestone_eta_date || hl.milestone_eta_date || "";
    const ganttTaskName = pkg.gantt_eta_task_name || hl.gantt_eta_task_name || "4.2 Sea Freight / Transit to China Port (CIF)";

    const discrepancyBannerHtml = hasDiscrepancy
        ? `<div class="discrepancy-alert-banner">
               <div class="alert-main">
                   <div class="alert-title" style="color:var(--accent-amber);font-size:13.5px;">
                       ⚠️ Schedule Alignment Alert: WPR Gantt Chart ETA (${esc(ganttEta)}) vs. Section 5 Key Milestone #10 (${esc(msEta)}) Discrepancy (${discDays > 0 ? `+${discDays}` : discDays} Days)
                   </div>
                   <div class="alert-detail" style="font-size:12px;line-height:1.45;">
                       • <strong>Governing Package ETA (Follows WPR Att. 3 Gantt Chart):</strong> <strong>${esc(ganttEta)}</strong> — from Task <em>"${esc(ganttTaskName)}"</em> (Cut-off: <strong>${esc(pkg.cut_off_date)}</strong>, Ref: <strong>${esc(pkg.wpr_source_file || pkg.document_no)}</strong>).<br/>
                       • <strong>Section 5 Key Milestone #10 ('Receipt at worksite'):</strong> Lists <strong>${esc(msEta)}</strong> (which corresponds to Gantt Task 4.1 <em>Loading, Inland Transport &amp; Export Customs Clearance</em> departure date, ${Math.abs(discDays)} days before CIF China Port arrival).<br/>
                       • <strong>Comparison vs. Master Schedule (SCD) ROS Date:</strong> Governing ETA (<strong>${esc(ganttEta)}</strong>) vs. Master Schedule ROS (<strong>${esc(hl.master_eta_date || hl.governing_ros_date || hl.ros_date)}</strong>) provides a <strong>+${esc(hl.ros_margin_days)} days buffer</strong> (Yard Erection: ${esc(hl.yard_erection_date || "2027-03-08")}).
                   </div>
               </div>
               <div style="display:flex;flex-direction:column;gap:5px;align-items:flex-end;">
                   <span class="badge badge-warning">⚡ Gantt ETA Governing: ${esc(ganttEta)}</span>
                   <span class="badge badge-danger">Milestone #10 Diff: ${discDays > 0 ? `+${discDays}d` : `${discDays}d`} (${esc(msEta)})</span>
                   <span class="badge badge-success">vs SCD ROS (${esc(hl.master_eta_date || hl.ros_date)}): +${esc(hl.ros_margin_days)}d Buffer</span>
               </div>
           </div>`
        : "";

    bannerEl.innerHTML = `
        ${discrepancyBannerHtml}
        <!-- 4.1A: Current Status Based on Key Milestone -->
        <div class="highlight-card hl-purple">
            <div class="highlight-head">
                <span class="highlight-tag">4.1A • Current Key Milestone Status</span>
                <span class="badge badge-purple">${esc(pkg.package_code)} • Cut-off: ${esc(pkg.cut_off_date || "—")}</span>
            </div>
            <div class="highlight-big">
                <span>${esc(hl.active_stage_summary || "In Progress")}</span>
            </div>
            <div class="highlight-detail">
                <strong>Latest Completed:</strong> ${esc(hl.current_stage_summary || "—")}<br/>
                <strong>Overall Progress:</strong> Actual <strong>${Number((pkg.overall_progress || {}).actual_pct || 0).toFixed(2)}%</strong> vs Plan <strong>${Number((pkg.overall_progress || {}).planned_pct || 0).toFixed(2)}%</strong> • Ref: <strong>${esc(pkg.has_live_wpr ? (pkg.wpr_source_file || pkg.report_no) : "SCD / PSR Baseline")}</strong>
            </div>
        </div>

        <!-- 4.1B: Number of Remaining Days to Next Key Milestone -->
        <div class="highlight-card ${remTone}">
            <div class="highlight-head">
                <span class="highlight-tag">4.1B • Countdown to Next Key Milestone</span>
                <span class="badge ${remBadge}">Cut-off: ${esc(hl.reference_date || pkg.cut_off_date)}</span>
            </div>
            <div class="highlight-big">
                <span>⏱️ ${esc(remDaysText)}</span>
            </div>
            <div class="highlight-detail">
                <strong>Next Milestone:</strong> ${esc(hl.next_milestone_name || "—")}<br/>
                <strong>Target Date:</strong> ${fmtBlank(hl.next_milestone_date)}
            </div>
        </div>

        <!-- 4.1C: How Many Days to Delivery (Governing ETA follows Gantt Chart if available) -->
        <div class="highlight-card ${isRosDelayed ? "hl-rose" : (hasDiscrepancy ? "hl-amber" : "hl-emerald")}">
            <div class="highlight-head">
                <span class="highlight-tag">4.1C • Countdown to Delivery (Governing ETA)</span>
                <span class="badge ${hasDiscrepancy ? "badge-warning" : "badge-info"}">${esc(hl.delivery_source || "Receipt at Worksite")}</span>
            </div>
            <div class="highlight-big">
                <span>🚢 ${esc(delDaysText)}</span>
            </div>
            <div class="highlight-detail">
                <strong>Governing Package ETA:</strong> <strong>${fmtBlank(hl.effective_delivery_date)}</strong> ${hasDiscrepancy ? `<span class="badge badge-warning" style="font-size:10px;padding:1px 5px;">⚠️ Milestone #10: ${esc(msEta)} (${discDays > 0 ? `+${discDays}d` : `${discDays}d`})</span>` : ""}<br/>
                <strong>Governing Target ROS:</strong> ${fmtBlank(hl.governing_ros_date || hl.ros_date)} (${esc(hl.governing_ros_source || "Master Schedule")})<br/>
                ${hl.master_eta_date ? `<strong>Master Schedule ROS:</strong> ${esc(hl.master_eta_date)} • ` : ""}<strong>Procurement Plan ROS:</strong> ${fmtBlank(hl.psr_ros_date || hl.ros_date)}
            </div>
        </div>

        <!-- 4.1D & Comment #6: Priority #1 Master Schedule & Procurement Plan ROS Delivery Check -->
        <div class="highlight-card ${rosTone}">
            <div class="highlight-head">
                <span class="highlight-tag">6 • 🎯 Governing ROS Delivery Compliance</span>
                <span class="badge ${rosBadge}">${isRosDelayed ? `🚨 +${hl.effective_delay_days || hl.ros_delay_days}d DELAY` : (hl.health_status_label || "Compliant")}</span>
            </div>
            <div class="highlight-big">
                <span>ROS: ${fmtBlank(hl.governing_ros_date || hl.ros_date)}</span>
            </div>
            <div class="highlight-detail">
                <strong>Governing Rule (ETA ≤ ROS):</strong> ${esc(hl.ros_status_label || "—")}<br/>
                <strong>Mitigation:</strong> ${esc(hl.smart_mitigation || "On track vs Master Schedule & Procurement Plan ROS.")}
            </div>
        </div>
    `;
}

/* ==========================================================================
   9-POINT PACKAGE MONITORING WORKSPACE RENDERER
   ========================================================================== */
function renderSelectedPackageWorkspace() {
    const pkg = getSelectedPackage();
    if (!pkg) return;

    // Point 3: Update Top Executive Summary card with active package context
    renderTab1ExecutiveSummary();

    // 4.1 & 6: Package Highlight Key Information Banner
    renderPackageHighlightBanner(pkg);

    // Stamp Cut-off Date & Reference Document into each of the 9 Elements
    renderElementRefStamps(pkg);

    // ----------------------------------------------------------------------
    // ELEMENTS 1 & 2: Project & Package Name / Identification
    // ----------------------------------------------------------------------
    document.getElementById("pkgHeaderTitle").textContent =
        `[${pkg.project}] ${pkg.package_code} — ${pkg.package_name}`;
    document.getElementById("pkgHeaderSub").textContent =
        `Discipline: ${pkg.discipline} • RFQ: ${pkg.rfq_no || "—"} • Cut-off Date: ${pkg.cut_off_date || "—"} • Pkg Weight: ${Number(pkg.package_weight_pct || 0).toFixed(2)}%`;

    const hl = pkg.package_highlights || {};
    const priorityClass = pkg.has_live_wpr ? "badge-success" : (pkg.using_sample_wpr ? "badge-info" : "badge-purple");
    const momBadgeHtml = pkg.has_live_mom && pkg.mom_summary && pkg.mom_summary.available
        ? `<span class="badge badge-info">📋 Live ${esc(pkg.mom_summary.mom_no)} (${pkg.mom_summary.open_count || 0} Open)</span>`
        : "";
    document.getElementById("pkgHeaderBadges").innerHTML = `
        <span class="badge ${priorityClass}">${esc(pkg.data_priority_badge)}</span>
        ${momBadgeHtml}
        ${hl.is_ros_delayed ? `<span class="badge badge-danger">🚨 ROS Delay (+${hl.ros_delay_days}d After ROS)</span>` : `<span class="badge badge-success">🔒 ROS Compliant</span>`}
        <span class="badge ${badgeClassForTone(pkg.health_tone)}">${esc(pkg.health_status)}</span>
    `;

    document.getElementById("pkgIdentityGrid").innerHTML = `
        <div class="kv-box">
            <div class="kv-label">1. Project</div>
            <div class="kv-val">${esc(pkg.project)} (${pkg.project === "ASK" ? "Aung Sinkha Phase 1A" : "Zawtika Phase 1F"})</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">2. Package Name &amp; Code</div>
            <div class="kv-val">${esc(pkg.package_code)} — ${esc(pkg.package_name)}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Discipline &amp; Portfolio Weight</div>
            <div class="kv-val">${esc(pkg.discipline)} • Weight: ${Number(pkg.package_weight_pct || 0).toFixed(2)}%</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Equipment Tag No.</div>
            <div class="kv-val">${fmtBlank(pkg.equipment_tag)}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">PO Number &amp; PO Date</div>
            <div class="kv-val">${fmtBlank(pkg.po_number)} ${pkg.po_date ? `(${esc(pkg.po_date)})` : ""}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">WPR Report No. &amp; Cut-off Date</div>
            <div class="kv-val">${fmtBlank(pkg.report_no)} (Rev. ${fmtBlank(pkg.revision)}) • ${fmtBlank(pkg.cut_off_date)}</div>
        </div>
        <div class="kv-box" style="${hl.is_ros_delayed ? "border-color:var(--accent-rose);background:rgba(244,63,94,0.09);" : ""}">
            <div class="kv-label">🔒 Immutable PSR ROS Date vs. ETA</div>
            <div class="kv-val">
                Fixed ROS: <strong>${fmtBlank(hl.ros_date || pkg.ros_date)}</strong> • ETA: <strong>${fmtBlank(hl.effective_delivery_date || pkg.eta_date)}</strong>
            </div>
        </div>
        <div class="kv-box" style="${hl.is_ros_delayed ? "border-color:var(--accent-rose);background:rgba(244,63,94,0.09);" : ""}">
            <div class="kv-label">ROS Float / Delay Status</div>
            <div class="kv-val" style="color:${hl.is_ros_delayed ? "var(--accent-rose)" : "var(--accent-emerald)"};">
                ${esc(hl.ros_status_label || (pkg.float_days !== null ? `${pkg.float_days} days` : "—"))}
            </div>
        </div>
    `;

    // ----------------------------------------------------------------------
    // ELEMENT 3: Manufacturer Detail (Strictly from KOM/ Folder; blank if absent)
    // ----------------------------------------------------------------------
    const mfr = pkg.manufacturer_detail || {};
    const komBadgeEl = document.getElementById("komStatusBadge");
    if (mfr.available) {
        komBadgeEl.innerHTML = `<span class="badge badge-success">✓ KOM Verified (${esc(mfr.kom_source_file)})</span>`;
    } else {
        komBadgeEl.innerHTML = `<span class="badge badge-muted">No KOM File in KOM/ (Blank per Spec)</span>`;
    }

    document.getElementById("manufacturerGrid").innerHTML = `
        <div class="kv-box">
            <div class="kv-label">Manufacturer / Vendor Name</div>
            <div class="kv-val">${mfr.available ? fmtBlank(mfr.manufacturer_name) : `<span class="blank-val">— (Blank: No KOM in KOM/)</span>`}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Vendor Code</div>
            <div class="kv-val">${mfr.available ? fmtBlank(mfr.vendor_code) : `<span class="blank-val">—</span>`}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Factory / Manufacturing Location</div>
            <div class="kv-val">${mfr.available ? fmtBlank(mfr.factory_location) : `<span class="blank-val">—</span>`}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Kick-Off Meeting (KOM) Date</div>
            <div class="kv-val">${mfr.available ? fmtBlank(mfr.kom_date) : `<span class="blank-val">—</span>`}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Contractual PO Reference (KOM)</div>
            <div class="kv-val">${mfr.available ? fmtBlank(mfr.po_number) : `<span class="blank-val">—</span>`}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Key Vendor Contacts (KOM)</div>
            <div class="kv-val">${mfr.available ? fmtBlank(mfr.key_contacts) : `<span class="blank-val">—</span>`}</div>
        </div>
    `;

    // ----------------------------------------------------------------------
    // ELEMENT 7: Overall Progress (Dynamic Package S-Curve + Weighted Project S-Curve)
    // ----------------------------------------------------------------------
    renderPackageSCurve(pkg);
    renderScurveWeightsTable();

    // ----------------------------------------------------------------------
    // ELEMENT 8: Look Ahead Plan (7–14 Days) — Dynamic per Package
    // ----------------------------------------------------------------------
    renderLookAheadPlan(pkg);

    // ----------------------------------------------------------------------
    // ELEMENT 9: Area of Concern — Dynamic per Package
    // ----------------------------------------------------------------------
    renderAreasOfConcern(pkg);

    // ----------------------------------------------------------------------
    // COMMENT #5: WPR Attachment 3 (Fabrication Gantt) & Attachment 4 (Supplier Material Status)
    // ----------------------------------------------------------------------
    renderAttachments3And4(pkg);

    // ----------------------------------------------------------------------
    // MOM INTEGRATION: Overall Situation, Key Concerns & Action Items
    // ----------------------------------------------------------------------
    renderPackageMomSummary(pkg);

    // ----------------------------------------------------------------------
    // ELEMENT 4: Key Milestone Table (12 Standard Milestones, WPR #1 > PSR #2)
    // ----------------------------------------------------------------------
    renderKeyMilestonesTable(pkg);

    // ----------------------------------------------------------------------
    // ELEMENT 5: Progress Tracking by Key Milestone
    // ----------------------------------------------------------------------
    renderMilestoneProgressTable(pkg);

    // ----------------------------------------------------------------------
    // ELEMENT 6: Status Vendor Document (6A VDRL + 6B Contractual Client Log)
    // ----------------------------------------------------------------------
    renderPackageDocuments();
}

function fmtElementRefStampHtml(metaObj, fallbackRef, fallbackCutOff) {
    const ref = (metaObj && metaObj.ref_doc) || fallbackRef || "Master Schedule (SCD) / PSR Baseline";
    const cut = (metaObj && metaObj.cut_off_date) || fallbackCutOff || "—";
    return `<div class="element-ref-stamp">
        <span>📄 <strong>Ref Doc:</strong> ${esc(ref)}</span>
        <span>•</span>
        <span>📅 <strong>Cut-off Date:</strong> ${esc(cut)}</span>
    </div>`;
}

function renderElementRefStamps(pkg) {
    if (!pkg) return;
    const em = pkg.element_metadata || {};
    const cut = pkg.cut_off_date || "2026-10-16";
    const wprRef = pkg.has_live_wpr
        ? `${pkg.wpr_source_file || pkg.document_no} (${pkg.report_no || "WPR"})`
        : "No WPR Uploaded — Master Schedule (SCD) & PSR Baseline";

    const mapIds = [
        ["elRefStamp12", em.element_1_2, wprRef, cut],
        ["elRefStamp3", em.element_3, (pkg.manufacturer_detail || {}).kom_source_file || "No KOM Uploaded (KOM/)", (pkg.manufacturer_detail || {}).kom_date || "—"],
        ["elRefStamp4", em.element_4, wprRef, cut],
        ["elRefStamp5", em.element_5, wprRef, cut],
        ["elRefStamp6", em.element_6, (pkg.vdrl_status || {}).source_file || "VDRL & Client Document Log (D2)", (pkg.vdrl_status || {}).today_date || cut],
        ["elRefStamp7", em.element_7, wprRef, cut],
        ["elRefStamp8", em.element_8, wprRef, cut],
        ["elRefStamp9", em.element_9, wprRef, cut],
        ["elRefStampAtt3", em.element_4, pkg.has_live_wpr ? `${pkg.wpr_source_file} (Attachment 3 Gantt)` : "Master Schedule (SCD) & PSR Baseline", cut],
        ["elRefStampAtt4", em.element_5, pkg.has_live_wpr ? `${pkg.wpr_source_file} (Attachment 4 Sub-Orders)` : "Master Schedule (SCD) & PSR Baseline", cut],
        ["elRefStampMom", em.mom, (pkg.mom_summary && pkg.mom_summary.source_file) || "No MOM Uploaded", (pkg.mom_summary && pkg.mom_summary.cut_off_date) || "—"],
    ];

    mapIds.forEach(([id, metaObj, fbRef, fbCut]) => {
        const el = document.getElementById(id);
        if (el) {
            el.innerHTML = fmtElementRefStampHtml(metaObj, fbRef, fbCut);
        }
    });
}

function renderPackageMomSummary(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const mom = pkg.mom_summary || {};
    const titleEl = document.getElementById("momCardTitle");
    const subEl = document.getElementById("momCardSubtitle");
    const badgesEl = document.getElementById("momCountsBadges");
    const bodyEl = document.getElementById("momSummaryBody");
    if (!bodyEl) return;

    if (titleEl) {
        titleEl.textContent = mom.available
            ? `Minutes of Meeting (${mom.mom_no || "MOM"}) — [${pkg.package_code}] ${pkg.package_name}`
            : `Minutes of Meeting (MOM) — [${pkg.package_code}] ${pkg.package_name}`;
    }
    if (subEl) {
        subEl.textContent = mom.available
            ? `Document: ${mom.document_no || mom.source_file} • Meeting / Cut-off Date: ${mom.cut_off_date || mom.meeting_date || "—"}`
            : `No MOM uploaded in MOM/ for ${pkg.package_code} — showing available Master Schedule (SCD) & VDRL information only`;
    }

    if (!mom.available) {
        if (badgesEl) {
            badgesEl.innerHTML = `<span class="badge badge-muted">No MOM Uploaded for ${esc(pkg.package_code)}</span>`;
        }
        bodyEl.innerHTML = `
            <div class="alert-item">
                <div class="alert-main">
                    <div class="alert-title">ℹ️ No Minutes of Meeting (MOM) File Uploaded for [${esc(pkg.package_code)}] ${esc(pkg.package_name)}</div>
                    <div class="alert-detail">
                        In Execution Mode, only verified data from the Master Schedule (SCD), Procurement Plan (PSR), and VDRL is displayed when a package does not yet have a MOM or WPR file. Upload a MOM PDF via <strong>+ Drop / Upload Weekly Files</strong> to populate this section.
                    </div>
                </div>
                <span class="badge badge-info">SCD &amp; VDRL Baseline Only</span>
            </div>
        `;
        return;
    }

    if (badgesEl) {
        badgesEl.innerHTML = `
            <span class="badge badge-purple">Total: ${mom.total_items || 0} Items</span>
            <span class="badge badge-danger">OPEN: ${mom.open_count || 0}</span>
            <span class="badge badge-success">CLOSED: ${mom.closed_count || 0}</span>
            <span class="badge badge-info">INFO: ${mom.info_count || 0}</span>
        `;
    }

    const statusFilter = document.getElementById("momStatusFilter")?.value || "ALL";
    const q = (document.getElementById("momSearchInput")?.value || "").trim().toUpperCase();

    const keyConcerns = mom.key_concerns || [];
    const concernsHtml = keyConcerns.map(c => {
        const sevClass = c.severity === "high" ? "sev-high" : "sev-medium";
        const badgeTone = c.severity === "high" ? "badge-danger" : "badge-warning";
        return `
            <div class="alert-item ${sevClass}">
                <div class="alert-main">
                    <div class="alert-title">${esc(c.title)}</div>
                    <div class="alert-detail">${esc(c.detail)}</div>
                </div>
                <div style="display:flex;flex-direction:column;gap:4px;align-items:flex-end;">
                    <span class="badge ${badgeTone}">${esc(c.status || "OPEN")} • Action: ${esc(c.action_by || "Vendor")}</span>
                    <span style="font-size:10.5px;color:var(--text-muted);">Target: ${esc(c.target_date || "—")}</span>
                </div>
            </div>
        `;
    }).join("");

    const actionItems = (mom.action_items || []).filter(it => {
        const st = String(it.status || "").toUpperCase();
        if (statusFilter !== "ALL" && st !== statusFilter) return false;
        if (q) {
            const hay = `${it.item_no} ${it.description} ${it.action_by} ${it.target_date} ${it.status}`.toUpperCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });

    const procItems = (mom.procurement_status || []).filter(p => {
        if (!q) return true;
        return `${p.item_no} ${p.component} ${p.plan_date} ${p.actual_date} ${p.status}`.toUpperCase().includes(q);
    });

    const actionsRowsHtml = actionItems.length
        ? actionItems.map(it => {
            const st = String(it.status || "INFO").toUpperCase();
            const stBadge = st === "OPEN" ? "badge-danger" : (st === "CLOSED" ? "badge-success" : "badge-info");
            return `
                <tr>
                    <td class="center"><strong>${esc(it.item_no)}</strong></td>
                    <td>${esc(it.description)}</td>
                    <td class="center"><span class="badge badge-muted">${esc(it.action_by || "—")}</span></td>
                    <td class="center">${fmtBlank(it.target_date)}</td>
                    <td class="center"><span class="badge ${stBadge}">${esc(st)}</span></td>
                </tr>
            `;
        }).join("")
        : `<tr><td colspan="5" class="center" style="padding:16px;color:var(--text-muted);">No MOM action items match the current filter.</td></tr>`;

    const procRowsHtml = procItems.length
        ? procItems.map(p => `
            <tr>
                <td class="center"><strong>${esc(p.item_no)}</strong></td>
                <td><strong>${esc(p.component)}</strong></td>
                <td class="center">${fmtBlank(p.plan_date)}</td>
                <td class="center">${fmtBlank(p.actual_date)}</td>
                <td><span class="badge badge-info">${esc(p.status || "In Progress")}</span></td>
            </tr>
        `).join("")
        : `<tr><td colspan="5" class="center" style="padding:16px;color:var(--text-muted);">No sub-vendor procurement rows in this MOM.</td></tr>`;

    bodyEl.innerHTML = `
        <!-- Overall MOM Situation Summary Box -->
        <div class="alert-item sev-lookahead" style="margin-bottom:12px;">
            <div class="alert-main">
                <div class="alert-title">📌 Overall Situation Summary — ${esc(mom.document_no || mom.mom_no)} (Cut-off / Meeting Date: ${esc(mom.cut_off_date || mom.meeting_date)})</div>
                <div class="alert-detail" style="line-height:1.45;">${esc(mom.overall_situation_summary || "")}</div>
            </div>
            <span class="badge badge-success">Incorporates ${esc(mom.source_file)}</span>
        </div>

        <!-- Highlighted Concerns from MOM -->
        ${concernsHtml ? `<div style="margin-bottom:14px;"><div class="kv-label" style="margin-bottom:6px;">🚨 Highlighted Concerns &amp; Critical Situations from ${esc(mom.mom_no)}</div><div class="concern-list">${concernsHtml}</div></div>` : ""}

        <!-- Dual MOM Tables: Action Items & Section 3.4 Procurement Status -->
        <div class="grid-equal-2col">
            <div>
                <div class="table-toolbar">
                    <div class="kv-label">MOM Action Items &amp; Technical Clarifications (${actionItems.length} shown)</div>
                </div>
                <div class="table-scroll" style="max-height:320px;overflow-y:auto;">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th class="center" style="width:11%;">Item #</th>
                                <th style="width:51%;">Discussion / Action Item Description</th>
                                <th class="center" style="width:12%;">Action By</th>
                                <th class="center" style="width:14%;">Target Date</th>
                                <th class="center" style="width:12%;">Status</th>
                            </tr>
                        </thead>
                        <tbody>${actionsRowsHtml}</tbody>
                    </table>
                </div>
            </div>

            <div>
                <div class="table-toolbar">
                    <div class="kv-label">MOM Section 3.4 — Sub-Vendor Procurement &amp; Expediting Status (${procItems.length} items)</div>
                </div>
                <div class="table-scroll" style="max-height:320px;overflow-y:auto;">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th class="center" style="width:10%;">#</th>
                                <th style="width:36%;">Sub-Component / Equipment</th>
                                <th class="center" style="width:15%;">Plan PO Date</th>
                                <th class="center" style="width:15%;">Actual PO Date</th>
                                <th style="width:24%;">Delivery / Expediting Status</th>
                            </tr>
                        </thead>
                        <tbody>${procRowsHtml}</tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
}

/* ==========================================================================
   ELEMENT 7: INTERACTIVE SVG S-CURVE RENDERER (PACKAGE vs WEIGHTED ALL PKGS)
   ========================================================================== */
function renderPackageSCurve(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const btnPkgLabel = document.getElementById("btnScurvePkgLabel");
    if (btnPkgLabel) {
        btnPkgLabel.textContent = `Package S-Curve (${pkg.package_code})`;
    }

    const isProjectView = APP_STATE.scurveView === "project";
    const projSc = (APP_STATE.snapshot && APP_STATE.snapshot.project_overall_scurve) || {};
    const sc = isProjectView ? projSc : (pkg.scurve || {});
    const pts = sc.points || [];
    const ov = pkg.overall_progress || {};

    const planPct = Number(isProjectView ? projSc.planned_to_date_pct : (ov.planned_pct || sc.planned_to_date_pct || 0)).toFixed(2);
    const actPct = Number(isProjectView ? projSc.actual_to_date_pct : (ov.actual_pct || sc.actual_to_date_pct || 0)).toFixed(2);
    const varPct = Number(isProjectView ? projSc.variance_pct : (ov.variance_pct || sc.variance_pct || 0)).toFixed(2);
    const varSign = Number(varPct) > 0 ? `+${varPct}%` : `${varPct}%`;

    document.getElementById("scurveCardTitle").textContent = isProjectView
        ? `Overall Progress S-Curve — Weighted Across All ${((APP_STATE.snapshot && APP_STATE.snapshot.packages) || []).length} ${APP_STATE.project} Packages`
        : `Overall Progress S-Curve — [${pkg.package_code}] ${pkg.package_name}`;

    document.getElementById("scurveSubtitle").textContent = isProjectView
        ? (projSc.calculation_basis || "Calculated from weighted progress across all packages")
        : `Dynamic Package S-Curve (${pkg.package_code}) • Pkg Portfolio Weight: ${Number(pkg.package_weight_pct || 0).toFixed(2)}% • Cut-off: ${pkg.cut_off_date}`;

    document.getElementById("scurveKpiPills").innerHTML = `
        <span class="badge badge-info">Planned: ${planPct}%</span>
        <span class="badge badge-success">Actual: ${actPct}%</span>
        <span class="badge ${Number(varPct) >= 0 ? "badge-success" : "badge-warning"}">Variance: ${varSign}</span>
    `;
    document.getElementById("scurveCutOffLabel").textContent =
        sc.current_week ? `Reporting Week: ${sc.current_week} (${sc.cut_off_date || ""})` : "";

    const container = document.getElementById("scurveSvgContainer");
    const labelColor = APP_STATE.theme === "light" ? "#475569" : "#94a3b8";
    const gridStroke = APP_STATE.theme === "light" ? "rgba(100,116,139,0.22)" : "rgba(148,163,184,0.14)";
    const dotStroke = APP_STATE.theme === "light" ? "#ffffff" : "#0a0f1d";

    if (!pts.length) {
        container.innerHTML = `<div style="padding:36px;text-align:center;color:var(--text-muted);">No S-Curve points available for ${esc(pkg.package_code)}.</div>`;
    } else {
        const W = 760, H = 245;
        const padL = 46, padR = 20, padT = 18, padB = 34;
        const plotW = W - padL - padR;
        const plotH = H - padT - padB;

        const xForIdx = (i) => padL + (i / Math.max(pts.length - 1, 1)) * plotW;
        const yForPct = (pct) => padT + plotH - (Math.min(Math.max(Number(pct || 0), 0), 100) / 100) * plotH;

        const plannedCoords = pts.map((p, i) => `${xForIdx(i).toFixed(1)},${yForPct(p.planned_pct).toFixed(1)}`);
        const plannedArea = `${xForIdx(0).toFixed(1)},${(padT + plotH).toFixed(1)} ` + plannedCoords.join(" ") + ` ${xForIdx(pts.length - 1).toFixed(1)},${(padT + plotH).toFixed(1)}`;

        const actualPts = pts.map((p, i) => ({ ...p, idx: i })).filter(p => p.actual_pct !== null && p.actual_pct !== undefined);
        const actualCoords = actualPts.map(p => `${xForIdx(p.idx).toFixed(1)},${yForPct(p.actual_pct).toFixed(1)}`);

        const cutOffIdx = actualPts.length > 0 ? actualPts[actualPts.length - 1].idx : 3;
        const cutX = xForIdx(cutOffIdx);

        let gridLines = "";
        [0, 25, 50, 75, 100].forEach(val => {
            const y = yForPct(val).toFixed(1);
            gridLines += `
                <line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="${gridStroke}" stroke-dasharray="3,3" />
                <text x="${padL - 8}" y="${Number(y) + 4}" text-anchor="end" fill="${labelColor}" font-size="10.5">${val}%</text>
            `;
        });

        let xLabels = "";
        pts.forEach((p, i) => {
            if (i === 0 || (i + 1) % 5 === 0 || i === pts.length - 1 || i === cutOffIdx) {
                const x = xForIdx(i).toFixed(1);
                const cutColor = APP_STATE.theme === "light" ? "#d97706" : "#fbbf24";
                xLabels += `<text x="${x}" y="${H - 10}" text-anchor="middle" fill="${i === cutOffIdx ? cutColor : labelColor}" font-size="10" font-weight="${i === cutOffIdx ? "700" : "400"}">${esc(p.week)}</text>`;
            }
        });

        const actualDots = actualPts.map(p => `
            <circle cx="${xForIdx(p.idx).toFixed(1)}" cy="${yForPct(p.actual_pct).toFixed(1)}" r="4.2" fill="#10b981" stroke="${dotStroke}" stroke-width="1.5">
                <title>${esc(p.week)} (${esc(p.date)}): Actual ${p.actual_pct}% vs Planned ${p.planned_pct}%</title>
            </circle>
        `).join("");

        container.innerHTML = `
            <svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block;">
                <defs>
                    <linearGradient id="planGrad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stop-color="#38bdf8" stop-opacity="0.28" />
                        <stop offset="100%" stop-color="#38bdf8" stop-opacity="0.0" />
                    </linearGradient>
                </defs>
                ${gridLines}
                <polygon points="${plannedArea}" fill="url(#planGrad)" />
                <polyline fill="none" stroke="#38bdf8" stroke-width="2.4" points="${plannedCoords.join(" ")}" />
                ${actualCoords.length > 1 ? `<polyline fill="none" stroke="#10b981" stroke-width="3" points="${actualCoords.join(" ")}" />` : ""}
                <line x1="${cutX.toFixed(1)}" y1="${padT}" x2="${cutX.toFixed(1)}" y2="${padT + plotH}" stroke="#f59e0b" stroke-width="1.6" stroke-dasharray="4,4" />
                ${actualDots}
                ${xLabels}
            </svg>
        `;
    }

    // Dynamic This Period Activities per Package
    const acts = (pkg.overall_progress && pkg.overall_progress.this_period_activities) || [];
    const titleLabel = document.getElementById("thisPeriodTitleLabel");
    if (titleLabel) {
        titleLabel.textContent = `WPR Section 1 — This Period Activities (${pkg.package_code} — ${pkg.package_name})`;
    }
    const actsEl = document.getElementById("thisPeriodActivitiesList");
    if (!acts.length) {
        actsEl.innerHTML = `<span class="blank-val">No current period activity bullets recorded</span>`;
    } else {
        actsEl.innerHTML = acts.map(a => `<span class="source-pill">• ${esc(a)}</span>`).join("");
    }
}

function renderScurveWeightsTable() {
    const projSc = (APP_STATE.snapshot && APP_STATE.snapshot.project_overall_scurve) || {};
    const rows = projSc.package_weights_breakdown || [];
    const q = (document.getElementById("scurveWeightsFilterInput")?.value || "").trim().toUpperCase();

    const noteEl = document.getElementById("pkgWeightsFormulaNote");
    if (noteEl) {
        noteEl.textContent = `Total Project Weighted Planned: ${Number(projSc.planned_to_date_pct || 0).toFixed(2)}% • Total Project Weighted Actual: ${Number(projSc.actual_to_date_pct || 0).toFixed(2)}% (Sum across ${rows.length} packages = 100% weight)`;
    }

    const filtered = rows.filter(r => {
        if (!q) return true;
        return `${r.package_code} ${r.package_name} ${r.discipline}`.toUpperCase().includes(q);
    });

    const sorted = sortRows(filtered, "scurveWeights");
    const tbody = document.getElementById("scurveWeightsTableBody");
    if (!tbody) return;

    tbody.innerHTML = sorted.map(r => {
        const isSel = r.package_code === APP_STATE.selectedPkgCode;
        return `
            <tr class="clickable-row ${isSel ? "selected-row" : ""}" onclick="selectPackage('${esc(r.package_code)}')">
                <td><strong>${esc(r.package_code)}</strong></td>
                <td>${esc(r.package_name)}</td>
                <td class="center"><span class="badge badge-info">${esc(r.discipline)}</span></td>
                <td class="num">${Number(r.package_weight_pct || 0).toFixed(2)}%</td>
                <td class="num">${Number(r.pkg_planned_pct || 0).toFixed(1)}%</td>
                <td class="num"><strong>${Number(r.pkg_actual_pct || 0).toFixed(1)}%</strong></td>
                <td class="num">${Number(r.weighted_planned_pct || 0).toFixed(2)}%</td>
                <td class="num"><strong>${Number(r.weighted_actual_pct || 0).toFixed(2)}%</strong></td>
            </tr>
        `;
    }).join("");
}

/* ==========================================================================
   ELEMENT 8: LOOK AHEAD PLAN (7–14 DAYS) — DYNAMIC PER PACKAGE
   ========================================================================== */
function renderLookAheadPlan(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const titleEl = document.getElementById("lookAheadCardTitle");
    if (titleEl) {
        titleEl.textContent = `Look Ahead Plan (7–14 Days) — [${pkg.package_code}]`;
    }

    const la = pkg.look_ahead_plan || {};
    const allItems = [
        ...(la.wpr_next_activities || []),
        ...(la.schedule_tasks_7_14d || []),
        ...(la.vdrl_submissions_7_14d || []),
        ...(la.upcoming_milestones || []),
    ];

    const q = (document.getElementById("lookAheadFilterInput")?.value || "").trim().toUpperCase();
    const filtered = allItems.filter(item => {
        if (!q) return true;
        return `${item.title} ${item.category} ${item.target_date} ${item.source}`.toUpperCase().includes(q);
    });

    document.getElementById("lookAheadCountBadge").textContent = `${filtered.length} Activity Item(s)`;
    const container = document.getElementById("lookAheadContainer");

    if (!filtered.length) {
        container.innerHTML = `<div class="alert-item sev-lookahead"><div class="alert-main"><div class="alert-title">No matching 7–14 day actions found</div><div class="alert-detail">Clear the filter or check upcoming milestones for ${esc(pkg.package_code)}.</div></div></div>`;
        return;
    }

    container.innerHTML = filtered.slice(0, 10).map(item => `
        <div class="alert-item sev-lookahead">
            <div class="alert-main">
                <div class="alert-title">${esc(item.title)}</div>
                <div class="alert-detail">${esc(item.category)} • Target Window: <strong>${esc(item.target_date)}</strong></div>
            </div>
            <span class="badge badge-info">${esc(item.source)}</span>
        </div>
    `).join("");
}

/* ==========================================================================
   ELEMENT 9: AREA OF CONCERN — DYNAMIC PER PACKAGE (Includes ROS Delay)
   ========================================================================== */
function renderAreasOfConcern(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const concerns = pkg.areas_of_concern || [];
    const q = (document.getElementById("concernFilterInput")?.value || "").trim().toUpperCase();
    const filtered = concerns.filter(c => {
        if (!q) return true;
        return `${c.category} ${c.title} ${c.detail} ${c.source}`.toUpperCase().includes(q);
    });

    document.getElementById("concernCountBadge").textContent = `${filtered.length} Concern(s)`;
    const container = document.getElementById("areaOfConcernContainer");

    if (!filtered.length) {
        container.innerHTML = `
            <div class="alert-item">
                <div class="alert-main">
                    <div class="alert-title">No Critical Concerns or Overdue Bottlenecks Match Filter</div>
                    <div class="alert-detail">Package milestones, ROS compliance, VDRL submissions, and Client review cycles are clear.</div>
                </div>
                <span class="badge badge-success">Clear</span>
            </div>
        `;
        return;
    }

    container.innerHTML = filtered.map(c => {
        const sevClass = c.severity === "high" ? "sev-high" : "sev-medium";
        const badgeTone = c.severity === "high" ? "badge-danger" : "badge-warning";
        return `
            <div class="alert-item ${sevClass}">
                <div class="alert-main">
                    <div class="alert-title">[${esc(c.category)}] ${esc(c.title)}</div>
                    <div class="alert-detail">${esc(c.detail)}</div>
                </div>
                <span class="badge ${badgeTone}">${esc(c.source)}</span>
            </div>
        `;
    }).join("");
}

/* ==========================================================================
   COMMENT #5: WPR ATTACHMENT 3 (GANTT) & ATTACHMENT 4 (SUPPLIER MATERIAL)
   ========================================================================== */
function renderAttachments3And4(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const att3 = pkg.attachment3_summary || {};
    const att4 = pkg.attachment4_summary || {};

    // --- Attachment 3 (Overall Fabrication Detail Plan — Gantt Chart Analysis) ---
    document.getElementById("att3Subtitle").textContent =
        `Critical Gate: ${att3.critical_gate || "—"}`;
    document.getElementById("att3TotalTasksBadge").textContent =
        `${att3.total_tasks || 61} Gantt Tasks (${att3.total_phases || 5} Phases)`;

    document.getElementById("att3SituationBanner").innerHTML = `
        <div class="alert-item sev-medium">
            <div class="alert-main">
                <div class="alert-title">📊 Attachment 3 Gantt Situation Analysis — [${esc(pkg.package_code)}]</div>
                <div class="alert-detail">${esc(att3.situation_summary || "—")}</div>
            </div>
            <span class="badge badge-info">Gantt Analysis</span>
        </div>
    `;

    const q3 = (document.getElementById("att3FilterInput")?.value || "").trim().toUpperCase();
    const phases = (att3.phases || []).filter(p => {
        if (!q3) return true;
        return `${p.phase_id} ${p.phase_name} ${p.start_date} ${p.finish_date} ${p.status} ${p.critical_note}`.toUpperCase().includes(q3);
    });
    const sortedPhases = sortRows(phases, "att3Phases");

    const att3Tbody = document.getElementById("att3PhasesTableBody");
    att3Tbody.innerHTML = sortedPhases.map(p => `
        <tr>
            <td><strong>${esc(p.phase_id)}</strong> — ${esc(p.phase_name)}</td>
            <td class="center"><span class="badge badge-muted">${esc(p.duration_days)}d</span></td>
            <td>${fmtBlank(p.start_date)} → ${fmtBlank(p.finish_date)}</td>
            <td class="center"><span class="badge ${badgeClassForTone(p.tone)}">${esc(p.status)}</span></td>
            <td>${esc(p.critical_note)}</td>
        </tr>
    `).join("");

    // --- Attachment 4 (Supplier Material Status Expediting Report) ---
    document.getElementById("att4Subtitle").textContent =
        `Target Material Delivery: ${att4.target_material_delivery_date || "—"} • Sub-vendor procurement holds & prerequisite gates`;
    document.getElementById("att4HoldCountBadge").textContent =
        `${att4.components_on_hold || 0} / ${att4.total_components || 4} Sub-Orders Held`;

    document.getElementById("att4SituationBanner").innerHTML = `
        <div class="alert-item ${(att4.components_on_hold || 0) > 0 ? "sev-high" : "sev-lookahead"}">
            <div class="alert-main">
                <div class="alert-title">📦 Attachment 4 Supplier Material Bottleneck Summary — [${esc(pkg.package_code)}]</div>
                <div class="alert-detail">${esc(att4.expediting_narrative || "—")}</div>
            </div>
            <span class="badge ${(att4.components_on_hold || 0) > 0 ? "badge-danger" : "badge-success"}">Sub-Order Gate</span>
        </div>
    `;

    const q4 = (document.getElementById("att4FilterInput")?.value || "").trim().toUpperCase();
    const matItems = (att4.items || []).filter(it => {
        if (!q4) return true;
        return `${it.item_no} ${it.component} ${it.sub_vendor_origin} ${it.prerequisite_gate} ${it.po_status} ${it.impact_on_production} ${it.risk_level}`.toUpperCase().includes(q4);
    });
    const sortedMat = sortRows(matItems, "att4Materials");

    const att4Tbody = document.getElementById("att4MaterialsTableBody");
    att4Tbody.innerHTML = sortedMat.map(it => `
        <tr>
            <td class="center"><strong>${esc(it.item_no)}</strong></td>
            <td>
                <strong>${esc(it.component)}</strong>
                <div style="font-size:11px;color:var(--text-muted);">${esc(it.sub_vendor_origin)} • Lead Time: ${esc(it.lead_time_days)}d</div>
            </td>
            <td><span class="badge ${badgeClassForTone(it.tone)}">${esc(it.prerequisite_gate)}</span></td>
            <td class="center"><strong>${esc(it.po_status)}</strong></td>
            <td>
                <div>Order: ${fmtBlank(it.planned_order_date)}</div>
                <div style="font-size:11px;color:var(--text-muted);">ETA: ${fmtBlank(it.planned_eta)}</div>
            </td>
            <td class="center"><span class="badge ${badgeClassForTone(it.tone)}">${esc(it.risk_level)}</span></td>
            <td>${esc(it.impact_on_production)}</td>
        </tr>
    `).join("");
}

/* ==========================================================================
   ELEMENT 4: KEY MILESTONES TABLE (WPR PRIORITY #1 > PSR PRIORITY #2)
   ========================================================================== */
function renderKeyMilestonesTable(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const rows = (pkg.key_milestones || []).map(m => ({
        ...m,
        effective_date: m.actual_date || m.forecast_date || "",
    }));

    document.getElementById("milestoneSourceBadge").textContent = pkg.has_live_wpr
        ? `Priority #1: Live ${pkg.report_no || "WPR"} (Cut-off: ${pkg.cut_off_date || "—"})`
        : "Priority #2: SCD & PSR Baseline Only (No WPR)";

    const discBannerEl = document.getElementById("element4DiscrepancyBanner");
    if (discBannerEl) {
        if (pkg.has_milestone_discrepancy) {
            discBannerEl.style.display = "block";
            discBannerEl.innerHTML = `
                <div class="alert-item sev-medium">
                    <div class="alert-main">
                        <div class="alert-title">⚠️ Milestone #10 vs. Gantt Chart ETA Discrepancy (+${esc(pkg.milestone_discrepancy_days)} Days)</div>
                        <div class="alert-detail">${esc(pkg.milestone_discrepancy_note)}</div>
                    </div>
                    <span class="badge badge-warning">Governing ETA: ${esc(pkg.gantt_eta_date)}</span>
                </div>
            `;
        } else {
            discBannerEl.style.display = "none";
            discBannerEl.innerHTML = "";
        }
    }

    const q = (document.getElementById("milestoneTableFilterInput")?.value || "").trim().toUpperCase();
    const filtered = rows.filter(m => {
        if (!q) return true;
        return `${m.no} ${m.description} ${m.planned_date} ${m.effective_date} ${m.status} ${m.source} ${m.remark} ${m.deviation_note}`.toUpperCase().includes(q);
    });

    const sorted = sortRows(filtered, "keyMilestones");
    const tbody = document.getElementById("keyMilestonesTableBody");

    tbody.innerHTML = sorted.map(m => {
        const fcOrAct = m.actual_date
            ? `<strong style="color:var(--accent-emerald);">Act: ${esc(m.actual_date)}</strong>`
            : (m.forecast_date ? `<span style="color:var(--accent-amber);font-weight:600;">Fcst: ${esc(m.forecast_date)}</span>` : `<span class="blank-val">—</span>`);

        const notes = [m.remark, m.deviation_note].filter(Boolean).join(" • ");
        const isRosRow = m.no === 12 || m.no === 11;
        const isDanger = m.tone === "danger";
        return `
            <tr class="${isDanger ? "ros-delay-row" : ""}">
                <td class="center"><strong>${m.no}</strong></td>
                <td>
                    <strong>${esc(m.description)}</strong>
                    ${isRosRow && m.no === 12 ? ` <span class="badge badge-purple">🔒 Immutable PSR ROS</span>` : ""}
                </td>
                <td>${fmtBlank(m.planned_date)}</td>
                <td>${fcOrAct}</td>
                <td class="center"><span class="badge ${badgeClassForTone(m.tone)}">${esc(m.status)}</span></td>
                <td>
                    <span class="badge badge-muted" style="margin-right:4px;">${esc(m.source)}</span>
                    ${notes ? `<span style="font-size:11.5px;color:${isDanger ? "var(--accent-rose)" : "var(--text-secondary)"};font-weight:${isDanger ? "700" : "400"};">${esc(notes)}</span>` : ""}
                </td>
            </tr>
        `;
    }).join("");
}

/* ==========================================================================
   ELEMENT 5: PROGRESS TRACKING BY KEY MILESTONE
   ========================================================================== */
function renderMilestoneProgressTable(pkg) {
    if (!pkg) pkg = getSelectedPackage();
    if (!pkg) return;

    const rows = pkg.milestone_progress || [];
    const ov = pkg.overall_progress || {};
    const tbody = document.getElementById("milestoneProgressTableBody");

    document.getElementById("milestoneProgressSummaryBadge").textContent =
        `Planned: ${Number(ov.planned_pct || 0).toFixed(1)}% • Actual: ${Number(ov.actual_pct || 0).toFixed(1)}%`;

    const q = (document.getElementById("milestoneProgFilterInput")?.value || "").trim().toUpperCase();
    const filtered = rows.filter(r => {
        if (!q) return true;
        return `${r.id} ${r.description} ${r.weight_pct} ${r.planned_pct} ${r.actual_pct}`.toUpperCase().includes(q);
    });

    const sorted = sortRows(filtered, "milestoneProgress");

    if (!sorted.length) {
        const msg = (!pkg.has_live_wpr && APP_STATE.mode === "live")
            ? `ℹ️ No WPR uploaded for [${esc(pkg.package_code)}] ${esc(pkg.package_name)} — In Execution Mode, only available Master Schedule (SCD), PSR & VDRL data is shown.`
            : `No matching milestone progress rows for ${esc(pkg.package_code)}.`;
        tbody.innerHTML = `<tr><td colspan="7" class="center" style="padding:24px;color:var(--text-muted);">${msg}</td></tr>`;
        return;
    }

    tbody.innerHTML = sorted.map(r => {
        const actPct = Number(r.actual_pct || 0);
        const planPct = Number(r.planned_pct || 0);
        const fillClass = actPct < planPct ? "warn" : "";
        return `
            <tr class="${r.is_sub ? "sub-row" : ""}">
                <td class="center"><strong>${esc(r.id)}</strong></td>
                <td>${esc(r.description)}</td>
                <td class="num">${Number(r.weight_pct || 0).toFixed(1)}%</td>
                <td class="num">${planPct.toFixed(1)}%</td>
                <td class="num"><strong>${actPct.toFixed(1)}%</strong></td>
                <td class="num">${Number(r.weighted_actual_pct || 0).toFixed(2)}%</td>
                <td>
                    <div class="progress-track">
                        <div class="progress-fill ${fillClass}" style="width:${Math.min(actPct, 100)}%;"></div>
                    </div>
                </td>
            </tr>
        `;
    }).join("");
}

/* ==========================================================================
   ELEMENT 6: DUAL STATUS VENDOR DOCUMENT (6A VDRL + 6B CLIENT LOG)
   ========================================================================== */
function renderPackageDocuments() {
    const pkg = getSelectedPackage();
    if (!pkg) return;

    const vdrl = pkg.vdrl_status || {};
    const clog = pkg.client_doc_status || {};
    const q = (document.getElementById("pkgDocSearchInput")?.value || "").trim().toUpperCase();
    const q6A = (document.getElementById("vdrlSubFilterInput")?.value || "").trim().toUpperCase();
    const q6B = (document.getElementById("clientDocSubFilterInput")?.value || "").trim().toUpperCase();
    const filterMode = document.getElementById("pkgDocFilterSelect")?.value || "ALL";

    // 6A Summary Cards
    document.getElementById("vdrlFileSubtitle").textContent = vdrl.available
        ? `File: ${vdrl.source_file} • Status Date: ${vdrl.today_date || "—"}`
        : "No VDRL workbook in VDRL/ for this package";
    document.getElementById("vdrlPctBadge").textContent = `${Number(vdrl.pct_submitted || 0).toFixed(1)}% Submitted (${vdrl.submitted || 0}/${vdrl.total_docs || 0})`;

    document.getElementById("vdrlSummaryStatsGrid").innerHTML = `
        <div class="kv-box">
            <div class="kv-label">Total VDRL Docs</div>
            <div class="kv-val">${vdrl.total_docs || 0} (Key: ${vdrl.key_docs_total || 0} / Non-Key: ${vdrl.non_key_total || 0})</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Submitted / Not Yet</div>
            <div class="kv-val"><span style="color:var(--accent-emerald);">${vdrl.submitted || 0} Submitted</span> • ${vdrl.not_yet || 0} Not Yet</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Overdue Submissions</div>
            <div class="kv-val"><span style="color:${(vdrl.overdue || 0) > 0 ? "var(--accent-rose)" : "var(--accent-emerald)"};">${vdrl.overdue || 0} Overdue</span> (${vdrl.key_docs_overdue || 0} Key)</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Production / Fab Impact</div>
            <div class="kv-val">Prod Overdue: ${vdrl.production_impact_overdue || 0}/${vdrl.production_impact_total || 0} • Fab Overdue: ${vdrl.fabrication_impact_overdue || 0}/${vdrl.fabrication_impact_total || 0}</div>
        </div>
    `;

    // 6B Summary Cards
    const logFile = (APP_STATE.snapshot.sources_summary && APP_STATE.snapshot.sources_summary.client_doc_log_file) || "ASK Vendor Document Log Report";
    document.getElementById("clientLogFileSubtitle").textContent = clog.available
        ? `Source: ${logFile} (Package Code: ${clog.raw_package_code || pkg.package_code})`
        : `No submittals logged under '${pkg.package_code}' in ${logFile} yet`;
    document.getElementById("clientLogTotalBadge").textContent = `${clog.total_submittals || 0} Contractual Submittal(s)`;

    document.getElementById("clientLogSummaryStatsGrid").innerHTML = `
        <div class="kv-box">
            <div class="kv-label">Workflow Pipeline</div>
            <div class="kv-val">In Review: ${clog.in_review || 0} • Submitted: ${clog.submitted || 0} • In Approval: ${clog.in_approval || 0}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Client Return Codes</div>
            <div class="kv-val"><span style="color:var(--accent-emerald);">APPR: ${clog.appr || 0}</span> • <span style="color:var(--accent-amber);">AWC: ${clog.awc || 0}</span> • REC: ${clog.rec || 0}</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Pending Return</div>
            <div class="kv-val">${clog.pending_return || 0} Submittal(s) Awaiting Client Return</div>
        </div>
        <div class="kv-box">
            <div class="kv-label">Overdue Client Reviews</div>
            <div class="kv-val"><span style="color:${(clog.overdue_reviews || 0) > 0 ? "var(--accent-rose)" : "var(--accent-emerald)"};">${clog.overdue_reviews || 0} Past Plan Finish Date</span></div>
        </div>
    `;

    // 6A Table: VDRL Documents (Filtered + Sortable)
    const vdrlDocs = (vdrl.documents || []).filter(d => {
        if (filterMode === "OVERDUE" && !String(d.status).includes("OVERDUE")) return false;
        if (filterMode === "KEY" && d.is_key !== "Y") return false;
        if (filterMode === "PROD_IMPACT" && d.impact_production !== "Y") return false;
        if (filterMode === "SUBMITTED" && !String(d.status).includes("SUBMIT")) return false;
        const hay = `${d.doc_no} ${d.title} ${d.status} ${d.doc_type} ${d.plan_date} ${d.actual_date}`.toUpperCase();
        if (q && !hay.includes(q)) return false;
        if (q6A && !hay.includes(q6A)) return false;
        return true;
    });

    const sortedVdrl = sortRows(vdrlDocs, "pkgVdrlDocs");
    const vdrlTbody = document.getElementById("pkgVdrlDocsTableBody");
    if (!sortedVdrl.length) {
        vdrlTbody.innerHTML = `<tr><td colspan="6" class="center" style="padding:18px;color:var(--text-muted);">No matching VDRL documents found.</td></tr>`;
    } else {
        vdrlTbody.innerHTML = sortedVdrl.map(d => {
            const st = String(d.status || "").toUpperCase();
            const stBadge = st.includes("SUBMIT") ? "badge-success" : (st.includes("OVERDUE") ? "badge-danger" : "badge-muted");
            return `
                <tr>
                    <td><strong>${esc(d.doc_no)}</strong></td>
                    <td>${esc(d.title)}</td>
                    <td class="center">
                        <span class="badge ${d.is_key === "Y" ? "badge-info" : "badge-muted"}">Key:${esc(d.is_key)}</span>
                        ${d.impact_production === "Y" ? `<span class="badge badge-warning">Prod:Y</span>` : ""}
                    </td>
                    <td>${fmtBlank(d.plan_date)}</td>
                    <td>${fmtBlank(d.actual_date)}</td>
                    <td class="center"><span class="badge ${stBadge}">${esc(st || "NOT YET")}</span></td>
                </tr>
            `;
        }).join("");
    }

    // 6B Table: Contractual Client Review Submittals (Filtered + Sortable)
    const clientDocs = (clog.documents || []).filter(d => {
        if (filterMode === "OVERDUE" && !d.is_overdue) return false;
        if (filterMode === "SUBMITTED" && d.status !== "Submitted") return false;
        const hay = `${d.submittal_id} ${d.doc_no} ${d.title} ${d.status} ${d.return_code} ${d.focal_point} ${d.transmittal_out_no}`.toUpperCase();
        if (q && !hay.includes(q)) return false;
        if (q6B && !hay.includes(q6B)) return false;
        return true;
    });

    const sortedClient = sortRows(clientDocs, "pkgClientDocs");
    const clientTbody = document.getElementById("pkgClientDocsTableBody");
    if (!sortedClient.length) {
        clientTbody.innerHTML = `<tr><td colspan="6" class="center" style="padding:18px;color:var(--text-muted);">No contractual submittals in Vendor Document Log Report for this filter/package. (See Tab 3 for all 271 project submittals).</td></tr>`;
    } else {
        clientTbody.innerHTML = sortedClient.map(d => {
            const rc = d.return_code || "PENDING";
            const rcBadge = rc === "APPR" ? "badge-success" : (rc === "AWC" ? "badge-warning" : (d.is_overdue ? "badge-danger" : "badge-info"));
            return `
                <tr>
                    <td><strong>${esc(d.submittal_id)}</strong><br/><span style="font-size:11px;color:var(--text-muted);">Rev.${esc(d.rev)} (${esc(d.purpose)})</span></td>
                    <td><strong>${esc(d.doc_no)}</strong><br/><span style="font-size:11.5px;color:var(--text-secondary);">${esc(d.title)}</span></td>
                    <td class="center"><span class="badge badge-purple">${esc(d.status)}</span></td>
                    <td>
                        <div style="font-size:11.5px;">FP: ${fmtBlank(d.focal_point)}</div>
                        <div style="font-size:11px;color:var(--text-muted);">Plan: ${fmtBlank(d.focal_point_plan_finish || d.reviewer_plan_finish)}</div>
                    </td>
                    <td class="center"><span class="badge ${rcBadge}">${esc(rc)}</span></td>
                    <td>
                        <div>${fmtBlank(d.transmittal_out_no)}</div>
                        <div style="font-size:11px;color:var(--text-muted);">${fmtBlank(d.transmittal_out_date)}</div>
                    </td>
                </tr>
            `;
        }).join("");
    }
}

/* ==========================================================================
   TAB 2: EXECUTIVE PORTFOLIO MATRIX RENDERER (SINGLE-WINDOW FIT, NO SLIDING BAR)
   ========================================================================== */
function renderPortfolioTable() {
    const pkgs = (APP_STATE.snapshot && APP_STATE.snapshot.packages) || [];
    const q = (document.getElementById("portfolioSearchInput")?.value || "").trim().toUpperCase();
    const covFilter = document.getElementById("portfolioCoverageFilter")?.value || "ALL";
    const discFilter = document.getElementById("portfolioDiscFilter")?.value || "ALL";
    const healthFilter = document.getElementById("portfolioHealthFilter")?.value || "ALL";

    const filtered = pkgs.filter(p => {
        const hl = p.package_highlights || {};
        const isCore = Boolean(p.is_core_package !== undefined ? p.is_core_package : ((p.vdrl_status || {}).available || (p.manufacturer_detail || {}).available));
        if (covFilter === "CORE" && !isCore) return false;
        if (covFilter === "PSR_ONLY" && isCore) return false;
        if (discFilter !== "ALL" && p.discipline !== discFilter) return false;
        if (healthFilter === "ROS_DELAY" && !hl.is_ros_delayed) return false;
        if (healthFilter !== "ALL" && healthFilter !== "ROS_DELAY" && p.health_tone !== healthFilter) return false;
        if (q) {
            const mfrName = (p.manufacturer_detail && p.manufacturer_detail.manufacturer_name) || "";
            const hay = `${p.package_code} ${p.package_name} ${p.discipline} ${mfrName} ${p.po_number} ${p.rfq_no || ""} ${p.equipment_tag} ${hl.ros_date} ${hl.ros_status_label}`.toUpperCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });

    const extractorMap = {
        package_code: p => p.package_code,
        package_name: p => p.package_name,
        discipline: p => p.discipline,
        manufacturer_name: p => (p.manufacturer_detail && p.manufacturer_detail.manufacturer_name) || "",
        package_weight_pct: p => Number(p.package_weight_pct || 0),
        planned_pct: p => Number((p.overall_progress || {}).planned_pct || 0),
        actual_pct: p => Number((p.overall_progress || {}).actual_pct || 0),
        next_milestone_days: p => (p.package_highlights || {}).remaining_days_to_next_milestone,
        days_to_delivery: p => (p.package_highlights || {}).days_to_delivery,
        ros_margin_days: p => (p.package_highlights || {}).ros_margin_days,
        vdrl_pct: p => Number((p.vdrl_status || {}).pct_submitted || 0),
        concerns_count: p => (p.areas_of_concern || []).length,
        health_status: p => p.health_status,
    };

    const sorted = sortRows(filtered, "portfolio", extractorMap);
    const tbody = document.getElementById("portfolioTableBody");

    tbody.innerHTML = sorted.map(p => {
        const mfr = p.manufacturer_detail || {};
        const vdrl = p.vdrl_status || {};
        const ov = p.overall_progress || {};
        const hl = p.package_highlights || {};
        const isSelected = p.package_code === APP_STATE.selectedPkgCode;
        const isRosDelayed = Boolean(hl.is_ros_delayed);

        const remDays = hl.remaining_days_to_next_milestone;
        const remBadge = remDays !== null && remDays !== undefined
            ? (remDays < 0 ? `${Math.abs(remDays)}d Overdue` : `In ${remDays}d`)
            : "—";

        let compactRosBadge = "On Track";
        if (isRosDelayed) {
            compactRosBadge = `🚨 +${hl.ros_delay_days}d Late`;
        } else if (hl.ros_margin_days !== null && hl.ros_margin_days !== undefined) {
            compactRosBadge = hl.ros_margin_days <= 14 ? `⚠️ ${hl.ros_margin_days}d Buf.` : `✓ ${hl.ros_margin_days}d Buf.`;
        }

        let compactHealth = p.health_status || "On Track";
        if (compactHealth.includes("Critical")) compactHealth = "Critical";
        else if (compactHealth.includes("Watch")) compactHealth = "Watch";

        return `
            <tr class="clickable-row ${isSelected ? "selected-row" : ""} ${isRosDelayed ? "ros-delay-row" : ""}" onclick="selectPackage('${esc(p.package_code)}', true)">
                <td><strong>${esc(p.package_code)}</strong></td>
                <td>
                    <strong>${esc(p.package_name)}</strong>
                    <div style="font-size:10.5px;color:var(--text-muted);">${esc(p.rfq_no || p.equipment_tag || "—")}${p.po_number ? ` • ${esc(p.po_number)}` : ""}</div>
                </td>
                <td class="center"><span class="badge badge-info">${esc(p.discipline)}</span></td>
                <td>
                    ${mfr.available
                        ? `<strong>${esc(mfr.manufacturer_name)}</strong><div style="font-size:10.5px;color:var(--text-muted);">KOM: ${esc(mfr.kom_date || "—")}</div>`
                        : `<span class="blank-val">— (No KOM)</span>`}
                </td>
                <td class="num">${Number(p.package_weight_pct || 0).toFixed(2)}%</td>
                <td class="num">${Number(ov.planned_pct || 0).toFixed(1)}%</td>
                <td class="num"><strong>${Number(ov.actual_pct || 0).toFixed(1)}%</strong></td>
                <td>
                    <div style="font-size:11.5px;font-weight:600;">${esc(hl.next_milestone_name || "—")}</div>
                    <div style="font-size:10.5px;color:var(--text-muted);">${fmtBlank(hl.next_milestone_date)} (${esc(remBadge)})</div>
                </td>
                <td>
                    <div style="font-size:11.5px;font-weight:700;">${fmtBlank(hl.effective_delivery_date)}</div>
                    <div style="font-size:10.5px;color:var(--text-muted);">${hl.days_to_delivery !== null && hl.days_to_delivery !== undefined ? `${hl.days_to_delivery}d to ETA` : "—"}</div>
                </td>
                <td>
                    <div style="font-size:11.5px;"><strong>${fmtBlank(hl.ros_date)}</strong></div>
                    <div>
                        <span class="badge ${isRosDelayed ? "badge-danger" : (hl.ros_margin_days !== null && hl.ros_margin_days <= 14 ? "badge-warning" : "badge-success")}">
                            ${esc(compactRosBadge)}
                        </span>
                    </div>
                </td>
                <td class="num">
                    ${vdrl.available
                        ? `<strong>${vdrl.submitted}/${vdrl.total_docs}</strong> (${Number(vdrl.pct_submitted || 0).toFixed(0)}%)<div style="font-size:10.5px;color:${vdrl.overdue > 0 ? "var(--accent-rose)" : "var(--text-muted)"};">${vdrl.overdue} Overdue</div>`
                        : `<span class="blank-val">PSR Only</span>`}
                </td>
                <td class="center"><span class="badge ${(p.areas_of_concern || []).length > 0 ? "badge-warning" : "badge-success"}">${(p.areas_of_concern || []).length}</span></td>
                <td class="center"><span class="badge ${badgeClassForTone(p.health_tone)}">${esc(compactHealth)}</span></td>
            </tr>
        `;
    }).join("");
}

/* ==========================================================================
   TAB 3: CONTRACTUAL CLIENT REVIEW & VDRL HUB RENDERER (FILTERABLE + SORTABLE)
   ========================================================================== */
function renderDocControlHub() {
    const overview = (APP_STATE.snapshot && APP_STATE.snapshot.client_doc_log_overview) || {};
    document.getElementById("hubClientLogSubtitle").textContent = overview.available
        ? `File: ${overview.source_file} • Last Updated: ${overview.last_updated || "—"} • Total Submittals: ${overview.total_submittals}`
        : `No Vendor Document Log Report found for ${APP_STATE.project}`;

    const q = (document.getElementById("hubPkgBreakdownFilter")?.value || "").trim().toUpperCase();
    const pkgRows = (overview.package_breakdown || []).filter(r => {
        if (!q) return true;
        return `${r.raw_package_code} ${r.package_code} ${r.package_name}`.toUpperCase().includes(q);
    });

    const sortedPkgRows = sortRows(pkgRows, "hubClientPkg");
    const pkgTbody = document.getElementById("hubClientPkgTableBody");
    if (!sortedPkgRows.length) {
        pkgTbody.innerHTML = `<tr><td colspan="9" class="center" style="padding:20px;color:var(--text-muted);">No matching packages in Client Document Log Report.</td></tr>`;
    } else {
        pkgTbody.innerHTML = sortedPkgRows.map(r => `
            <tr class="clickable-row" onclick="selectPackage('${esc(r.package_code)}', true)">
                <td><strong>${esc(r.raw_package_code)}</strong> ${r.package_code !== r.raw_package_code ? `<span style="font-size:11px;color:var(--accent-blue);">(${esc(r.package_code)})</span>` : ""}</td>
                <td>${esc(r.package_name)}</td>
                <td class="num"><strong>${r.total_submittals}</strong></td>
                <td class="num">${r.in_review}</td>
                <td class="num">${r.submitted}</td>
                <td class="num">${r.in_approval}</td>
                <td class="num" style="color:var(--accent-emerald);font-weight:700;">${r.appr}</td>
                <td class="num" style="color:var(--accent-amber);font-weight:700;">${r.awc}</td>
                <td class="num" style="color:${r.overdue_reviews > 0 ? "var(--accent-rose)" : "var(--text-muted)"};font-weight:700;">${r.overdue_reviews}</td>
            </tr>
        `).join("");
    }

    renderHubSubmittals();
}

function renderHubSubmittals() {
    const overview = (APP_STATE.snapshot && APP_STATE.snapshot.client_doc_log_overview) || {};
    const rows = overview.recent_submittals || [];
    const q = (document.getElementById("hubSubmittalSearch")?.value || "").trim().toUpperCase();
    const retFilter = document.getElementById("hubReturnFilter")?.value || "ALL";

    const filtered = rows.filter(r => {
        if (retFilter === "PENDING" && r.return_code !== "PENDING") return false;
        if (retFilter === "APPR" && r.return_code !== "APPR") return false;
        if (retFilter === "AWC" && r.return_code !== "AWC") return false;
        if (retFilter === "OVERDUE" && !r.is_overdue) return false;
        if (q) {
            const hay = `${r.submittal_id} ${r.vendor_package_code} ${r.vendor_package_name} ${r.doc_no} ${r.title} ${r.status} ${r.focal_point}`.toUpperCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    });

    const sorted = sortRows(filtered, "hubSubmittals");
    const tbody = document.getElementById("hubSubmittalsTableBody");
    if (!sorted.length) {
        tbody.innerHTML = `<tr><td colspan="7" class="center" style="padding:20px;color:var(--text-muted);">No matching submittals found.</td></tr>`;
        return;
    }

    tbody.innerHTML = sorted.slice(0, 120).map(d => {
        const rc = d.return_code || "PENDING";
        const rcBadge = rc === "APPR" ? "badge-success" : (rc === "AWC" ? "badge-warning" : (d.is_overdue ? "badge-danger" : "badge-info"));
        return `
            <tr>
                <td><strong>${esc(d.submittal_id)}</strong><div style="font-size:11px;color:var(--text-muted);">In: ${fmtBlank(d.actual_submit_in)}</div></td>
                <td><span class="badge badge-info">${esc(d.vendor_package_code)}</span></td>
                <td><strong>${esc(d.doc_no)}</strong><div style="font-size:11.5px;color:var(--text-secondary);">${esc(d.title)}</div></td>
                <td class="center"><span class="badge badge-purple">${esc(d.status)}</span></td>
                <td>
                    <div style="font-size:11.5px;">${fmtBlank(d.focal_point)}</div>
                    <div style="font-size:11px;color:var(--text-muted);">Plan: ${fmtBlank(d.focal_point_plan_finish || d.reviewer_plan_finish)}</div>
                </td>
                <td class="center"><span class="badge ${rcBadge}">${esc(rc)}</span></td>
                <td>
                    <div>${fmtBlank(d.transmittal_out_no)}</div>
                    <div style="font-size:11px;color:var(--text-muted);">${fmtBlank(d.transmittal_out_date)}</div>
                </td>
            </tr>
        `;
    }).join("");
}

/* ==========================================================================
   WEEKLY DYNAMIC FILE UPLOAD MODAL (WPR, VDRL, VDLR, PSR, KOM)
   ========================================================================== */
function openUploadModal() {
    document.getElementById("uploadStatusBox").textContent = "";
    document.getElementById("uploadModalBackdrop").classList.add("open");
}

function closeUploadModal() {
    document.getElementById("uploadModalBackdrop").classList.remove("open");
}

function setupDragAndDrop() {
    const dz = document.getElementById("dropzoneArea");
    if (!dz) return;
    dz.addEventListener("dragover", (e) => {
        e.preventDefault();
        dz.classList.add("dragover");
    });
    dz.addEventListener("dragleave", () => dz.classList.remove("dragover"));
    dz.addEventListener("drop", (e) => {
        e.preventDefault();
        dz.classList.remove("dragover");
        if (e.dataTransfer && e.dataTransfer.files.length) {
            handleFilesSelected(e.dataTransfer.files);
        }
    });
}

async function handleFilesSelected(fileList) {
    const statusBox = document.getElementById("uploadStatusBox");
    const docType = document.getElementById("uploadDocType").value || "WPR";
    const files = Array.from(fileList || []);
    if (!files.length) return;

    statusBox.innerHTML = `Uploading & extracting ${files.length} file(s)...`;

    const results = [];
    for (const file of files) {
        const b64 = await readFileAsBase64(file);
        const resp = await fetch("/api/upload", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                doc_type: docType,
                filename: file.name,
                content_base64: b64,
            }),
        });
        const resJson = await resp.json();
        if (!resp.ok) {
            results.push(`❌ ${esc(file.name)}: ${esc(resJson.error || "Upload failed")}`);
        } else {
            let msg = `✅ Saved <strong>${esc(resJson.filename)}</strong> → <code>${esc(resJson.saved_to)}</code>`;
            if (resJson.extraction) {
                const ext = resJson.extraction;
                if (ext.tasks_count !== undefined) {
                    msg += `<br/>&nbsp;&nbsp;↳ 📊 Extracted <strong>${esc(ext.package)}</strong> (${esc(ext.report_no)}): <strong>${ext.tasks_count} Gantt tasks</strong>, <strong>${ext.scurve_count} S-curve weeks</strong>, <strong>${ext.sub_orders_count} sub-orders</strong> synced to <em>Export Data File.xlsx</em> (Plan ${ext.plan_pct}%, Actual ${ext.actual_pct}%)`;
                } else if (ext.action_items_count !== undefined) {
                    msg += `<br/>&nbsp;&nbsp;↳ 📋 Extracted <strong>${esc(ext.package)}</strong> (${esc(ext.mom_no)}): <strong>${ext.action_items_count} action items</strong>, <strong>${ext.procurement_count} procurement items</strong> synced to <em>Export Data File.xlsx</em>`;
                }
            }
            results.push(msg);
        }
    }

    statusBox.innerHTML = results.join("<br/>");
    APP_STATE.lastUploadBannerMsg = results.join(" • ");
    await refreshDashboard(true);
}

async function triggerSyncAll() {
    const statusBox = document.getElementById("uploadStatusBox");
    statusBox.innerHTML = `⏳ Scanning and extracting reports into Export Data File_ASK.xlsx and Export Data File_Z1F.xlsx...`;
    try {
        const resp = await fetch("/api/sync-all", { method: "POST" });
        const resJson = await resp.json();
        if (resp.ok) {
            const r = resJson.results || {};
            const ask = r.ASK || {};
            const z1f = r.Z1F || {};
            statusBox.innerHTML = `✅ <strong>Sync Complete for All Projects!</strong><br/>
                • <strong>ASK (${ask.export_filename || 'Export Data File_ASK.xlsx'})</strong>: ${ask.kom_count || 0} KOM, ${ask.wpr_count || 0} WPR, ${ask.mom_count || 0} MOM, ${ask.lookahead_7d_count || 0} 7d Lookahead (${ask.file_size_kb || 0} KB)<br/>
                • <strong>Z1F (${z1f.export_filename || 'Export Data File_Z1F.xlsx'})</strong>: ${z1f.kom_count || 0} KOM, ${z1f.wpr_count || 0} WPR, ${z1f.mom_count || 0} MOM, ${z1f.lookahead_7d_count || 0} 7d Lookahead (${z1f.file_size_kb || 0} KB)`;
            APP_STATE.lastUploadBannerMsg = `Synced ASK (${ask.wpr_count || 0} WPR, ${ask.mom_count || 0} MOM) & Z1F (${z1f.wpr_count || 0} WPR, ${z1f.mom_count || 0} MOM) into Export Data Files.`;
            await refreshDashboard(true);
        } else {
            statusBox.innerHTML = `❌ Sync failed: ${esc(resJson.error || "Unknown error")}`;
        }
    } catch (e) {
        statusBox.innerHTML = `❌ Error: ${esc(e.message)}`;
    }
}

function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const result = String(reader.result || "");
            const commaIdx = result.indexOf(",");
            resolve(commaIdx >= 0 ? result.slice(commaIdx + 1) : result);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

/* ==========================================================================
   TAB 4: 7 & 14 DAYS LOOK-AHEAD HUB (Master Schedule & Packages)
   ========================================================================== */
function setLookAheadWindow(days) {
    APP_STATE.lookaheadWindow = days;
    const btn7 = document.getElementById("btnWindow7d");
    const btn14 = document.getElementById("btnWindow14d");
    if (btn7) btn7.classList.toggle("active", days === 7);
    if (btn14) btn14.classList.toggle("active", days === 14);

    const b = document.getElementById("lookAheadWindowBadge");
    if (b) {
        b.textContent = days === 7 ? "7-Day Look-Ahead Window" : "14-Day Look-Ahead Window";
        b.className = days === 7 ? "badge badge-info" : "badge badge-purple";
    }
    renderLookAheadHub();
}

function renderLookAheadHub() {
    const hub = (APP_STATE.snapshot && APP_STATE.snapshot.lookahead_hub) || {};
    const sum7 = hub.summary_7d || { total: 0, master_schedule_count: 0, package_count: 0, critical_path_count: 0 };
    const sum14 = hub.summary_14d || { total: 0, master_schedule_count: 0, package_count: 0, critical_path_count: 0 };

    const el7Total = document.getElementById("kpiLookAhead7dTotal");
    if (el7Total) el7Total.textContent = sum7.total || 0;
    const el7Sub = document.getElementById("kpiLookAhead7dSub");
    if (el7Sub) el7Sub.textContent = `${sum7.master_schedule_count || 0} Master Plan • ${sum7.package_count || 0} Package`;

    const el14Total = document.getElementById("kpiLookAhead14dTotal");
    if (el14Total) el14Total.textContent = sum14.total || 0;
    const el14Sub = document.getElementById("kpiLookAhead14dSub");
    if (el14Sub) el14Sub.textContent = `${sum14.master_schedule_count || 0} Master Plan • ${sum14.package_count || 0} Package`;

    const activeSum = APP_STATE.lookaheadWindow === 7 ? sum7 : sum14;
    const elCrit = document.getElementById("kpiLookAheadCritCount");
    if (elCrit) elCrit.textContent = activeSum.critical_path_count || 0;

    const elPkg = document.getElementById("kpiLookAheadPkgCount");
    if (elPkg) elPkg.textContent = activeSum.package_count || 0;

    // Resolve active project Master Schedule document reference dynamically
    const isZ1F = (APP_STATE.snapshot && APP_STATE.snapshot.project === "Z1F") || APP_STATE.project === "Z1F";
    const src = (APP_STATE.snapshot && APP_STATE.snapshot.sources_summary) || {};
    const masterDoc = (src.master_schedule_file || (isZ1F ? "MM-ZTK-1F-GEN-PJM-SCD-0002.pdf" : "MM-ASK-1A-GEN01-PJM-SCD-0001.pdf")).replace(/\.pdf$/i, "");
    const masterRev = src.master_schedule_rev || (isZ1F ? "C1" : "B1");

    // Dynamic Scope Filter Option alignment with selected project
    const scopeSel = document.getElementById("lookAheadScopeFilter");
    if (scopeSel) {
        const optMaster = scopeSel.querySelector('option[value="MASTER"]');
        if (optMaster) {
            optMaster.textContent = `Master Schedule Only (${masterDoc})`;
        }
    }

    const sub = document.getElementById("lookAheadSubtitle");
    if (sub) {
        const cutOff = hub.cut_off_date || (isZ1F ? "2026-09-16" : "2026-10-16");
        const endDt = APP_STATE.lookaheadWindow === 7 ? (hub.window_7d_end || "") : (hub.window_14d_end || "");
        sub.textContent = `Cut-off: ${cutOff}${endDt ? ` → Window Target: ${endDt}` : ""} • Integrated from ${masterDoc} (Rev. ${masterRev}) & Package Schedules`;
    }

    renderLookAheadTable();
}

function renderLookAheadTable() {
    const hub = (APP_STATE.snapshot && APP_STATE.snapshot.lookahead_hub) || {};
    const rawActs = (APP_STATE.lookaheadWindow === 7 ? hub.activities_7d : hub.activities_14d) || [];

    const scopeFilter = (document.getElementById("lookAheadScopeFilter") || {}).value || "ALL";
    const critFilter = (document.getElementById("lookAheadCritFilter") || {}).value || "ALL";
    const catFilter = (document.getElementById("lookAheadCatFilter") || {}).value || "ALL";
    const searchVal = ((document.getElementById("lookAheadSearchInput") || {}).value || "").toLowerCase().trim();

    let filtered = rawActs.filter(a => {
        if (scopeFilter === "MASTER" && a.scope !== "Master Schedule") return false;
        if (scopeFilter === "PACKAGE" && a.scope !== "Package") return false;
        if (critFilter === "CRITICAL_ONLY" && a.critical_path !== "YES") return false;
        if (catFilter !== "ALL" && !String(a.category || "").toLowerCase().includes(catFilter.toLowerCase())) return false;
        if (searchVal) {
            const haystack = `${a.package_or_id} ${a.package_name} ${a.activity_name} ${a.activity_id} ${a.source} ${a.category}`.toLowerCase();
            if (!haystack.includes(searchVal)) return false;
        }
        return true;
    });

    const sortConf = APP_STATE.sortState.lookahead || { col: "start_date", dir: "asc" };
    filtered.sort((x, y) => {
        let v1 = x[sortConf.col];
        let v2 = y[sortConf.col];
        if (v1 === null || v1 === undefined) v1 = "";
        if (v2 === null || v2 === undefined) v2 = "";
        if (typeof v1 === "string") {
            const cmp = v1.localeCompare(String(v2));
            return sortConf.dir === "asc" ? cmp : -cmp;
        }
        return sortConf.dir === "asc" ? v1 - v2 : v2 - v1;
    });

    const tbody = document.getElementById("lookAheadTableBody");
    if (!tbody) return;

    if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="center blank-val" style="padding:24px;">No activities found matching filters in this ${APP_STATE.lookaheadWindow}-day window.</td></tr>`;
        return;
    }

    tbody.innerHTML = filtered.map(a => {
        const isMaster = a.scope === "Master Schedule";
        const scopeBadge = isMaster
            ? `<span class="badge badge-purple" style="font-size:11px;">Master Plan</span>`
            : `<span class="badge badge-info" style="font-size:11px;">Package (${esc(a.package_or_id)})</span>`;
        const isCrit = a.critical_path === "YES";
        const critBadge = isCrit
            ? `<span class="badge badge-danger">🔴 TF=0 Critical</span>`
            : `<span class="badge badge-muted">Standard</span>`;
        const statusBadge = isCrit
            ? `<span class="badge badge-danger">${esc(a.status || 'Critical')}</span>`
            : `<span class="badge badge-success">${esc(a.status || 'Scheduled')}</span>`;

        return `<tr>
            <td class="center">${scopeBadge}</td>
            <td style="font-weight:700;color:var(--text-primary);font-family:monospace;font-size:12px;">${esc(a.package_or_id)}</td>
            <td>
                <div style="font-weight:600;color:var(--text-primary);">${esc(a.activity_name)}</div>
                <div style="font-size:11px;color:var(--text-muted);">${esc(a.package_name)}</div>
            </td>
            <td><span class="badge badge-info">${esc(a.category)}</span></td>
            <td class="center font-mono">${fmtBlank(a.start_date)}</td>
            <td class="center font-mono" style="font-weight:700;">${fmtBlank(a.finish_date)}</td>
            <td class="center">${critBadge}</td>
            <td style="font-size:11.5px;color:var(--text-secondary);">${esc(a.source)}</td>
            <td class="center">${statusBadge}</td>
        </tr>`;
    }).join("");
}
