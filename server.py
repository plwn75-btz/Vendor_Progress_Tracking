"""Standalone HTTP Server for the ASK & Z1F Weekly Project Management & Vendor Package Monitoring Dashboard.
Follows Lesson_Learn.md and Handoff_Report.md:
- Zero-cache headers on all responses (Cache-Control: no-cache, no-store, must-revalidate)
- Dynamic mtime + size fingerprinting across WPR/, VDRL/, KOM/, PSR, and Vendor Document Log Report
- Supports dropping files into designated folders OR uploading directly via POST /api/upload
- Binds to PORT environment variable (default 8080) for local and Render.com hosting
"""

import base64
import datetime
import glob
import hashlib
import http.server
import importlib
import json
import os
import re
import socketserver
import threading
import urllib.error
import urllib.request
from urllib.parse import parse_qs, urlparse
from typing import Any, Dict, List, Tuple

import vendor_dashboard.models
import vendor_dashboard.ingestors
import vendor_dashboard.reconciler

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))

_CACHE_LOCK = threading.Lock()
_SNAPSHOT_CACHE: Dict[str, Tuple[Tuple[Any, ...], Dict[str, Any]]] = {}


def _compute_workspace_fingerprint() -> Tuple[Any, ...]:
    """Compute a fast (rel_path, mtime_ns, size) tuple across all 5 dynamic document locations + backend code."""
    entries = []
    watched_dirs = [
        WORKSPACE_DIR,
        os.path.join(WORKSPACE_DIR, "WPR"),
        os.path.join(WORKSPACE_DIR, "MOM"),
        os.path.join(WORKSPACE_DIR, "VDRL"),
        os.path.join(WORKSPACE_DIR, "KOM"),
        os.path.join(WORKSPACE_DIR, "VDLR"),
        os.path.join(WORKSPACE_DIR, "PSR"),
        os.path.join(WORKSPACE_DIR, "vendor_dashboard"),
    ]
    valid_exts = (".xlsx", ".xls", ".pdf", ".docx", ".pptx", ".py")
    for d in watched_dirs:
        if not os.path.isdir(d):
            continue
        for fname in sorted(os.listdir(d)):
            if fname.startswith("~$"):
                continue
            if not fname.lower().endswith(valid_exts):
                continue
            fpath = os.path.join(d, fname)
            if os.path.isfile(fpath):
                try:
                    st = os.stat(fpath)
                    entries.append((os.path.relpath(fpath, WORKSPACE_DIR), st.st_mtime_ns, st.st_ctime_ns, st.st_size))
                except OSError:
                    pass
    return tuple(entries)


def get_cached_snapshot(project: str = "ASK", mode: str = "live", force_refresh: bool = False) -> Dict[str, Any]:
    proj = project.upper() if project.upper() in ("ASK", "Z1F") else "ASK"
    eff_mode = mode.lower() if mode.lower() in ("sample", "live") else "live"
    cache_key = f"{proj}::{eff_mode}"
    fp = _compute_workspace_fingerprint()

    with _CACHE_LOCK:
        if force_refresh:
            _SNAPSHOT_CACHE.clear()
            vendor_dashboard.ingestors._INGEST_CACHE.clear()
        elif cache_key in _SNAPSHOT_CACHE:
            cached_fp, cached_data = _SNAPSHOT_CACHE[cache_key]
            if cached_fp == fp:
                return cached_data
            else:
                # Workspace files changed: clear parser cache to reload freshly modified files
                vendor_dashboard.ingestors._INGEST_CACHE.clear()

        # Hot-reload backend modules so running server processes always use latest reconciler logic
        importlib.reload(vendor_dashboard.models)
        importlib.reload(vendor_dashboard.ingestors)
        importlib.reload(vendor_dashboard.reconciler)
        snapshot = vendor_dashboard.reconciler.build_project_snapshot(WORKSPACE_DIR, project=proj, mode=eff_mode)
        _SNAPSHOT_CACHE[cache_key] = (fp, snapshot)
        return snapshot


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    """HTTP Request Handler with zero-cache headers and REST API routes."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, directory=WORKSPACE_DIR, **kwargs)

    def end_headers(self) -> None:
        # Lesson_Learn.md §2.8: Prevent stale browser caching on weekly data & CSS/JS updates
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()

    def _send_json(self, status_code: int, payload: Dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        route = parsed.path
        qs = parse_qs(parsed.query)

        if route == "/api/dashboard":
            project = qs.get("project", ["ASK"])[0]
            mode = qs.get("mode", ["live"])[0]
            force = qs.get("refresh", ["0"])[0] == "1"
            try:
                data = get_cached_snapshot(project=project, mode=mode, force_refresh=force)
                self._send_json(200, data)
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        if route == "/api/package":
            project = qs.get("project", ["ASK"])[0]
            mode = qs.get("mode", ["live"])[0]
            code = qs.get("code", [""])[0].upper()
            try:
                data = get_cached_snapshot(project=project, mode=mode)
                pkg = next((p for p in data.get("packages", []) if p["package_code"].upper() == code), None)
                if not pkg:
                    self._send_json(404, {"error": f"Package '{code}' not found in {project}"})
                else:
                    self._send_json(200, pkg)
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        if route == "/api/health":
            self._send_json(200, {"status": "ok", "workspace": WORKSPACE_DIR})
            return

        if route == "/api/github-status":
            token = os.environ.get("GITHUB_TOKEN", "").strip()
            repo = os.environ.get("GITHUB_REPO", "").strip()
            branch = os.environ.get("GITHUB_BRANCH", "main").strip() or "main"
            self._send_json(200, {
                "configured": bool(token and repo),
                "has_token": bool(token),
                "repo": repo,
                "branch": branch,
                "upload_url": f"https://github.com/{repo}/upload/{branch}" if repo else "",
            })
            return

        if route == "/api/download-export-file":
            query = parse_qs(parsed.query)
            proj = query.get("project", ["ASK"])[0].upper()
            if proj not in ("ASK", "Z1F"):
                proj = "ASK"
            from vendor_dashboard.ingestors import get_export_data_path, get_export_data_filename
            export_filename = get_export_data_filename(proj)
            export_path = get_export_data_path(WORKSPACE_DIR, project=proj, check_legacy=True)
            if os.path.isfile(export_path):
                with open(export_path, "rb") as f:
                    data = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{export_filename}"')
                self.send_header("Content-Length", str(len(data)))
                self._add_no_cache_headers()
                self.end_headers()
                self.wfile.write(data)
                return
            else:
                self._send_json(404, {"error": f"{export_filename} not found on server."})
                return

        return super().do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/sync-all":
            try:
                from export_dashboard_data import sync_project_export_data
                query = parse_qs(parsed.query)
                req_proj = query.get("project", ["ALL"])[0].upper()
                targets = ["ASK", "Z1F"] if req_proj == "ALL" else ([req_proj] if req_proj in ("ASK", "Z1F") else ["ASK", "Z1F"])
                results_summary = {}
                for p in targets:
                    results_summary[p] = sync_project_export_data(WORKSPACE_DIR, project=p)

                with _CACHE_LOCK:
                    _SNAPSHOT_CACHE.clear()
                    vendor_dashboard.ingestors._INGEST_CACHE.clear()

                self._send_json(200, {"status": "synced", "results": results_summary})
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        if parsed.path == "/api/upload":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                raw = self.rfile.read(length)
                payload = json.loads(raw.decode("utf-8"))

                doc_type = str(payload.get("doc_type", "WPR")).upper()
                filename = os.path.basename(str(payload.get("filename", "")).strip())
                content_b64 = str(payload.get("content_base64", ""))

                if not filename or not content_b64:
                    self._send_json(400, {"error": "Both filename and content_base64 are required."})
                    return

                # Smart Drop Location Routing (Supports explicit doc_type or auto-detection from filename)
                u = filename.upper()
                if doc_type == "WPR" or ("WPR" in u and u.endswith((".PDF", ".DOCX"))):
                    target_dir = os.path.join(WORKSPACE_DIR, "WPR")
                    doc_type = "WPR"
                elif doc_type == "MOM" or ("MOM" in u and u.endswith(".PDF")):
                    target_dir = os.path.join(WORKSPACE_DIR, "MOM")
                    doc_type = "MOM"
                elif doc_type == "VDRL" or ("VDRL" in u and u.endswith(".XLSX")):
                    target_dir = os.path.join(WORKSPACE_DIR, "VDRL")
                    doc_type = "VDRL"
                elif doc_type == "KOM" or ("KOM" in u or u.endswith(".PPTX")):
                    target_dir = os.path.join(WORKSPACE_DIR, "KOM")
                    doc_type = "KOM"
                elif doc_type in ("VDLR", "CLIENT_LOG") or ("VENDOR DOCUMENT LOG REPORT" in u or "LOG REPORT" in u):
                    target_dir = WORKSPACE_DIR
                    doc_type = "CLIENT_LOG"
                elif doc_type in ("PSR", "PROCUREMENT_PLAN") or ("PROCUREMENT PLAN" in u or "PSR" in u):
                    target_dir = WORKSPACE_DIR
                    doc_type = "PSR"
                elif "PJM-SCD" in u or "MASTER" in u or "SCHEDULE" in u:
                    target_dir = WORKSPACE_DIR
                    doc_type = "MASTER_SCHEDULE"
                else:
                    target_dir = os.path.join(WORKSPACE_DIR, "WPR")

                os.makedirs(target_dir, exist_ok=True)
                dest_path = os.path.join(target_dir, filename)
                file_bytes = base64.b64decode(content_b64)
                with open(dest_path, "wb") as f:
                    f.write(file_bytes)

                # Automatic Extraction & Sync into project-specific Export Data File
                extraction_info = None
                is_z1f = "ZTK" in u or "Z1F" in u
                upload_proj = "Z1F" if is_z1f else "ASK"
                from vendor_dashboard.ingestors import get_export_data_path
                export_path = get_export_data_path(WORKSPACE_DIR, project=upload_proj)
                if doc_type == "WPR" and u.endswith((".PDF", ".DOCX")):
                    from vendor_dashboard.extractors.wpr_mom_extractor import extract_wpr_pdf
                    from vendor_dashboard.extractors.excel_sync import sync_wpr_data_to_excel
                    from vendor_dashboard.ingestors import parse_wpr_file
                    try:
                        if u.endswith(".DOCX"):
                            pw = parse_wpr_file(dest_path)
                            pkg_code = pw.get("package_code") or filename.split("-")[4] if "-" in filename else filename
                            rep_no = pw.get("report_no") or "WPR-001"
                            cut_dt = pw.get("cut_off_date") or pw.get("report_date") or ""
                            sc_obj = pw.get("scurve") or {}
                            sc_pts = sc_obj.get("points", []) if isinstance(sc_obj, dict) else list(sc_obj)
                            plan_p = float(pw.get("overall_planned_pct") or (sc_obj.get("planned_to_date_pct") if isinstance(sc_obj, dict) else 0.0) or 0.0)
                            act_p = float(pw.get("overall_actual_pct") or (sc_obj.get("actual_to_date_pct") if isinstance(sc_obj, dict) else 0.0) or 0.0)
                            scurve_rows = [
                                {
                                    "package": pkg_code,
                                    "report_no": rep_no,
                                    "cutoff_date": cut_dt,
                                    "week_no": int(pt.get("week_num") or (re.findall(r"\d+", str(pt.get("week") or pt.get("week_no") or "0")) or [0])[0]),
                                    "week_date": pt.get("date") or pt.get("week_date") or "",
                                    "plan_cum_pct": float(pt.get("planned_pct") or pt.get("planned_cum") or pt.get("plan_cum_pct") or 0.0),
                                    "actual_cum_pct": float(pt.get("actual_pct") or pt.get("actual_cum") or pt.get("actual_cum_pct") or 0.0) if pt.get("actual_pct") is not None else None,
                                    "plan_inc_pct": float(pt.get("planned_inc") or pt.get("plan_inc_pct") or 0.0),
                                    "actual_inc_pct": float(pt.get("actual_inc") or pt.get("actual_inc_pct") or 0.0),
                                }
                                for pt in sc_pts if isinstance(pt, dict)
                            ]
                            wpr_rec = {
                                "summary": {
                                    "package": pkg_code, "report_no": rep_no, "cutoff_date": cut_dt,
                                    "po_number": pw.get("po_number") or "", "vendor_name": pw.get("wpr_vendor_name") or "",
                                    "target_delivery": "", "overall_plan_pct": plan_p, "overall_actual_pct": act_p,
                                    "inc_plan_pct": 0.0, "inc_actual_pct": 0.0, "deviation_pct": round(act_p - plan_p, 2),
                                    "status_tone": "On Track" if act_p >= plan_p else "Behind",
                                    "executive_highlights": "; ".join(pw.get("this_period_activities", [])[:3]),
                                    "areas_of_concern": "; ".join(pw.get("outstanding_issues", [])[:3]),
                                    "critical_path_notes": "",
                                },
                                "scurve": scurve_rows,
                                "gantt_tasks": pw.get("gantt_tasks", []),
                                "sub_orders": [],
                                "milestones": pw.get("milestones", []),
                                "milestone_progress": pw.get("milestone_progress", []),
                            }
                        else:
                            wpr_rec = extract_wpr_pdf(file_bytes, filename=filename)
                            if not wpr_rec.get("milestones") or not wpr_rec.get("milestone_progress"):
                                pw_pdf = parse_wpr_file(dest_path)
                                wpr_rec.setdefault("milestones", pw_pdf.get("milestones", []))
                                wpr_rec.setdefault("milestone_progress", pw_pdf.get("milestone_progress", []))
                        stats = sync_wpr_data_to_excel(export_path, [wpr_rec])
                        extraction_info = {
                            "package": wpr_rec["summary"]["package"],
                            "report_no": wpr_rec["summary"]["report_no"],
                            "cutoff_date": wpr_rec["summary"]["cutoff_date"],
                            "plan_pct": wpr_rec["summary"]["overall_plan_pct"],
                            "actual_pct": wpr_rec["summary"]["overall_actual_pct"],
                            "deviation_pct": wpr_rec["summary"]["deviation_pct"],
                            "tasks_count": len(wpr_rec["gantt_tasks"]),
                            "scurve_count": len(wpr_rec["scurve"]),
                            "sub_orders_count": len(wpr_rec["sub_orders"]),
                            "stats": stats,
                        }
                    except Exception as e:
                        print(f"[WARN] Failed extracting uploaded WPR {filename}: {e}")

                elif doc_type == "MOM" and u.endswith(".PDF"):
                    from vendor_dashboard.extractors.wpr_mom_extractor import extract_mom_pdf
                    from vendor_dashboard.extractors.excel_sync import sync_mom_data_to_excel
                    try:
                        mom_rec = extract_mom_pdf(file_bytes, filename=filename)
                        stats = sync_mom_data_to_excel(export_path, [mom_rec])
                        extraction_info = {
                            "package": mom_rec["metadata"]["package"],
                            "mom_no": mom_rec["metadata"]["mom_no"],
                            "meeting_date": mom_rec["metadata"]["meeting_date"],
                            "action_items_count": len(mom_rec["action_items"]),
                            "procurement_count": len(mom_rec["procurement_status"]),
                            "stats": stats,
                        }
                    except Exception as e:
                        print(f"[WARN] Failed extracting uploaded MOM {filename}: {e}")

                with _CACHE_LOCK:
                    _SNAPSHOT_CACHE.clear()
                    vendor_dashboard.ingestors._INGEST_CACHE.clear()

                self._send_json(200, {
                    "status": "uploaded",
                    "doc_type": doc_type,
                    "filename": filename,
                    "saved_to": os.path.relpath(dest_path, WORKSPACE_DIR),
                    "extraction": extraction_info,
                })
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        if parsed.path == "/api/push-github":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                raw = self.rfile.read(length) if length > 0 else b"{}"
                payload = json.loads(raw.decode("utf-8")) if raw else {}
                token = str(payload.get("token") or os.environ.get("GITHUB_TOKEN", "")).strip()
                repo = str(payload.get("repo") or os.environ.get("GITHUB_REPO", "")).strip()
                branch = str(payload.get("branch") or os.environ.get("GITHUB_BRANCH", "main")).strip() or "main"
                commit_msg = str(payload.get("message") or "").strip()

                if not token or not repo:
                    self._send_json(400, {
                        "error": "GITHUB_TOKEN and GITHUB_REPO are not configured in Render Environment Variables.",
                        "needs_config": True,
                    })
                    return

                result = _push_files_to_github(
                    workspace_dir=WORKSPACE_DIR,
                    token=token,
                    repo=repo,
                    branch=branch,
                    commit_msg=commit_msg,
                )
                self._send_json(200, result)
            except urllib.error.HTTPError as exc:
                err_body = exc.read().decode("utf-8", errors="ignore")
                self._send_json(502, {"error": f"GitHub API HTTP {exc.code}: {err_body[:300]}"})
            except Exception as exc:
                self._send_json(500, {"error": str(exc)})
            return

        self._send_json(404, {"error": "Unknown POST endpoint"})


def _collect_github_render_files(workspace_dir: str) -> List[str]:
    """Return root-relative filenames of Render-required data files to sync to GitHub."""
    candidates: List[str] = []
    for name in ("Export Data File_ASK.xlsx", "Export Data File_Z1F.xlsx"):
        if os.path.isfile(os.path.join(workspace_dir, name)):
            candidates.append(name)
    for pattern in ("*Procurement Plan*.xlsx", "*Vendor Document Log Report*.xlsx", "*Vendor Document Log Report*.xls"):
        for path in sorted(glob.glob(os.path.join(workspace_dir, pattern))):
            bn = os.path.basename(path)
            if not bn.startswith("~$") and bn not in candidates:
                candidates.append(bn)
    return candidates


def _github_api_request(url: str, token: str, method: str = "GET", body: Dict[str, Any] = None) -> Dict[str, Any]:
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "Vendor-Tracking-Dashboard",
    }
    data = None
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _push_files_to_github(
    workspace_dir: str,
    token: str,
    repo: str,
    branch: str = "main",
    commit_msg: str = "",
) -> Dict[str, Any]:
    """Commit all changed Render data files in a single atomic commit via GitHub Git Data API."""
    files = _collect_github_render_files(workspace_dir)
    if not files:
        return {"status": "up_to_date", "message": "No Render data files found in workspace.", "pushed_files": []}

    api_base = f"https://api.github.com/repos/{repo}"
    ref_data = _github_api_request(f"{api_base}/git/ref/heads/{branch}", token)
    head_sha = ref_data["object"]["sha"]

    commit_data = _github_api_request(f"{api_base}/git/commits/{head_sha}", token)
    base_tree_sha = commit_data["tree"]["sha"]

    tree_data = _github_api_request(f"{api_base}/git/trees/{base_tree_sha}", token)
    existing_shas = {item["path"]: item["sha"] for item in tree_data.get("tree", []) if item.get("type") == "blob"}

    tree_entries: List[Dict[str, str]] = []
    pushed_files: List[str] = []
    for fname in files:
        fpath = os.path.join(workspace_dir, fname)
        with open(fpath, "rb") as f:
            raw = f.read()
        local_git_sha = hashlib.sha1(f"blob {len(raw)}\0".encode("utf-8") + raw).hexdigest()
        if existing_shas.get(fname) == local_git_sha:
            continue

        blob_resp = _github_api_request(
            f"{api_base}/git/blobs",
            token,
            method="POST",
            body={"content": base64.b64encode(raw).decode("ascii"), "encoding": "base64"},
        )
        tree_entries.append({
            "path": fname,
            "mode": "100644",
            "type": "blob",
            "sha": blob_resp["sha"],
        })
        pushed_files.append(fname)

    if not tree_entries:
        return {
            "status": "up_to_date",
            "message": "All Render data files already match GitHub.",
            "files_checked": len(files),
            "pushed_files": [],
        }

    new_tree = _github_api_request(
        f"{api_base}/git/trees",
        token,
        method="POST",
        body={"base_tree": base_tree_sha, "tree": tree_entries},
    )
    ts_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    msg = commit_msg or f"Dashboard weekly data sync ({ts_str}): {', '.join(pushed_files)}"
    new_commit = _github_api_request(
        f"{api_base}/git/commits",
        token,
        method="POST",
        body={"message": msg, "tree": new_tree["sha"], "parents": [head_sha]},
    )
    _github_api_request(
        f"{api_base}/git/refs/heads/{branch}",
        token,
        method="PATCH",
        body={"sha": new_commit["sha"]},
    )
    return {
        "status": "pushed",
        "repo": repo,
        "branch": branch,
        "commit_sha": new_commit["sha"],
        "commit_url": f"https://github.com/{repo}/commit/{new_commit['sha']}",
        "pushed_files": pushed_files,
    }


class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


def run_server(port: int = 8080) -> None:
    server = ThreadingHTTPServer(("0.0.0.0", port), DashboardHandler)
    print(f"Vendor Package Monitoring Dashboard (ASK & Z1F) serving on http://localhost:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    port_env = int(os.environ.get("PORT", "8080"))
    run_server(port_env)

