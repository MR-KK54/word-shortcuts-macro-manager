"""Word & PDF Page Exporter Pro - backend server.

Implements the API contract used by static/app.js:
  POST /api/upload, /api/inspect, /api/naming-preview, /api/export
  GET  /api/preview/<name>, /api/output-preview/<job_id>/<name>
  GET  /api/job/<id>, /api/download/<job_id>/<name>
  POST /api/job/<id>/cancel, /api/clear-storage
  static files served from the project root.
"""

import os
import re
import shutil
import threading
import time
import uuid

from flask import Flask, jsonify, request, send_file, send_from_directory, Response
from lxml import etree

from engine import convert, naming, paginate, pdf_split, ranges, docx_trim, verify
from engine.renderer import Renderer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNTIME_DIR = os.environ.get("RUNTIME_DIR") or os.path.join(BASE_DIR, "runtime")
UPLOADS_DIR = os.path.join(RUNTIME_DIR, "uploads")
JOBS_DIR = os.path.join(RUNTIME_DIR, "jobs")
RENDERS_DIR = os.path.join(RUNTIME_DIR, "renders")

for d in (UPLOADS_DIR, JOBS_DIR, RENDERS_DIR):
    os.makedirs(d, exist_ok=True)

app = Flask(__name__, static_folder=None, static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024
renderer = Renderer(RENDERS_DIR)

JOBS = {}
JOBS_LOCK = threading.Lock()


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

_WORD_EXT = {".docx", ".doc", ".docm", ".dotx", ".dotm", ".rtf"}
_PDF_EXT = {".pdf"}


def _clean_name(name):
    name = os.path.basename(name or "").replace("\\", "/").split("/")[-1]
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()
    return name or "file"


def resolve_input(name):
    """Resolve an export 'file' entry (uploaded name or server path)."""
    clean = _clean_name(name)
    up = os.path.join(UPLOADS_DIR, clean)
    if os.path.exists(up):
        return up
    if os.path.exists(name) and os.path.isfile(name):
        return name
    if os.path.exists(clean):
        return clean
    raise FileNotFoundError(f"File not found on server: {name}")


def classify(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in _PDF_EXT:
        return "pdf"
    if ext in _WORD_EXT or ext == ".pdf":
        return "word"
    return "word"


def _word_namespace():
    return "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def count_sections(path):
    import zipfile

    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            target = "word/document.xml"
            if target not in names:
                return 1
            root = etree.fromstring(z.read(target))
            ns = {"w": _word_namespace()}
            return len(root.findall(".//w:sectPr", ns)) or 1
    except Exception:
        return 1


def get_document_info(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in _PDF_EXT:
        return pdf_split.pdf_page_count(path), 1
    # Word family: paginate with MS Word COM (Windows) or Word's recorded
    # pagination markers / explicit breaks (Linux/Render).
    docx_path = path
    tmp = None
    if ext != ".docx":
        tmp = os.path.join(RUNTIME_DIR, "tmp_inspect")
        os.makedirs(tmp, exist_ok=True)
        try:
            docx_path = convert.normalize_to_docx(path, tmp)
        except Exception:
            docx_path = path
    try:
        page_count, _ = paginate.paginate_docx(docx_path)
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)
    return page_count, count_sections(path)


# --------------------------------------------------------------------------
# static + PWA
# --------------------------------------------------------------------------

@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/manifest.json")
def manifest():
    return send_from_directory(BASE_DIR, "manifest.json")


@app.route("/sw.js")
def service_worker():
    return send_from_directory(BASE_DIR, "sw.js")


@app.route("/static/<path:filename>")
def static_files(filename):
    return send_from_directory(os.path.join(BASE_DIR, "static"), filename)


# --------------------------------------------------------------------------
# upload / inspect / naming
# --------------------------------------------------------------------------

@app.route("/api/upload", methods=["POST"])
def api_upload():
    files = request.files.getlist("files")
    saved = []
    for f in files:
        if not f.filename:
            continue
        name = _clean_name(f.filename)
        dest = os.path.join(UPLOADS_DIR, name)
        f.save(dest)
        saved.append({"name": name, "size": os.path.getsize(dest)})
    return jsonify({"files": saved})


@app.route("/api/inspect", methods=["POST"])
def api_inspect():
    data = request.get_json(silent=True) or {}
    name = data.get("name")
    try:
        path = resolve_input(name)
        page_count, section_count = get_document_info(path)
        return jsonify({"name": name, "page_count": page_count, "section_count": section_count})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/naming-preview", methods=["POST"])
def api_naming_preview():
    data = request.get_json(silent=True) or {}
    pattern = data.get("pattern", "{original_name}_pages_{start_page}-{end_page}")
    fmt = data.get("format", "docx")
    sample = data.get("sample", "SampleReport.docx")
    try:
        ext = "pdf" if os.path.splitext(sample)[1].lower() == ".pdf" else fmt
        preview = naming.apply_template(pattern, sample, 1, 2, ext)
        return jsonify({"preview": preview})
    except Exception as e:
        return jsonify({"preview": "[Pattern Error]", "error": str(e)})


# --------------------------------------------------------------------------
# previews
# --------------------------------------------------------------------------

def _render_preview_response(path, page_param, width_param):
    try:
        page = int(page_param or 1)
        width = int(width_param or 1200)
    except ValueError:
        page, width = 1, 1200
    try:
        png, total = renderer.render_page(path, page, width)
    except Exception as e:
        # Transient failures (Word busy, first-open dialog) are retried once.
        import time

        time.sleep(1.0)
        try:
            png, total = renderer.render_page(path, page, width)
        except Exception as e2:
            return jsonify({"error": f"Preview render failed: {e2}"}), 500
    return Response(png, mimetype="image/png", headers={"X-Total-Pages": str(total), "Cache-Control": "no-store"})


@app.route("/api/preview/<path:name>")
def api_preview(name):
    try:
        path = resolve_input(name)
    except Exception as e:
        return jsonify({"error": str(e)}), 404
    return _render_preview_response(path, request.args.get("page"), request.args.get("w"))


@app.route("/api/output-preview/<job_id>/<path:name>")
def api_output_preview(job_id, name):
    safe = _clean_name(name)
    path = os.path.join(JOBS_DIR, job_id, "outputs", safe)
    if not os.path.exists(path):
        return jsonify({"error": "Output not found"}), 404
    return _render_preview_response(path, request.args.get("page"), request.args.get("w"))


@app.route("/api/download/<job_id>/<path:name>")
def api_download(job_id, name):
    safe = _clean_name(name)
    path = os.path.join(JOBS_DIR, job_id, "outputs", safe)
    if not os.path.exists(path):
        return jsonify({"error": "Output not found"}), 404
    return send_file(path, as_attachment=True, download_name=safe)


@app.route("/api/diagnostics")
def api_diagnostics():
    """Report which rendering engines are detected on this server."""
    import getpass
    import shutil as _shutil
    import sys as _sys

    from engine import convert as _convert
    from engine import word_com as _word

    soffice = _convert.find_soffice()
    try:
        import os as _os

        euid = _os.geteuid()
    except Exception:
        euid = None
    return jsonify(
        {
            "platform": os.name,
            "runtime_user": getpass.getuser(),
            "euid": euid,
            "has_apt_get": bool(_shutil.which("apt-get") or os.path.exists("/usr/bin/apt-get")),
            "has_sudo": bool(_shutil.which("sudo")),
            "word_com_available": bool(_word.word_available()),
            "soffice_path": soffice,
            "soffice_available": bool(_convert.soffice_available()),
            "soffice_install_attempted": _convert._soffice_install_tried,
            "soffice_install_log": _convert._soffice_install_log,
            "soffice_install_error": _convert._soffice_install_error,
            "python_renderer_available": True,
            "python": _sys.version,
            "hint": (
                "OK: MS Word COM available"
                if _word.word_available()
                else ("OK: LibreOffice available" if _convert.soffice_available() else "OK: Pure-Python PyMuPDF fallback renderer active")
            ),
        }
    )


# --------------------------------------------------------------------------
# export jobs
# --------------------------------------------------------------------------

def _new_job(body):
    job_id = uuid.uuid4().hex
    job = {
        "id": job_id,
        "status": "queued",
        "current_status": "Queued",
        "completed": 0,
        "total": 0,
        "success_count": 0,
        "fail_count": 0,
        "logs": [{"level": "info", "message": "Job queued."}],
        "errors": [],
        "outputs": [],
        "request": body,
        "cancel": False,
    }
    JOBS[job_id] = job
    return job


def _log(job, level, message):
    job["logs"].append({"level": level, "message": message})
    if len(job["logs"]) > 2000:
        job["logs"] = job["logs"][-1000:]


def _process_one(job, src_path):
    """Build outputs for every range of a single source file.

    Returns (outputs, expected_count).
    """
    spec = job["request"].get("range", "1-end")
    outputs = []
    src_ext = os.path.splitext(src_path)[1].lower()

    # PDF source -------------------------------------------------------------
    if src_ext in _PDF_EXT:
        page_count = pdf_split.pdf_page_count(src_path)
        ranges_list = ranges.parse_range_spec(spec, page_count)
        for start, end in ranges_list:
            _check_cancel(job)
            try:
                name = naming.apply_template(
                    job["request"].get("naming_pattern", "{original_name}_pages_{start_page}-{end_page}"),
                    os.path.basename(src_path), start, end, "pdf")
                out = os.path.join(JOBS_DIR, job["id"], "outputs", name)
                pdf_split.split_pdf_range(src_path, start, end, out)
                outputs.append(name)
            except Exception as e:
                job["errors"].append(f"{os.path.basename(src_path)} pages {start}-{end}: {e}")
                job["fail_count"] += 1
            finally:
                job["completed"] += 1
        return outputs, len(ranges_list)

    # Word family ------------------------------------------------------------
    work_dir = os.path.join(JOBS_DIR, job["id"], "work")
    os.makedirs(work_dir, exist_ok=True)
    docx_path = src_path
    if src_ext != ".docx":
        try:
            docx_path = convert.normalize_to_docx(src_path, work_dir)
        except Exception as e:
            job["errors"].append(f"{os.path.basename(src_path)}: {e}")
            return outputs, 1

    # Determine effective target extension.
    fmt = job["request"].get("format", "same")
    if fmt == "same":
        fmt = "docx"
        if src_ext in (".doc", ".rtf", ".docm", ".dot", ".dotx", ".dotm"):
            fmt = src_ext.lstrip(".") if src_ext != ".dot" else "doc"
    target_ext = fmt.lower().lstrip(".")

    try:
        page_count, boundaries = paginate.paginate_docx(docx_path)
    except Exception as e:
        job["errors"].append(f"{os.path.basename(src_path)}: {e}")
        return outputs, 1

    ranges_list = ranges.parse_range_spec(spec, page_count)
    for start, end in ranges_list:
        _check_cancel(job)
        try:
            name = naming.apply_template(
                job["request"].get("naming_pattern", "{original_name}_pages_{start_page}-{end_page}"),
                os.path.basename(src_path), start, end, target_ext)
            out = os.path.join(JOBS_DIR, job["id"], "outputs", name)
            os.makedirs(os.path.dirname(out), exist_ok=True)
            split_docx = os.path.join(work_dir, "_split_tmp.docx")
            docx_trim.split_docx_range(docx_path, start, end, boundaries, split_docx)
            if _verify_enabled():
                _verify_output(job, docx_path, split_docx, start, end, work_dir)
            convert.convert_docx_to(split_docx, target_ext, out)
            outputs.append(name)
        except Exception as e:
            job["errors"].append(f"{os.path.basename(src_path)} pages {start}-{end}: {e}")
            job["fail_count"] += 1
        finally:
            job["completed"] += 1
    return outputs, len(ranges_list)


def _check_cancel(job):
    if job["cancel"]:
        raise _Cancelled()


def _verify_enabled():
    return os.environ.get("VERIFY_EXPORT", "1") not in ("0", "false", "off")


def _verify_output(job, src_path, out_path, start, end, work_dir):
    """Render-compare the exported file against the source pages; auto-correct."""
    _log(job, "info", f"Verifying {os.path.basename(out_path)} against source pages {start}-{end}...")
    vdir = os.path.join(work_dir, "verify")
    try:
        report = verify.verify_export(src_path, out_path, start, end, vdir)
    except Exception as e:
        _log(job, "warn", f"Verification skipped for {os.path.basename(out_path)}: {e}")
        return
    pc = report["page_count"]
    if pc["actual"] != pc["expected"]:
        _log(job, "error",
             f"{os.path.basename(out_path)} FAIL: expected {pc['expected']} page(s), "
             f"got {pc['actual']}.")
    for se in report.get("structural_errors", []):
        _log(job, "error", f"{os.path.basename(out_path)} structural defect: {se}")
    if report["corrected"]:
        _log(job, "info",
             f"{os.path.basename(out_path)} auto-corrected ({report['corrected']} pass(es)).")
    for entry in report["pages"]:
        if not entry["match"]:
            _log(job, "warn",
                 f"{os.path.basename(out_path)} page {entry['page']} differs from "
                 f"source page {entry['source_page']}: {entry['detail']}")
        elif not entry["text_match"]:
            _log(job, "warn",
                 f"{os.path.basename(out_path)} page {entry['page']} text mismatch "
                 f"vs source page {entry['source_page']}.")
    if report["pass"]:
        _log(job, "success",
             f"{os.path.basename(out_path)} verified: {pc['actual']} page(s) match source.")


class _Cancelled(Exception):
    pass


def _worker(job):
    try:
        job["status"] = "running"
        job["current_status"] = "Starting..."
        _log(job, "info", "Job started (engine: trimming/high-fidelity).")
        total_outputs = 0
        files = job["request"].get("files", [])
        for fname in files:
            _check_cancel(job)
            try:
                src = resolve_input(fname)
                job["current_status"] = f"Processing {os.path.basename(src)}..."
                outputs, expected = _process_one(job, src)
                total_outputs += expected
                job["total"] = total_outputs
                job["outputs"].extend(outputs)
                _log(job, "success", f"Finished {os.path.basename(src)} ({len(outputs)} output(s)).")
            except _Cancelled:
                raise
            except Exception as e:
                job["errors"].append(f"{fname}: {e}")
                job["fail_count"] += 1
                _log(job, "error", f"Failed {fname}: {e}")

        job["total"] = max(job["total"], job["completed"])
        job["success_count"] = len(job["outputs"])
        if job["cancel"]:
            job["status"] = "cancelled"
            _log(job, "warn", "Job cancelled by user.")
        else:
            job["status"] = "done"
            _log(job, "info", f"Job complete: {len(job['outputs'])} file(s) produced.")
        job["current_status"] = "Complete" if job["status"] == "done" else "Cancelled"
    except _Cancelled:
        job["status"] = "cancelled"
        job["current_status"] = "Cancelled"
        _log(job, "warn", "Job cancelled.")
    except Exception as e:
        job["status"] = "error"
        job["current_status"] = "Error"
        job["errors"].append(str(e))
        _log(job, "error", "Job error: " + str(e))


@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.get_json(silent=True) or {}
    files = data.get("files")
    if not files:
        return jsonify({"error": "No files selected."}), 400
    if not data.get("range", "").strip():
        return jsonify({"error": "Enter a valid page range."}), 400
    job = _new_job(data)
    os.makedirs(os.path.join(JOBS_DIR, job["id"], "outputs"), exist_ok=True)
    os.makedirs(os.path.join(JOBS_DIR, job["id"], "work"), exist_ok=True)
    t = threading.Thread(target=_worker, args=(job,), daemon=True)
    t.start()
    return jsonify({"job_id": job["id"]})


@app.route("/api/job/<job_id>")
def api_job(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "job_id": job_id,
        "status": job["status"],
        "current_status": job["current_status"],
        "completed": job["completed"],
        "total": job["total"],
        "success_count": job["success_count"],
        "fail_count": job["fail_count"],
        "logs": job["logs"],
        "errors": job["errors"],
        "outputs": job["outputs"],
    })


@app.route("/api/job/<job_id>/cancel", methods=["POST"])
def api_job_cancel(job_id):
    job = JOBS.get(job_id)
    if job:
        job["cancel"] = True
    return jsonify({"ok": True})


@app.route("/api/clear-storage", methods=["POST"])
def api_clear_storage():
    freed = 0
    count = 0
    for d in (UPLOADS_DIR, JOBS_DIR, RENDERS_DIR):
        for root, dirs, files in os.walk(d):
            for f in files:
                p = os.path.join(root, f)
                try:
                    freed += os.path.getsize(p)
                    os.remove(p)
                    count += 1
                except Exception:
                    pass
    JOBS.clear()
    return jsonify({"file_count": count, "reclaimed_mb": round(freed / (1024 * 1024), 2)})


@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "Uploaded files exceed the 512 MB limit."}), 413


@app.errorhandler(404)
def not_found(e):
    if request.path.startswith("/api/"):
        return jsonify({"error": "Not found"}), 404
    return send_from_directory(BASE_DIR, "index.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, threaded=True)
