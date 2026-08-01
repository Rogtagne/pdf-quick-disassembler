"""
PDF Quick Disassembler — Servidor Web
GoTopo 2026

Uso:
    python web_app.py                  # desarrollo local  http://localhost:5000
    gunicorn web_app:app               # producción (Render.com)

Variables de entorno opcionales:
    PQD_MAX_MB      Tamaño máximo de archivo en MB  (default: 50)
    PQD_UPLOAD_DIR  Carpeta temporal de uploads     (default: /tmp/pqd_uploads)
"""

from __future__ import annotations

import os
import sys
import uuid
import threading
import tempfile
import shutil
from pathlib import Path
from flask import (
    Flask, request, jsonify, send_file,
    render_template, abort
)

# ── Asegurar que los motores del proyecto sean importables ────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from topo_vs_pdf import (
    fase1_convert,
    fase2_convert,
    tabla_excel_recover,
    POPPLER_PATH,
)

# ── Configuración ─────────────────────────────────────────────────────────────
MAX_MB      = int(os.environ.get("PQD_MAX_MB", 50))
UPLOAD_ROOT = Path(os.environ.get("PQD_UPLOAD_DIR", tempfile.gettempdir())) / "pqd_uploads"
UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

# ── Registro de trabajos en memoria ──────────────────────────────────────────
# { job_id: { "status": "running"|"done"|"error", "pct": int,
#             "msg": str, "result": Path|None } }
_jobs: dict = {}
_jobs_lock  = threading.Lock()


def _job_update(job_id: str, pct: int, msg: str):
    with _jobs_lock:
        if job_id in _jobs:
            _jobs[job_id]["pct"] = pct
            _jobs[job_id]["msg"] = msg


def _job_path(job_id: str) -> Path:
    return UPLOAD_ROOT / job_id


def _cleanup_job(job_id: str, delay_seconds: int = 300):
    """Elimina archivos temporales del trabajo después de N segundos."""
    import time
    def _do():
        time.sleep(delay_seconds)
        try:
            shutil.rmtree(_job_path(job_id), ignore_errors=True)
            with _jobs_lock:
                _jobs.pop(job_id, None)
        except Exception:
            pass
    threading.Thread(target=_do, daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
#  RUTAS — UI
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html",
                           poppler_ok=bool(POPPLER_PATH),
                           max_mb=MAX_MB)


# ══════════════════════════════════════════════════════════════════════════════
#  RUTAS — API
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/api/status")
def api_status():
    return jsonify({"ok": True, "poppler": bool(POPPLER_PATH), "max_mb": MAX_MB})


@app.route("/api/job/<job_id>")
def api_job(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        abort(404)
    # Serializar a JSON-safe: convertir Path a str
    safe = {k: (str(v) if isinstance(v, Path) else v)
            for k, v in job.items()}
    return jsonify(safe)


@app.route("/api/download/<job_id>")
def api_download(job_id: str):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job or job["status"] != "done":
        abort(404)
    result = Path(job["result"]) if job["result"] else None
    if not result or not result.exists():
        abort(404)
    return send_file(
        str(result),
        as_attachment=True,
        download_name=result.name,
    )


# ── FASE 1 ────────────────────────────────────────────────────────────────────

@app.route("/api/fase1", methods=["POST"])
def api_fase1():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "Sin archivo"}), 400

    ext = Path(f.filename or "").suffix.lower()
    if ext not in (".docx", ".pptx"):
        return jsonify({"error": "Solo se aceptan .docx y .pptx"}), 400

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / f"input{ext}"
    dst_path = job_dir / (Path(f.filename or "output").stem + ".pdf")
    f.save(str(src_path))

    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None}

    def worker():
        try:
            fase1_convert(
                str(src_path), str(dst_path),
                progress_cb=lambda p, m: _job_update(job_id, p, m),
            )
            with _jobs_lock:
                _jobs[job_id]["status"] = "done"
                _jobs[job_id]["result"] = str(dst_path)
        except Exception as e:
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["msg"]    = str(e)
        finally:
            _cleanup_job(job_id)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── FASE 2 ────────────────────────────────────────────────────────────────────

@app.route("/api/fase2", methods=["POST"])
def api_fase2():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "Sin archivo"}), 400
    if Path(f.filename or "").suffix.lower() != ".pdf":
        return jsonify({"error": "Solo se aceptan archivos .pdf"}), 400

    dpi              = int(request.form.get("dpi", 150))
    include_text     = request.form.get("include_text", "1") == "1"
    text_mode        = request.form.get("text_mode", "paragraphs")
    composition_mode = request.form.get("composition_mode", "background")
    preserve_breaks  = request.form.get("preserve_line_breaks", "1") == "1"

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / "input.pdf"
    dst_path = job_dir / (Path(f.filename or "output").stem + "_editable.pptx")
    f.save(str(src_path))

    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None}

    def worker():
        try:
            fase2_convert(
                str(src_path), str(dst_path),
                dpi=dpi,
                include_text=include_text,
                text_mode=text_mode,
                composition_mode=composition_mode,
                preserve_line_breaks=preserve_breaks,
                progress_cb=lambda p, m: _job_update(job_id, p, m),
            )
            with _jobs_lock:
                _jobs[job_id]["status"] = "done"
                _jobs[job_id]["result"] = str(dst_path)
        except Exception as e:
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["msg"]    = str(e)
        finally:
            _cleanup_job(job_id)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── PDF TABLE RECOVERY ────────────────────────────────────────────────────────

@app.route("/api/table_recovery", methods=["POST"])
def api_table_recovery():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "Sin archivo"}), 400
    if Path(f.filename or "").suffix.lower() != ".pdf":
        return jsonify({"error": "Solo se aceptan archivos .pdf"}), 400

    gap_tol       = float(request.form.get("gap_tol", 18))
    max_font_size = float(request.form.get("max_font_size", 16))

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / "input.pdf"
    out_dir  = job_dir / "output"
    f.save(str(src_path))

    # El resultado descargable será el XLSX; si no hay tablas, un txt de aviso.
    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None,
                          "recovered": 0}

    def worker():
        try:
            result = tabla_excel_recover(
                src_path=str(src_path),
                out_dir=out_dir,
                gap_tol=gap_tol,
                max_font_size=max_font_size,
                progress_cb=lambda p, m: _job_update(job_id, p, m),
            )
            with _jobs_lock:
                _jobs[job_id]["recovered"] = result["recovered"]
                if result["msg"] == "sin_tablas":
                    _jobs[job_id]["status"] = "done"
                    _jobs[job_id]["msg"]    = "⚠️ Sin tablas compatibles"
                    _jobs[job_id]["result"] = None
                else:
                    _jobs[job_id]["status"] = "done"
                    _jobs[job_id]["result"] = str(result["xlsx"])
        except Exception as e:
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["msg"]    = str(e)
        finally:
            _cleanup_job(job_id)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("PDF Quick Disassembler — Servidor Web")
    print(f"Poppler: {POPPLER_PATH or 'NO ENCONTRADO'}")
    print("Abre: http://localhost:5000")
    app.run(debug=True, host="0.0.0.0", port=5000)
