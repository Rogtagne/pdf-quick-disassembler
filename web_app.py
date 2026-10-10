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
import time
import tempfile
import shutil
from pathlib import Path
from flask import (
    Flask, request, jsonify, send_file,
    render_template, abort
)

# ── Asegurar que los motores del proyecto sean importables ────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from utils.logging import get_logger
from topo_vs_pdf import (
    fase1_convert,
    fase2_convert,
    fase2_word_convert,
    tabla_excel_recover,
    POPPLER_PATH,
)

log = get_logger("web_app")

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


def _reject(route: str, reason: str, filename: str | None = None):
    """Registra y devuelve un error 400 de validación."""
    log.warning("%s rechazado: %s (archivo=%r)", route, reason, filename)
    return jsonify({"error": reason}), 400


def _cleanup_job(job_id: str, delay_seconds: int = 300):
    """Elimina archivos temporales del trabajo después de N segundos."""
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
        log.warning("/api/job: trabajo no encontrado [%s]", job_id)
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
        log.warning("/api/download: trabajo no disponible [%s] (status=%s)",
                    job_id, job["status"] if job else "inexistente")
        abort(404)
    result = Path(job["result"]) if job["result"] else None
    if not result or not result.exists():
        log.warning("/api/download: resultado no encontrado [%s] (%s)", job_id, result)
        abort(404)
    log.info("/api/download [%s]: %s", job_id, result.name)
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
        return _reject("/api/fase1", "Sin archivo")

    ext = Path(f.filename or "").suffix.lower()
    if ext not in (".docx", ".pptx"):
        return _reject("/api/fase1", "Solo se aceptan .docx y .pptx", f.filename)

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / f"input{ext}"
    dst_path = job_dir / (Path(f.filename or "output").stem + ".pdf")
    f.save(str(src_path))
    log.info("/api/fase1 [%s] inicio: archivo=%r", job_id, f.filename)

    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None}

    def worker():
        t0 = time.monotonic()
        try:
            fase1_convert(
                str(src_path), str(dst_path),
                progress_cb=lambda p, m: _job_update(job_id, p, m),
            )
            with _jobs_lock:
                _jobs[job_id]["status"] = "done"
                _jobs[job_id]["result"] = str(dst_path)
            log.info("/api/fase1 [%s] terminado en %.1fs", job_id, time.monotonic() - t0)
        except Exception as e:
            log.exception("/api/fase1 [%s] error tras %.1fs", job_id, time.monotonic() - t0)
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
        return _reject("/api/fase2", "Sin archivo")
    if Path(f.filename or "").suffix.lower() != ".pdf":
        return _reject("/api/fase2", "Solo se aceptan archivos .pdf", f.filename)

    dpi              = int(request.form.get("dpi", 150))
    include_text     = request.form.get("include_text", "1") == "1"
    text_mode        = request.form.get("text_mode", "paragraphs")
    composition_mode = request.form.get("composition_mode", "objects")
    preserve_breaks  = request.form.get("preserve_line_breaks", "1") == "1"

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / "input.pdf"
    dst_path = job_dir / (Path(f.filename or "output").stem + "_editable.pptx")
    f.save(str(src_path))
    log.info("/api/fase2 [%s] inicio: archivo=%r dpi=%d texto=%s modo_texto=%s "
             "composicion=%s saltos=%s", job_id, f.filename, dpi, include_text,
             text_mode, composition_mode, preserve_breaks)

    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None}

    def worker():
        t0 = time.monotonic()
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
            log.info("/api/fase2 [%s] terminado en %.1fs", job_id, time.monotonic() - t0)
        except Exception as e:
            log.exception("/api/fase2 [%s] error tras %.1fs", job_id, time.monotonic() - t0)
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["msg"]    = str(e)
        finally:
            _cleanup_job(job_id)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── FASE 2 (variante) — PDF a WORD ───────────────────────────────────────────

@app.route("/api/fase2_word", methods=["POST"])
def api_fase2_word():
    f = request.files.get("file")
    if not f:
        return _reject("/api/fase2_word", "Sin archivo")
    if Path(f.filename or "").suffix.lower() != ".pdf":
        return _reject("/api/fase2_word", "Solo se aceptan archivos .pdf", f.filename)

    include_images = request.form.get("include_images", "1") == "1"
    image_dpi      = int(request.form.get("image_dpi", 150))

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / "input.pdf"
    dst_path = job_dir / (Path(f.filename or "output").stem + "_editable.docx")
    f.save(str(src_path))
    log.info("/api/fase2_word [%s] inicio: archivo=%r imagenes=%s dpi=%d",
             job_id, f.filename, include_images, image_dpi)

    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None}

    def worker():
        t0 = time.monotonic()
        try:
            fase2_word_convert(
                str(src_path), str(dst_path),
                include_images=include_images,
                image_dpi=image_dpi,
                progress_cb=lambda p, m: _job_update(job_id, p, m),
            )
            with _jobs_lock:
                _jobs[job_id]["status"] = "done"
                _jobs[job_id]["result"] = str(dst_path)
            log.info("/api/fase2_word [%s] terminado en %.1fs", job_id, time.monotonic() - t0)
        except Exception as e:
            log.exception("/api/fase2_word [%s] error tras %.1fs", job_id, time.monotonic() - t0)
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["msg"]    = str(e)
        finally:
            _cleanup_job(job_id)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── FASE 3 — EXCEL TABLE RECOVERY ────────────────────────────────────────────

@app.route("/api/table_recovery", methods=["POST"])
def api_table_recovery():
    f = request.files.get("file")
    if not f:
        return _reject("/api/table_recovery", "Sin archivo")
    if Path(f.filename or "").suffix.lower() != ".pdf":
        return _reject("/api/table_recovery", "Solo se aceptan archivos .pdf", f.filename)

    gap_tol       = float(request.form.get("gap_tol", 18))
    max_font_size = float(request.form.get("max_font_size", 16))

    job_id  = uuid.uuid4().hex
    job_dir = _job_path(job_id)
    job_dir.mkdir(parents=True)

    src_path = job_dir / "input.pdf"
    out_dir  = job_dir / "output"
    f.save(str(src_path))
    log.info("/api/table_recovery [%s] inicio: archivo=%r gap_tol=%s max_font=%s",
             job_id, f.filename, gap_tol, max_font_size)

    # El resultado descargable será el XLSX; si no hay tablas, un txt de aviso.
    with _jobs_lock:
        _jobs[job_id] = {"status": "running", "pct": 0,
                          "msg": "Iniciando…", "result": None,
                          "recovered": 0, "grid_pages": 0,
                          "band_pages": 0, "warn_cells": 0}

    def worker():
        t0 = time.monotonic()
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
                # Se devuelve el detalle de calidad para que la pantalla pueda
                # avisar si hay celdas que conviene verificar contra el PDF.
                _jobs[job_id]["grid_pages"] = result.get("grid_pages", 0)
                _jobs[job_id]["band_pages"] = result.get("band_pages", 0)
                _jobs[job_id]["warn_cells"] = result.get("warn_cells", 0)

                if result["msg"] == "sin_tablas":
                    _jobs[job_id]["status"] = "done"
                    _jobs[job_id]["msg"]    = "⚠️ Sin tablas compatibles"
                    _jobs[job_id]["result"] = None
                else:
                    _jobs[job_id]["status"] = "done"
                    _jobs[job_id]["result"] = str(result["xlsx"])
            log.info("/api/table_recovery [%s] terminado en %.1fs: tablas=%s "
                     "grid=%s bandas=%s celdas_a_verificar=%s%s",
                     job_id, time.monotonic() - t0, result["recovered"],
                     result.get("grid_pages", 0), result.get("band_pages", 0),
                     result.get("warn_cells", 0),
                     " (sin tablas)" if result["msg"] == "sin_tablas" else "")
        except Exception as e:
            log.exception("/api/table_recovery [%s] error tras %.1fs",
                          job_id, time.monotonic() - t0)
            with _jobs_lock:
                _jobs[job_id]["status"] = "error"
                _jobs[job_id]["msg"]    = str(e)
        finally:
            _cleanup_job(job_id)

    threading.Thread(target=worker, daemon=True).start()
    return jsonify({"job_id": job_id})


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    log.info("PDF Quick Disassembler — Servidor Web")
    if POPPLER_PATH:
        log.info("Poppler: %s", POPPLER_PATH)
    else:
        log.warning("Poppler: NO ENCONTRADO")
    log.info("Abre: http://localhost:5000")
    app.run(debug=True, host="0.0.0.0", port=5000)
