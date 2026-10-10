"""
Utilidad de logging — PDF Quick Disassembler
GoTopo 2026

Uso:
    from utils.logging import get_logger
    log = get_logger(__name__)
    log.info("Procesando %s", nombre)

Variables de entorno opcionales:
    PQD_LOG_LEVEL   Nivel de log: DEBUG, INFO, WARNING, ERROR  (default: INFO)
    PQD_LOG_FILE    Ruta de archivo de log adicional           (default: solo consola)
"""

from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

# ── Configuración ─────────────────────────────────────────────────────────────
ROOT_NAME   = "pqd"
LOG_FORMAT  = "%(asctime)s [%(levelname)s] [pid %(process)d] %(name)s: %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

_configured = False


def setup_logging(level: str | int | None = None, log_file: str | None = None) -> logging.Logger:
    """Configura el logger raíz del proyecto una sola vez (idempotente)."""
    global _configured
    root = logging.getLogger(ROOT_NAME)
    if _configured:
        return root

    level    = level or os.environ.get("PQD_LOG_LEVEL", "INFO")
    log_file = log_file or os.environ.get("PQD_LOG_FILE")

    root.setLevel(level.upper() if isinstance(level, str) else level)
    root.propagate = False
    formatter = logging.Formatter(LOG_FORMAT, DATE_FORMAT)

    # Consola (stdout → visible en los logs de Render / gunicorn)
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    root.addHandler(console)

    # Archivo opcional con rotación (5 MB x 3 respaldos)
    if log_file:
        file_handler = RotatingFileHandler(
            log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)

    _configured = True
    return root


def get_logger(name: str | None = None) -> logging.Logger:
    """Devuelve un logger hijo de 'pqd', configurando el logging si hace falta."""
    setup_logging()
    if not name or name == ROOT_NAME:
        return logging.getLogger(ROOT_NAME)
    if name.startswith(ROOT_NAME + "."):
        return logging.getLogger(name)
    return logging.getLogger(f"{ROOT_NAME}.{name}")
