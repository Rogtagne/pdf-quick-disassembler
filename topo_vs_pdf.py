"""
PDF Quick Disassembler  v3.4  —  Tabla Excel
Produced by GoTopo  —  2026

FASE 1: DOCX / PPTX  →  PDF
FASE 2: PDF  →  PPTX editable  (fondo fiel o modo objetos independientes)
TABLA EXCEL: PDF con tablas de Excel  →  XLSX + CSV
"""

import os
import sys
import threading
import subprocess
import tempfile
import shutil
from pathlib import Path

# Tkinter solo se carga en modo desktop (no en servidor web Linux)
_WEB_MODE = (os.environ.get("PQD_WEB_MODE") == "1"
             or any("gunicorn" in a for a in sys.argv))

if not _WEB_MODE:
    try:
        from tkinter import (
            Tk, ttk, StringVar, Label, Button, Frame,
            filedialog, messagebox
        )
        from tkinter.ttk import Progressbar, Notebook
        _TKINTER_OK = True
    except ImportError:
        _TKINTER_OK = False
else:
    _TKINTER_OK = False

# ── Poppler ───────────────────────────────────────────────────────────────────
def _setup_poppler():
    # Linux / Render.com: pdftoppm instalado via apt (poppler-utils)
    for lp in ["/usr/bin", "/usr/local/bin", "/opt/homebrew/bin"]:
        if Path(lp, "pdftoppm").exists():
            return lp  # ya en PATH del sistema, solo retornar la ruta

    # Windows: buscar en C:\Poppler y rutas comunes
    candidates = []
    poppler_root = Path(r"C:\Poppler")
    if poppler_root.exists():
        for sub in sorted(poppler_root.iterdir(), reverse=True):
            for rel in ["Library/bin", "bin"]:
                p = sub / rel
                if p.exists() and any(p.glob("pdftoppm*")):
                    candidates.insert(0, str(p))
    candidates += [
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files\poppler\bin",
    ]
    for c in candidates:
        if Path(c).exists():
            if c not in os.environ.get("PATH", ""):
                os.environ["PATH"] = c + os.pathsep + os.environ.get("PATH", "")
            return c
    return None

POPPLER_PATH = _setup_poppler()

# ── Paleta de colores ─────────────────────────────────────────────────────────
C_BG      = "#12172B"   # fondo profundo
C_PANEL   = "#1C2240"   # paneles
C_CARD    = "#252D50"   # tarjetas
C_CARD2   = "#1E2840"   # tarjeta secundaria
C_ACCENT  = "#4F7EFF"   # azul principal
C_GREEN   = "#00C896"   # verde éxito
C_WARN    = "#FF6B6B"   # error
C_GOLD    = "#F5C518"   # dorado GoTopo
C_TEXT    = "#E8ECFF"   # texto principal
C_MUTED   = "#7880AA"   # texto secundario
C_BORDER  = "#2E3860"   # bordes


# ══════════════════════════════════════════════════════════════════════════════
#  CONVERSIÓN — FASE 1
# ══════════════════════════════════════════════════════════════════════════════

def _libre_convert(src: Path, dst_dir: Path) -> Path:
    for soffice in ["soffice",
                    r"C:\Program Files\LibreOffice\program\soffice.exe",
                    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        try:
            r = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", str(dst_dir), str(src)],
                capture_output=True, timeout=120)
            out = dst_dir / (src.stem + ".pdf")
            if r.returncode == 0 and out.exists():
                return out
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    raise RuntimeError("LibreOffice no encontrado")


def _docx_to_pdf_python(src: Path, dst: Path):
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm
    doc  = Document(str(src))
    pdfd = SimpleDocTemplate(str(dst), pagesize=letter,
                             leftMargin=2*cm, rightMargin=2*cm,
                             topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story  = []
    for para in doc.paragraphs:
        text  = para.text.strip()
        pname = (para.style.name or "") if para.style else ""
        if not text:
            story.append(Spacer(1, 6)); continue
        st = (styles["h1"] if pname.startswith("Heading 1") else
              styles["h2"] if pname.startswith("Heading 2") else
              styles["Normal"])
        story.append(Paragraph(text, st))
        story.append(Spacer(1, 4))
    if not story:
        story.append(Paragraph("(documento vacío)", styles["Normal"]))
    pdfd.build(story)


def _pptx_to_pdf_python(src: Path, dst: Path):
    from pptx import Presentation as _Prs
    from reportlab.pdfgen import canvas as rl_canvas
    prs  = _Prs(str(src))
    w_pt = prs.slide_width  / 12700   # type: ignore
    h_pt = prs.slide_height / 12700   # type: ignore
    c    = rl_canvas.Canvas(str(dst), pagesize=(w_pt, h_pt))
    for i, slide in enumerate(prs.slides):
        c.setPageSize((w_pt, h_pt))
        c.setFont("Helvetica-Bold", 14)
        c.setFillColorRGB(0.2, 0.2, 0.6)
        c.drawCentredString(w_pt/2, h_pt/2 + 16, f"Diapositiva {i+1}")
        c.setFont("Helvetica", 8)
        c.setFillColorRGB(0.5, 0.5, 0.5)
        c.drawCentredString(w_pt/2, h_pt/2 - 8,
                            "(instale LibreOffice para fidelidad total)")
        y = h_pt - 36
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            for para in shape.text_frame.paragraphs:  # type: ignore
                line = para.text.strip()
                if line and y > 16:
                    c.setFont("Helvetica", 8)
                    c.setFillColorRGB(0.15, 0.15, 0.15)
                    c.drawCentredString(w_pt/2, y, line[:90])
                    y -= 11
        c.showPage()
    c.save()


def fase1_convert(src_path: str, dst_path: str, progress_cb=None) -> bool:
    src = Path(src_path)
    dst = Path(dst_path)
    ext = src.suffix.lower()
    if progress_cb: progress_cb(10, f"Analizando {src.name}…")
    tmp = Path(tempfile.mkdtemp())
    try:
        try:
            if progress_cb: progress_cb(30, "Convirtiendo con LibreOffice…")
            shutil.copy2(_libre_convert(src, tmp), dst)
            if progress_cb: progress_cb(100, "✅  Conversión completada con LibreOffice")
            return True
        except Exception:
            pass
        if progress_cb: progress_cb(50, "Motor Python (solo texto)…")
        if ext == ".docx":   _docx_to_pdf_python(src, dst)
        elif ext == ".pptx": _pptx_to_pdf_python(src, dst)
        else: raise ValueError(f"Formato no soportado: {ext}")
        if progress_cb: progress_cb(100, "✅  Completado (instale LibreOffice para imágenes)")
        return True
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ══════════════════════════════════════════════════════════════════════════════
#  CONVERSIÓN — FASE 2
# ══════════════════════════════════════════════════════════════════════════════

def _is_bullet_or_numbered(text: str) -> bool:
    """
    Detecta inicios típicos de listas para evitar que se peguen al párrafo
    anterior cuando se agrupa el texto extraído del PDF.
    """
    import re

    return bool(re.match(r"^(\u2022|\-|\*|\d+[\.)]|[a-zA-Z][\.)])\s+", text.strip()))


def _join_paragraph_line(previous: str, new_line: str) -> str:
    """
    Une líneas consecutivas en un párrafo continuo.

    Caso especial:
        "administra-" + "ción"  ->  "administración"
    """
    previous = previous.rstrip()
    new_line = new_line.strip()

    if not previous:
        return new_line

    if previous.endswith("-") and new_line:
        return previous[:-1] + new_line

    return previous + " " + new_line


def _extract_text_lines(pdf_path: str):
    """
    Extrae texto embebido del PDF y lo devuelve agrupado por líneas.

    Esta función conserva el comportamiento original del MVP, pero queda
    separada para poder ofrecer dos modos en la Fase 2:

        1) Líneas exactas: máxima fidelidad posicional, muchas cajas pequeñas.
        2) Párrafos editables: menos cajas, edición más natural.

    Retorno:
        [(page_width_pt, page_height_pt, [line_blocks]), ...]
    """
    import pdfplumber

    pages = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pw, ph = float(page.width), float(page.height)

            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
                extra_attrs=["size", "fontname"],
            )

            if not words:
                pages.append((pw, ph, []))
                continue

            # Agrupar primero por coordenada vertical aproximada.
            raw_lines: dict = {}
            for word in words:
                key = round(float(word.get("top", 0)) / 4) * 4
                raw_lines.setdefault(key, []).append(word)

            line_blocks = []

            for _, line_words in sorted(raw_lines.items()):
                line_words = sorted(line_words, key=lambda item: float(item["x0"]))

                # Si una misma altura contiene columnas o textos muy separados,
                # se separan en segmentos para no mezclar columnas o números.
                segments = []
                current_segment = []
                previous_x1 = None
                gap_limit = max(42.0, pw * 0.18)

                for word in line_words:
                    x0 = float(word["x0"])
                    if (
                        previous_x1 is not None
                        and current_segment
                        and (x0 - previous_x1) > gap_limit
                    ):
                        segments.append(current_segment)
                        current_segment = []
                    current_segment.append(word)
                    previous_x1 = float(word["x1"])

                if current_segment:
                    segments.append(current_segment)

                for segment in segments:
                    text = " ".join(str(word["text"]) for word in segment).strip()
                    if not text:
                        continue

                    x0 = float(min(word["x0"] for word in segment))
                    y0 = float(min(word.get("top", 0) for word in segment))
                    x1 = float(max(word["x1"] for word in segment))
                    y1 = float(max(word.get("bottom", y0 + 12) for word in segment))

                    first = segment[0]
                    size = float(first.get("size", 11) or 11)
                    fontname = str(first.get("fontname", "") or "")
                    bold = "Bold" in fontname or "bold" in fontname

                    line_blocks.append(
                        {
                            "text": text,
                            "x0": x0,
                            "y0": y0,
                            "x1": x1,
                            "y1": y1,
                            "size": size,
                            "bold": bold,
                            "height": max(y1 - y0, size, 8),
                            "kind": "line",
                        }
                    )

            pages.append((pw, ph, line_blocks))

    return pages


def _group_lines_into_paragraphs(line_pages):
    """
    Reconstruye párrafos aproximados a partir de líneas de texto PDF.

    v3.3 mejora dos puntos importantes:
        1) Conserva la lista interna de líneas originales dentro de cada párrafo.
           Esto permite mantener saltos de línea visuales dentro de una sola caja
           editable de PowerPoint, evitando que el layout se desacomode.
        2) Calcula ancho y alto usando la caja real del párrafo en el PDF, no un
           tamaño genérico. Así la ventana del párrafo queda mucho más correlacionada
           con el ancho visual original.

    Importante:
        Un PDF no conserva párrafos reales como un DOCX. Normalmente solo conserva
        caracteres/palabras con coordenadas. Por eso esta función usa heurísticas
        controladas: separación vertical, sangría, bullets, cambios de tamaño y
        títulos en negrita.
    """
    paragraph_pages = []

    for pw, ph, lines in line_pages:
        if not lines:
            paragraph_pages.append((pw, ph, []))
            continue

        paragraphs = []
        current = None

        for line in sorted(lines, key=lambda item: (item["y0"], item["x0"])):
            if current is None:
                current = {
                    "text": line["text"],
                    "lines": [line["text"]],
                    "line_heights": [float(line.get("height", 8))],
                    "line_widths": [float(line["x1"] - line["x0"])],
                    "x0": line["x0"],
                    "y0": line["y0"],
                    "x1": line["x1"],
                    "y1": line["y1"],
                    "size": line["size"],
                    "bold": line["bold"],
                    "last_line": line,
                    "line_count": 1,
                }
                continue

            previous = current["last_line"]
            vertical_gap = float(line["y0"] - previous["y1"])
            avg_height = max(float(previous.get("height", 8)), float(line.get("height", 8)), 8)
            indent_change = abs(float(line["x0"] - previous["x0"]))
            size_change = abs(float(line["size"] - previous["size"]))
            same_baseline = abs(float(line["y0"] - previous["y0"])) < avg_height * 0.40

            new_paragraph = False

            # Si hay dos segmentos en la misma altura pero muy separados en X,
            # probablemente son columnas/campos distintos y no un mismo párrafo.
            if same_baseline and float(line["x0"]) > float(previous["x1"]) + 35:
                new_paragraph = True

            # Separación vertical claramente mayor al interlineado normal.
            if vertical_gap > avg_height * 0.85:
                new_paragraph = True

            # Cambio fuerte de sangría: típico de título, lista o nuevo bloque.
            if indent_change > 28 and vertical_gap > avg_height * 0.18:
                new_paragraph = True

            # Listas: cada bullet/numeral conviene conservarlo como bloque editable.
            if _is_bullet_or_numbered(line["text"]):
                new_paragraph = True

            # Título breve en negrita seguido de texto: separar título y cuerpo.
            if previous["bold"] and len(previous["text"]) < 90 and vertical_gap > avg_height * 0.20:
                new_paragraph = True

            # Cambio importante de tamaño de fuente: probable encabezado o nota.
            if size_change >= 2.0 and vertical_gap > avg_height * 0.10:
                new_paragraph = True

            if new_paragraph:
                paragraphs.append(current)
                current = {
                    "text": line["text"],
                    "lines": [line["text"]],
                    "line_heights": [float(line.get("height", 8))],
                    "line_widths": [float(line["x1"] - line["x0"])],
                    "x0": line["x0"],
                    "y0": line["y0"],
                    "x1": line["x1"],
                    "y1": line["y1"],
                    "size": line["size"],
                    "bold": line["bold"],
                    "last_line": line,
                    "line_count": 1,
                }
            else:
                current["text"] = _join_paragraph_line(current["text"], line["text"])
                current["lines"].append(line["text"])
                current["line_heights"].append(float(line.get("height", 8)))
                current["line_widths"].append(float(line["x1"] - line["x0"]))
                current["x0"] = min(current["x0"], line["x0"])
                current["y0"] = min(current["y0"], line["y0"])
                current["x1"] = max(current["x1"], line["x1"])
                current["y1"] = max(current["y1"], line["y1"])
                current["size"] = min(current["size"], line["size"])
                current["bold"] = bool(current["bold"] and line["bold"])
                current["last_line"] = line
                current["line_count"] += 1

        if current:
            paragraphs.append(current)

        blocks = []
        for paragraph in paragraphs:
            line_widths = paragraph.get("line_widths", []) or [float(paragraph["x1"] - paragraph["x0"])]
            line_heights = paragraph.get("line_heights", []) or [float(paragraph["size"])]
            avg_line_height = sum(line_heights) / max(len(line_heights), 1)
            max_line_width = max(line_widths)

            blocks.append(
                {
                    "text": paragraph["text"].strip(),
                    "visual_text": "\n".join(paragraph.get("lines", [])).strip(),
                    "lines": list(paragraph.get("lines", [])),
                    "x0": float(paragraph["x0"]),
                    "y0": float(paragraph["y0"]),
                    "x1": float(paragraph["x1"]),
                    "y1": float(paragraph["y1"]),
                    "size": float(paragraph["size"]),
                    "bold": bool(paragraph["bold"]),
                    "kind": "paragraph",
                    "line_count": int(paragraph["line_count"]),
                    "avg_line_height": float(avg_line_height),
                    "max_line_width": float(max_line_width),
                }
            )

        paragraph_pages.append((pw, ph, blocks))

    return paragraph_pages

def _extract_text_blocks(pdf_path: str, mode: str = "paragraphs"):
    """
    Extrae bloques de texto con posición usando pdfplumber.

    Parámetros:
        mode="paragraphs"
            Agrupa líneas en párrafos editables. Es el nuevo modo recomendado.

        mode="lines"
            Mantiene el comportamiento original: una caja de texto por línea.
            Útil cuando se necesita máxima fidelidad posicional.
    """
    line_pages = _extract_text_lines(pdf_path)

    if mode == "lines":
        return line_pages

    return _group_lines_into_paragraphs(line_pages)


def _extract_pdf_image_regions(pdf_path: str, dpi: int = 150):
    """
    Extrae regiones de imagen visibles del PDF como objetos independientes.

    Se usa PyMuPDF porque permite localizar imágenes con coordenadas dentro de
    cada página. Para mantener la apariencia real, cada región se rasteriza como
    PNG usando un recorte exacto del PDF, no como página completa.

    Retorno:
        [
            {
                "page_index": 0,
                "x0": ..., "y0": ..., "x1": ..., "y1": ...,
                "png": bytes,
            },
            ...
        ]
    """
    import fitz  # PyMuPDF

    regions_by_page = []
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)

    with fitz.open(pdf_path) as doc:
        for page_index, page in enumerate(doc):
            page_regions = []
            page_area = float(page.rect.width * page.rect.height)
            seen = set()

            try:
                image_infos = page.get_image_info(xrefs=True)
            except TypeError:
                image_infos = page.get_image_info()

            for info in image_infos:
                bbox = info.get("bbox")
                if not bbox:
                    continue

                rect = fitz.Rect(bbox)
                rect = rect & page.rect

                # Filtrar imágenes diminutas, máscaras o residuos invisibles.
                if rect.width < 8 or rect.height < 8:
                    continue
                if (rect.width * rect.height) < 96:
                    continue

                # Deduplicación por coordenadas redondeadas.
                key = (
                    round(rect.x0, 1), round(rect.y0, 1),
                    round(rect.x1, 1), round(rect.y1, 1),
                    info.get("xref"),
                )
                if key in seen:
                    continue
                seen.add(key)

                # Si la imagen cubre casi toda la página es el fondo de página
                # completo (página escaneada o fondo plano). Se omite para no
                # duplicar toda la página como objeto independiente.
                if page_area and (rect.width * rect.height) / page_area > 0.995:
                    continue

                try:
                    pix = page.get_pixmap(matrix=matrix, clip=rect, alpha=False)
                    png_bytes = pix.tobytes("png")
                except Exception:
                    continue

                page_regions.append(
                    {
                        "page_index": page_index,
                        "x0": float(rect.x0),
                        "y0": float(rect.y0),
                        "x1": float(rect.x1),
                        "y1": float(rect.y1),
                        "png": png_bytes,
                    }
                )

            regions_by_page.append(page_regions)

    return regions_by_page


def _add_text_blocks_to_slide(slide, blocks, pw, ph, lft, top, iw, ih, slide_w,
                              text_mode: str, composition_mode: str,
                              preserve_line_breaks: bool):
    """
    Inserta bloques de texto en una diapositiva.

    En modo fondo fiel, el texto se mantiene blanco sobre la imagen base, como
    en el MVP original. En modo objetos independientes, no hay imagen base de
    página, por lo que el texto se inserta visible en negro.
    """
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor

    if not blocks or pw == 0 or ph == 0:
        return

    sx = iw / (pw * 12700)
    sy = ih / (ph * 12700)
    for blk in blocks:
        bx = lft + int(float(blk["x0"]) * 12700 * sx)
        by = top + int(float(blk["y0"]) * 12700 * sy)
        block_kind = str(blk.get("kind", text_mode))
        is_paragraph = block_kind == "paragraph"
        line_count = max(int(blk.get("line_count", 1)), 1)
        source_width_pt = max(float(blk.get("max_line_width", 0)), float(blk["x1"] - blk["x0"]))
        source_height_pt = max(float(blk["y1"] - blk["y0"]), float(blk.get("avg_line_height", blk["size"])) * line_count)

        if is_paragraph:
            # v3.3: ancho correlacionado directamente con el párrafo original.
            # Se evita agrandarlo de forma excesiva para que PowerPoint no cambie
            # demasiado la estructura visual del texto.
            bw = max(int(source_width_pt * 12700 * sx * 1.045), 120000)
            bh = max(int(source_height_pt * 12700 * sy * 1.22), 70000)
        else:
            bw = max(int((float(blk["x1"]) - float(blk["x0"])) * 12700 * sx * 1.08), 60000)
            bh = max(int((float(blk["y1"]) - float(blk["y0"])) * 12700 * sy * 1.10), 36000)

        # No dejar que la caja salga de la diapositiva por el lado derecho.
        max_available_w = max(slide_w - bx - 25000, 60000)
        bw = min(bw, max_available_w)

        tb = slide.shapes.add_textbox(Emu(bx), Emu(by), Emu(bw), Emu(bh))
        tf = tb.text_frame
        tf.word_wrap = True if is_paragraph else False

        # Márgenes internos mínimos para que la posición PDF → PPTX sea precisa.
        tf.margin_left = Emu(0)
        tf.margin_right = Emu(0)
        tf.margin_top = Emu(0)
        tf.margin_bottom = Emu(0)

        p = tf.paragraphs[0]
        run = p.add_run()

        if is_paragraph and preserve_line_breaks:
            # Una sola caja por párrafo, pero con saltos internos originales.
            # Esto mantiene el layout casi idéntico y evita cientos de mini cajas.
            run.text = str(blk.get("visual_text") or blk.get("text", ""))
        else:
            run.text = str(blk.get("text", ""))

        run.font.size = Pt(float(blk.get("size", 11) or 11))
        run.font.bold = bool(blk.get("bold", False))

        if composition_mode == "objects":
            # En modo objetos no hay fondo de página: texto visible.
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # type: ignore
        else:
            # En modo fondo fiel, se conserva la estrategia anterior.
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)  # type: ignore

        tb.fill.background()


def fase2_convert(src_path: str, dst_path: str, dpi: int = 150,
                  include_text: bool = True, text_mode: str = "paragraphs",
                  composition_mode: str = "background",
                  preserve_line_breaks: bool = True,
                  progress_cb=None) -> bool:
    from pypdf import PdfReader
    from pdf2image import convert_from_path
    from pptx import Presentation
    from pptx.util import Emu
    import io

    src = Path(src_path)
    reader = PdfReader(str(src))
    n = len(reader.pages)

    if progress_cb:
        progress_cb(5, "Leyendo PDF…")

    text_pages = []
    if include_text:
        if progress_cb:
            progress_cb(8, "Extrayendo texto y agrupando bloques…")
        try:
            text_pages = _extract_text_blocks(src_path, mode=text_mode)
        except Exception as e:
            if progress_cb:
                progress_cb(8, f"Extracción omitida: {e}")

    images = []
    image_regions = []

    if composition_mode == "objects":
        if progress_cb:
            progress_cb(12, "Extrayendo imágenes independientes del PDF…")
        try:
            image_regions = _extract_pdf_image_regions(src_path, dpi=dpi)
        except Exception as e:
            image_regions = [[] for _ in range(n)]
            if progress_cb:
                progress_cb(12, f"Extracción de imágenes omitida: {e}")
    else:
        if progress_cb:
            progress_cb(12, f"Renderizando {n} páginas a {dpi} DPI…")
        kwargs: dict = {"dpi": dpi}
        if POPPLER_PATH:
            kwargs["poppler_path"] = POPPLER_PATH
        images = convert_from_path(str(src), **kwargs)

    prs = Presentation()
    prs.slides._sldIdLst.clear()  # type: ignore
    blank = prs.slide_layouts[6]

    for i, pdf_page in enumerate(reader.pages):
        if progress_cb:
            progress_cb(12 + int(80 * ((i + 1) / max(n, 1))), f"Página {i+1}/{n}…")

        w_pt = float(pdf_page.mediabox.width)
        h_pt = float(pdf_page.mediabox.height)
        w_emu = int(w_pt * 12700)
        h_emu = int(h_pt * 12700)

        if i == 0:
            prs.slide_width = Emu(w_emu)
            prs.slide_height = Emu(h_emu)

        slide = prs.slides.add_slide(blank)
        slide_w = int(prs.slide_width)
        slide_h = int(prs.slide_height)

        # Si todas las páginas tienen el mismo tamaño, ratio=1. Si hay mezcla de
        # orientaciones o tamaños, se centra proporcionalmente.
        ratio = min(slide_w / w_emu, slide_h / h_emu)
        iw = int(w_emu * ratio)
        ih = int(h_emu * ratio)
        lft = (slide_w - iw) // 2
        top = (slide_h - ih) // 2

        if composition_mode == "objects":
            # ── Imágenes independientes ─────────────────────────────────────
            # No se inserta la imagen completa de la página. Solo se colocan las
            # imágenes internas detectadas: fotos, tablas JPG, gráficos raster, etc.
            page_regions = image_regions[i] if i < len(image_regions) else []
            sx = iw / (w_pt * 12700) if w_pt else 1
            sy = ih / (h_pt * 12700) if h_pt else 1

            for region in page_regions:
                rx = lft + int(float(region["x0"]) * 12700 * sx)
                ry = top + int(float(region["y0"]) * 12700 * sy)
                rw = int((float(region["x1"]) - float(region["x0"])) * 12700 * sx)
                rh = int((float(region["y1"]) - float(region["y0"])) * 12700 * sy)
                if rw <= 0 or rh <= 0:
                    continue
                buf = io.BytesIO(region["png"])
                slide.shapes.add_picture(buf, Emu(rx), Emu(ry), Emu(rw), Emu(rh))
        else:
            # ── Imagen de fondo completa ────────────────────────────────────
            # Modo de máxima fidelidad visual: el PPTX conserva el aspecto del PDF.
            if i < len(images):
                buf = io.BytesIO()
                images[i].save(buf, format="PNG")
                buf.seek(0)
                slide.shapes.add_picture(buf, Emu(lft), Emu(top), Emu(iw), Emu(ih))

        # ── Texto editable ──────────────────────────────────────────────────
        if include_text and i < len(text_pages):
            pw, ph, blocks = text_pages[i]
            _add_text_blocks_to_slide(
                slide=slide,
                blocks=blocks,
                pw=pw,
                ph=ph,
                lft=lft,
                top=top,
                iw=iw,
                ih=ih,
                slide_w=slide_w,
                text_mode=text_mode,
                composition_mode=composition_mode,
                preserve_line_breaks=preserve_line_breaks,
            )

    if progress_cb:
        progress_cb(95, "Guardando PPTX…")
    prs.save(dst_path)

    if progress_cb:
        if composition_mode == "objects":
            progress_cb(100, "✅  PPTX listo — objetos independientes")
        else:
            progress_cb(100, "✅  PPTX listo — fondo fiel + texto editable")
    return True


# ══════════════════════════════════════════════════════════════════════════════
#  UI — WIDGETS COMPARTIDOS
# ══════════════════════════════════════════════════════════════════════════════

def _open_folder(path: Path):
    folder = str(path.parent if path.is_file() else path)
    if   sys.platform == "win32":  os.startfile(folder)
    elif sys.platform == "darwin": subprocess.run(["open",      folder])
    else:                          subprocess.run(["xdg-open",  folder])
# ══════════════════════════════════════════════════════════════════════════════
#  MOTOR — RECUPERACIÓN DE TABLAS EXCEL
#  (basado en PDF Table Recovery Mini v0.1 — GoTopo 2026)
# ══════════════════════════════════════════════════════════════════════════════

import csv
import re as _re
from dataclasses import dataclass
from statistics import median as _median
from typing import List, Optional, Sequence, Tuple


@dataclass
class _Cell:
    text: str
    x0: float
    x1: float
    top: float
    bottom: float


@dataclass
class _RecoveredPage:
    page_number: int
    width: float
    height: float
    word_count: int
    rows: List[List[str]]


def _parse_crop(value: Optional[str]):
    if not value:
        return None
    parts = [float(x.strip()) for x in value.split(",")]
    if len(parts) != 4:
        raise ValueError("Recorte debe tener formato x0,y0,x1,y1")
    return tuple(parts)


def _looks_numeric_financial(text: str) -> bool:
    s = text.strip()
    if not any(ch.isdigit() for ch in s):
        return False
    if "%" in s:
        return False
    return bool(_re.fullmatch(r"[₡$€()\-\s.,0-9]+", s))


def _format_thousands_spaces(digits: str) -> str:
    if not digits:
        return ""
    groups: List[str] = []
    while digits:
        groups.append(digits[-3:])
        digits = digits[:-3]
    return " ".join(reversed(groups))


def _normalize_financial_cell(text: str) -> str:
    s = " ".join(text.strip().split())
    if not _looks_numeric_financial(s):
        return s
    has_colon  = "₡" in s
    has_dollar = "$" in s
    negative   = "(" in s and ")" in s
    leading_minus = s.lstrip().startswith("-")
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return s
    number = _format_thousands_spaces(digits)
    prefix = "₡" if has_colon else "$" if has_dollar else ""
    if negative:
        return f"({prefix}{number})"
    if leading_minus:
        return f"-{prefix}{number}"
    return f"{prefix}{number}"


def _get_words_table(page, max_font_size: float, crop):
    target = page.crop(crop) if crop else page
    words  = target.extract_words(
        x_tolerance=2, y_tolerance=3,
        keep_blank_chars=False, use_text_flow=False,
        extra_attrs=["size", "fontname"],
    ) or []
    filtered = []
    for w in words:
        text = str(w.get("text", "")).strip()
        if not text:
            continue
        size = float(w.get("size", 10) or 10)
        if size > max_font_size:
            continue
        filtered.append(w)
    return filtered


def _group_words_into_lines_table(words, y_tol: float):
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (float(w.get("top", 0)), float(w.get("x0", 0))))
    lines: List[List[dict]] = []
    for w in sorted_words:
        top    = float(w.get("top", 0))
        placed = False
        for line in lines:
            line_top = _median(float(x.get("top", 0)) for x in line)
            if abs(top - line_top) <= y_tol:
                line.append(w)
                placed = True
                break
        if not placed:
            lines.append([w])
    for line in lines:
        line.sort(key=lambda w: float(w.get("x0", 0)))
    lines.sort(key=lambda line: _median(float(w.get("top", 0)) for w in line))
    return lines


def _words_to_cells_table(line, gap_tol: float) -> List[_Cell]:
    if not line:
        return []
    cells: List[_Cell] = []
    cur_words: List[dict] = []

    def flush():
        if not cur_words:
            return
        text = " ".join(str(w["text"]) for w in cur_words)
        text = _normalize_financial_cell(text)
        cells.append(_Cell(
            text   = text,
            x0     = float(min(w["x0"]     for w in cur_words)),
            x1     = float(max(w["x1"]     for w in cur_words)),
            top    = float(min(w.get("top",    0) for w in cur_words)),
            bottom = float(max(w.get("bottom", 0) for w in cur_words)),
        ))
        cur_words.clear()

    prev_x1 = None
    for w in sorted(line, key=lambda x: float(x.get("x0", 0))):
        x0 = float(w["x0"])
        if prev_x1 is not None and (x0 - prev_x1) > gap_tol:
            flush()
        cur_words.append(w)
        prev_x1 = float(w["x1"])
    flush()
    return cells


def _recover_page_table(page, page_number: int, min_words: int,
                         max_font_size: float, y_tol: float,
                         gap_tol: float, crop) -> Optional[_RecoveredPage]:
    words = _get_words_table(page, max_font_size=max_font_size, crop=crop)
    if len(words) < min_words:
        return None
    lines = _group_words_into_lines_table(words, y_tol=y_tol)
    rows: List[List[str]] = []
    for line in lines:
        cells = _words_to_cells_table(line, gap_tol=gap_tol)
        if cells:
            rows.append([c.text for c in cells])
    table_like = sum(1 for r in rows if len(r) >= 2)
    if table_like < 5:
        return None
    return _RecoveredPage(
        page_number=page_number,
        width=float(page.width),
        height=float(page.height),
        word_count=len(words),
        rows=rows,
    )


def _export_csv_tables(pages: List[_RecoveredPage], out_dir: Path) -> List[Path]:
    out_paths = []
    for rp in pages:
        path = out_dir / f"pagina_{rp.page_number:03d}_tabla.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rp.rows)
        out_paths.append(path)
    return out_paths


def _export_xlsx_tables(pages: List[_RecoveredPage], out_path: Path) -> Optional[Path]:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
    except Exception as e:
        return None

    wb   = Workbook()
    thin = Side(style="thin", color="DDDDDD")
    header_fill = PatternFill("solid", fgColor="E8F1FF")

    if pages and wb.active:
        wb.remove(wb.active)

    for rp in pages:
        ws_name = _re.sub(r"[\\/*?:\[\]]", "_", f"Pagina_{rp.page_number}")[:31]
        ws = wb.create_sheet(ws_name)
        ws.append([f"Página {rp.page_number}", f"Palabras seleccionables: {rp.word_count}"])
        ws.append([])
        for row in rp.rows:
            ws.append(row)
        max_cols = max((len(r) for r in rp.rows), default=1)
        for col_idx in range(1, max_cols + 1):
            max_len = 8
            for row_idx in range(1, ws.max_row + 1):
                val = ws.cell(row_idx, col_idx).value
                if val is not None:
                    max_len = max(max_len, min(len(str(val)) + 2, 42))
                ws.cell(row_idx, col_idx).border = Border(
                    top=thin, left=thin, right=thin, bottom=thin)
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len
        for row_idx in range(3, min(ws.max_row, 5) + 1):
            for col_idx in range(1, max_cols + 1):
                c = ws.cell(row_idx, col_idx)
                c.font = Font(bold=True)
                c.fill = header_fill
        ws.freeze_panes = "A4"

    wb.save(out_path)
    return out_path


def tabla_excel_recover(src_path: str, out_dir: Path,
                         max_font_size: float = 16.0,
                         y_tol: float = 4.0,
                         gap_tol: float = 18.0,
                         min_words: int = 80,
                         crop_str: Optional[str] = None,
                         progress_cb=None) -> dict:
    """
    Orquestador de recuperación de tablas Excel desde PDF.
    Retorna dict con claves: recovered (int), xlsx (Path|None), csvs (list), msg (str)
    """
    import pdfplumber

    crop = _parse_crop(crop_str) if crop_str else None
    recovered: List[_RecoveredPage] = []

    if progress_cb: progress_cb(5, "Abriendo PDF…")

    with pdfplumber.open(src_path) as pdf:
        n = len(pdf.pages)
        if progress_cb: progress_cb(10, f"Analizando {n} páginas…")

        for i, page in enumerate(pdf.pages):
            pct = 10 + int(75 * (i / max(n, 1)))
            if progress_cb: progress_cb(pct, f"Escaneando página {i+1}/{n}…")

            rp = _recover_page_table(
                page=page,
                page_number=i + 1,
                min_words=min_words,
                max_font_size=max_font_size,
                y_tol=y_tol,
                gap_tol=gap_tol,
                crop=crop,
            )
            if rp:
                recovered.append(rp)

    if not recovered:
        if progress_cb: progress_cb(100, "⚠️  Sin tablas compatibles encontradas")
        return {"recovered": 0, "xlsx": None, "csvs": [], "msg": "sin_tablas"}

    out_dir.mkdir(parents=True, exist_ok=True)

    if progress_cb: progress_cb(88, "Exportando CSV…")
    csvs = _export_csv_tables(recovered, out_dir)

    if progress_cb: progress_cb(94, "Exportando Excel…")
    stem  = Path(src_path).stem
    xlsx  = _export_xlsx_tables(recovered, out_dir / f"{stem}_tablas_recuperadas.xlsx")

    if progress_cb: progress_cb(100, f"✅  {len(recovered)} tabla(s) recuperadas")
    return {
        "recovered": len(recovered),
        "xlsx": xlsx,
        "csvs": csvs,
        "msg":  "ok",
    }




# ── Interfaz gráfica: solo modo desktop ──
if _TKINTER_OK:


    class DropZone(Frame):
        def __init__(self, parent, allowed_ext: list, label: str,
                     on_file=None, **kw):
            super().__init__(parent, bg=C_CARD,
                             highlightbackground=C_BORDER,
                             highlightthickness=2, **kw)
            self.allowed_ext = [e.lower() for e in allowed_ext]
            self.on_file     = on_file
            self._path       = ""
            self._build(label)
            self._try_dnd()

        def _build(self, label: str):
            inner = Frame(self, bg=C_CARD, padx=22, pady=14)
            inner.pack(fill="both", expand=True)

            Label(inner, text="📂", font=("Segoe UI Emoji", 26),
                  bg=C_CARD, fg=C_ACCENT).pack()
            Label(inner, text=label,
                  font=("Segoe UI", 11, "bold"),
                  bg=C_CARD, fg=C_TEXT).pack(pady=(4, 1))
            Label(inner,
                  text="  •  ".join(e.upper() for e in self.allowed_ext),
                  font=("Segoe UI", 9), bg=C_CARD, fg=C_MUTED).pack()
            Label(inner, text="Arrastra aquí  o",
                  font=("Segoe UI", 9), bg=C_CARD, fg=C_MUTED).pack(pady=(8, 4))

            Button(inner, text="  Seleccionar archivo  ",
                   font=("Segoe UI", 10, "bold"),
                   bg=C_ACCENT, fg="white", activebackground="#3A6AEE",
                   relief="flat", padx=16, pady=6, cursor="hand2",
                   command=self._browse).pack()

            self.lbl = Label(inner, text="",
                              font=("Segoe UI", 9, "bold"),
                              bg=C_CARD, fg=C_GREEN, wraplength=400)
            self.lbl.pack(pady=(8, 0))

        def _try_dnd(self):
            try:
                self.drop_target_register("DND_Files")   # type: ignore
                self.dnd_bind("<<Drop>>", lambda e: self._set(e.data.strip().strip("{}")))  # type: ignore
                self.config(highlightbackground=C_ACCENT)
            except Exception:
                pass

        def _browse(self):
            flt = [(f"{e.upper()} files", f"*{e}") for e in self.allowed_ext]
            flt.append(("Todos los archivos", "*.*"))
            p = filedialog.askopenfilename(filetypes=flt)
            if p: self._set(p)

        def _set(self, path: str):
            if Path(path).suffix.lower() not in self.allowed_ext:
                messagebox.showerror("Formato no válido",
                    f"Selecciona: {', '.join(e.upper() for e in self.allowed_ext)}")
                return
            self._path = path
            self.lbl.config(text=f"✅  {Path(path).name}")
            self.config(highlightbackground=C_GREEN)
            if self.on_file: self.on_file(path)

        def get(self) -> str: return self._path
        def clear(self):
            self._path = ""
            self.lbl.config(text="")
            self.config(highlightbackground=C_BORDER)


    class StatusBar(Frame):
        def __init__(self, parent, **kw):
            super().__init__(parent, bg=C_PANEL, **kw)
            self._msg = StringVar(value="Listo")
            Label(self, textvariable=self._msg, bg=C_PANEL, fg=C_MUTED,
                  font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(4, 0))
            self.bar = Progressbar(self, mode="determinate", maximum=100)
            self.bar.pack(fill="x", padx=12, pady=(2, 6))

        def set_status(self, pct: int, msg: str):
            self.bar["value"] = pct
            self._msg.set(msg)
            self.update_idletasks()


    # ══════════════════════════════════════════════════════════════════════════════
    #  TAB FASE 1
    # ══════════════════════════════════════════════════════════════════════════════

    class Fase1Tab(Frame):
        def __init__(self, parent, **kw):
            super().__init__(parent, bg=C_PANEL, **kw)
            self._out = ""
            self._build()

        def _build(self):
            # Header tab
            hdr = Frame(self, bg=C_BG, pady=10)
            hdr.pack(fill="x")
            Label(hdr, text="FASE 1", font=("Segoe UI", 13, "bold"),
                  bg=C_BG, fg=C_GOLD).pack(side="left", padx=(20, 6))
            Label(hdr, text="Convertir documento a PDF",
                  font=("Segoe UI", 12), bg=C_BG, fg=C_TEXT).pack(side="left")

            # Separador
            Frame(self, bg=C_BORDER, height=1).pack(fill="x", padx=18)

            # Info box
            ib = Frame(self, bg=C_CARD2, padx=16, pady=10)
            ib.pack(fill="x", padx=20, pady=(14, 0))
            Label(ib, text="DOCX  →  PDF     •     PPTX  →  PDF",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_CARD2, fg=C_ACCENT).pack(anchor="w")
            Label(ib,
                  text="Con LibreOffice: texto + imágenes + tipografía original\n"
                       "Sin LibreOffice: solo texto (motor Python)",
                  font=("Segoe UI", 8), bg=C_CARD2, fg=C_MUTED,
                  justify="left").pack(anchor="w", pady=(4, 0))

            # Drop zone
            self.drop = DropZone(self, [".docx", ".pptx"],
                                 "Documento a convertir")
            self.drop.pack(fill="x", padx=20, pady=16, ipady=4)

            # Botón
            btn_row = Frame(self, bg=C_PANEL)
            btn_row.pack(pady=4)
            Button(btn_row, text="  ⚡  Convertir a PDF  ",
                   font=("Segoe UI", 12, "bold"),
                   bg=C_ACCENT, fg="white", activebackground="#3A6AEE",
                   relief="flat", padx=20, pady=9,
                   cursor="hand2", command=self._run).pack()

            # Status
            self.status = StatusBar(self)
            self.status.pack(fill="x", padx=18, pady=(12, 0))

            # Resultado
            self.lbl_res = Label(self, text="",
                                  font=("Segoe UI", 9),
                                  bg=C_PANEL, fg=C_GREEN, wraplength=480)
            self.lbl_res.pack(pady=6)

            Button(self, text="📂  Abrir carpeta de destino",
                   font=("Segoe UI", 9), bg=C_CARD, fg=C_MUTED,
                   activebackground=C_BORDER,
                   relief="flat", padx=12, pady=4,
                   cursor="hand2",
                   command=lambda: _open_folder(Path(self._out))).pack()

        def _run(self):
            src = self.drop.get()
            if not src:
                messagebox.showwarning("Sin archivo", "Selecciona un archivo primero.")
                return
            dst = filedialog.asksaveasfilename(
                title="Guardar PDF como…",
                initialfile=Path(src).stem + ".pdf",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")])
            if not dst: return
            self._out = dst
            self.lbl_res.config(text="")

            def run():
                try:
                    fase1_convert(src, dst,
                        progress_cb=lambda p, m: self.after(
                            0, lambda p=p, m=m: self.status.set_status(p, m)))
                    self.after(0, lambda: self.lbl_res.config(
                        text=f"✅  {Path(dst).name}  guardado", fg=C_GREEN))
                except Exception as e:
                    self.after(0, lambda: self.lbl_res.config(
                        text=f"❌  {e}", fg=C_WARN))
                    self.after(0, lambda: self.status.set_status(0, "Error"))

            threading.Thread(target=run, daemon=True).start()


    # ══════════════════════════════════════════════════════════════════════════════
    #  TAB FASE 2
    # ══════════════════════════════════════════════════════════════════════════════

    class Fase2Tab(Frame):
        def __init__(self, parent, **kw):
            super().__init__(parent, bg=C_PANEL, **kw)
            self._out = ""
            self._build()

        def _build(self):
            # Header tab
            hdr = Frame(self, bg=C_BG, pady=10)
            hdr.pack(fill="x")
            Label(hdr, text="FASE 2", font=("Segoe UI", 13, "bold"),
                  bg=C_BG, fg=C_GOLD).pack(side="left", padx=(20, 6))
            Label(hdr, text="Descomponer PDF en PowerPoint editable",
                  font=("Segoe UI", 12), bg=C_BG, fg=C_TEXT).pack(side="left")

            Frame(self, bg=C_BORDER, height=1).pack(fill="x", padx=18)

            # Info box
            ib = Frame(self, bg=C_CARD2, padx=16, pady=10)
            ib.pack(fill="x", padx=20, pady=(14, 0))
            Label(ib, text="PDF  →  PPTX  (una diapositiva por página)",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_CARD2, fg=C_ACCENT).pack(anchor="w")
            Label(ib,
                  text="• Modo fondo fiel: imagen completa de página + texto editable\n"
                       "• Modo objetos: sin imagen base, con fotos/tablas/gráficos como objetos independientes\n"
                       "• Texto en párrafos: una caja editable por bloque, con ancho correlacionado al original\n"
                       "• Reordena, elimina o agrega páginas en PowerPoint",
                  font=("Segoe UI", 8), bg=C_CARD2, fg=C_MUTED,
                  justify="left").pack(anchor="w", pady=(4, 0))

            # Poppler status
            pcolor = C_GREEN if POPPLER_PATH else C_WARN
            pmsg   = (f"✅  Poppler detectado" if POPPLER_PATH
                      else "⚠️  Poppler no encontrado — requerido solo para modo fondo fiel")
            Label(self, text=pmsg, font=("Segoe UI", 8),
                  bg=C_PANEL, fg=pcolor).pack(anchor="w", padx=22, pady=(8, 0))

            # Opciones
            opt = Frame(self, bg=C_PANEL)
            opt.pack(fill="x", padx=22, pady=(8, 0))

            # DPI
            dpi_row = Frame(opt, bg=C_PANEL)
            dpi_row.pack(anchor="w")
            Label(dpi_row, text="Calidad imagen:",
                  bg=C_PANEL, fg=C_TEXT, font=("Segoe UI", 9)).pack(side="left")
            self.dpi_var = StringVar(value="150")
            for v, l in [("96","Baja"), ("150","Media ✓"), ("200","Alta"), ("300","Máxima")]:
                ttk.Radiobutton(dpi_row, text=l,
                                 variable=self.dpi_var, value=v).pack(side="left", padx=5)

            # Toggle texto
            txt_row = Frame(opt, bg=C_PANEL)
            txt_row.pack(anchor="w", pady=(6, 0))
            self.txt_var = StringVar(value="1")
            ttk.Checkbutton(txt_row,
                             text="Incluir texto editable",
                             variable=self.txt_var,
                             onvalue="1", offvalue="0").pack(side="left")

            # Modo de agrupación del texto embebido.
            # Párrafos: recomendado para edición real.
            # Líneas: conserva el comportamiento original del MVP.
            mode_row = Frame(opt, bg=C_PANEL)
            mode_row.pack(anchor="w", pady=(6, 0))
            Label(mode_row, text="Agrupar texto:",
                  bg=C_PANEL, fg=C_TEXT, font=("Segoe UI", 9)).pack(side="left")
            self.text_mode_var = StringVar(value="paragraphs")
            for value, label in [("paragraphs", "Párrafos editables ✓"),
                                 ("lines", "Líneas exactas")]:
                ttk.Radiobutton(mode_row, text=label,
                                 variable=self.text_mode_var,
                                 value=value).pack(side="left", padx=5)

            # Composición de salida.
            # Fondo fiel: comportamiento estable de v3.2.
            # Objetos: elimina la imagen base y extrae fotos/tablas JPG/gráficos raster.
            comp_row = Frame(opt, bg=C_PANEL)
            comp_row.pack(anchor="w", pady=(6, 0))
            Label(comp_row, text="Composición:",
                  bg=C_PANEL, fg=C_TEXT, font=("Segoe UI", 9)).pack(side="left")
            self.composition_var = StringVar(value="background")
            for value, label in [("background", "Fondo fiel"),
                                 ("objects", "Objetos independientes EXP")]:
                ttk.Radiobutton(comp_row, text=label,
                                 variable=self.composition_var,
                                 value=value).pack(side="left", padx=5)

            # v3.3: conserva saltos visuales dentro de una sola caja de párrafo.
            # Mejora mucho la estabilidad de ancho/estructura sin volver a mini cajas.
            br_row = Frame(opt, bg=C_PANEL)
            br_row.pack(anchor="w", pady=(6, 0))
            self.line_breaks_var = StringVar(value="1")
            ttk.Checkbutton(br_row,
                             text="Mantener saltos visuales dentro del párrafo",
                             variable=self.line_breaks_var,
                             onvalue="1", offvalue="0").pack(side="left")

            # Drop zone
            self.drop = DropZone(self, [".pdf"], "PDF a descomponer")
            self.drop.pack(fill="x", padx=20, pady=12, ipady=4)

            # Botón
            btn_row = Frame(self, bg=C_PANEL)
            btn_row.pack(pady=4)
            Button(btn_row, text="  🔓  Descomponer PDF  ",
                   font=("Segoe UI", 12, "bold"),
                   bg=C_GREEN, fg="#071A12", activebackground="#00A87E",
                   relief="flat", padx=20, pady=9,
                   cursor="hand2", command=self._run).pack()

            # Status
            self.status = StatusBar(self)
            self.status.pack(fill="x", padx=18, pady=(12, 0))

            # Resultado
            self.lbl_res = Label(self, text="",
                                  font=("Segoe UI", 9),
                                  bg=C_PANEL, fg=C_GREEN, wraplength=480)
            self.lbl_res.pack(pady=6)

            Button(self, text="📂  Abrir carpeta de destino",
                   font=("Segoe UI", 9), bg=C_CARD, fg=C_MUTED,
                   activebackground=C_BORDER,
                   relief="flat", padx=12, pady=4,
                   cursor="hand2",
                   command=lambda: _open_folder(Path(self._out))).pack()

        def _run(self):
            src = self.drop.get()
            if not src:
                messagebox.showwarning("Sin archivo", "Selecciona un PDF primero.")
                return

            composition_mode = self.composition_var.get()
            if composition_mode == "background" and not POPPLER_PATH:
                messagebox.showerror("Poppler no encontrado",
                    "Verifica la instalación de Poppler en C:\\Poppler\\ o usa el modo Objetos independientes.")
                return

            dst = filedialog.asksaveasfilename(
                title="Guardar PPTX como…",
                initialfile=Path(src).stem + "_editable.pptx",
                defaultextension=".pptx",
                filetypes=[("PowerPoint", "*.pptx")])
            if not dst: return

            self._out               = dst
            dpi                     = int(self.dpi_var.get())
            include_text            = self.txt_var.get() == "1"
            text_mode               = self.text_mode_var.get()
            preserve_line_breaks    = self.line_breaks_var.get() == "1"
            self.lbl_res.config(text="")

            def run():
                try:
                    fase2_convert(src, dst, dpi=dpi,
                        include_text=include_text,
                        text_mode=text_mode,
                        composition_mode=composition_mode,
                        preserve_line_breaks=preserve_line_breaks,
                        progress_cb=lambda p, m: self.after(
                            0, lambda p=p, m=m: self.status.set_status(p, m)))

                    if composition_mode == "objects":
                        base = "objetos independientes"
                    else:
                        base = "fondo fiel"

                    if include_text and text_mode == "paragraphs":
                        txt = "párrafos editables"
                    elif include_text:
                        txt = "líneas exactas"
                    else:
                        txt = "sin texto editable"

                    self.after(0, lambda: self.lbl_res.config(
                        text=f"✅  {Path(dst).name}  ({base} + {txt})", fg=C_GREEN))
                except Exception as e:
                    self.after(0, lambda: self.lbl_res.config(
                        text=f"❌  {e}", fg=C_WARN))
                    self.after(0, lambda: self.status.set_status(0, "Error"))

            threading.Thread(target=run, daemon=True).start()


    # ══════════════════════════════════════════════════════════════════════════════
    #  VENTANA PRINCIPAL
    # ══════════════════════════════════════════════════════════════════════════════

    class App(Tk):
        def __init__(self):
            super().__init__()
            self.title("PDF Quick Disassembler")
            self.geometry("660x800")
            self.minsize(600, 730)
            self.configure(bg=C_BG)
            self._styles()
            self._build()

        def _styles(self):
            s = ttk.Style(self)
            s.theme_use("clam")
            s.configure("TNotebook",     background=C_BG,    borderwidth=0)
            s.configure("TNotebook.Tab", background=C_PANEL, foreground=C_MUTED,
                         padding=[18, 9], font=("Segoe UI", 10, "bold"))
            s.map("TNotebook.Tab",
                  background=[("selected", C_CARD)],
                  foreground=[("selected", C_GOLD)])
            s.configure("Horizontal.TProgressbar",
                         troughcolor=C_BORDER, background=C_ACCENT,
                         borderwidth=0, thickness=9)
            for w in ("TRadiobutton", "TCheckbutton"):
                s.configure(w, background=C_PANEL, foreground=C_TEXT,
                             font=("Segoe UI", 9))
                s.map(w, background=[("active", C_PANEL)],
                          foreground=[("active", C_GREEN)])

        def _build(self):
            # ── Splash header ─────────────────────────────────────────────────────
            splash = Frame(self, bg=C_BG, pady=0)
            splash.pack(fill="x")

            # Banda de color superior
            band = Frame(splash, bg=C_GOLD, height=4)
            band.pack(fill="x")

            content = Frame(splash, bg=C_BG, pady=14)
            content.pack(fill="x")

            # Logo + título
            left = Frame(content, bg=C_BG)
            left.pack(side="left", padx=(20, 0))
            Label(left, text="⚡", font=("Segoe UI Emoji", 32),
                  bg=C_BG, fg=C_GOLD).pack(side="left", padx=(0, 10))

            titles = Frame(content, bg=C_BG)
            titles.pack(side="left")
            Label(titles, text="PDF Quick Disassembler",
                  font=("Segoe UI", 18, "bold"),
                  bg=C_BG, fg=C_TEXT).pack(anchor="w")
            Label(titles, text="by  GoTopo  •  v3.4 Tabla Excel  •  2026 R.Montagne",
                  font=("Segoe UI", 9),
                  bg=C_BG, fg=C_GOLD).pack(anchor="w")
            Label(titles,
                  text="Desarma, edita y reconstruye documentos PDF",
                  font=("Segoe UI", 9),
                  bg=C_BG, fg=C_MUTED).pack(anchor="w", pady=(2, 0))

            # Línea separadora
            Frame(self, bg=C_GOLD, height=1).pack(fill="x")

            # ── Notebook ──────────────────────────────────────────────────────────
            nb = Notebook(self)
            nb.pack(fill="both", expand=True, pady=0)
            nb.add(Fase1Tab(nb),       text="   📄  FASE 1  —  a PDF   ")
            nb.add(Fase2Tab(nb),       text="   🔓  FASE 2  —  a PowerPoint   ")
            nb.add(TablaExcelTab(nb),  text="   📊  PDF Table Recovery   ")

            # ── Footer ────────────────────────────────────────────────────────────
            Frame(self, bg=C_GOLD, height=1).pack(fill="x")
            footer = Frame(self, bg=C_BG, pady=6)
            footer.pack(fill="x")
            Label(footer,
                  text="GoTopo  •  Democratizando el acceso a la justicia  •  Costa Rica",
                  font=("Segoe UI", 8), bg=C_BG, fg=C_MUTED).pack()


    # ══════════════════════════════════════════════════════════════════════════════
    #  TAB — TABLA EXCEL
    # ══════════════════════════════════════════════════════════════════════════════

    class TablaExcelTab(Frame):
        """Pestaña para recuperar tablas financieras desde PDFs con texto embebido."""

        def __init__(self, parent, **kw):
            super().__init__(parent, bg=C_PANEL, **kw)
            self._out_dir: Optional[Path] = None
            self._build()

        def _build(self):
            # Header
            hdr = Frame(self, bg=C_BG, pady=10)
            hdr.pack(fill="x")
            Label(hdr, text="PDF Table Recovery", font=("Segoe UI", 13, "bold"),
                  bg=C_BG, fg=C_GOLD).pack(side="left", padx=(20, 6))
            Label(hdr, text="Recuperar tablas financieras desde PDF",
                  font=("Segoe UI", 12), bg=C_BG, fg=C_TEXT).pack(side="left")

            Frame(self, bg=C_BORDER, height=1).pack(fill="x", padx=18)

            # ── Advertencia crítica ───────────────────────────────────────────────
            warn_frame = Frame(self, bg="#2A1E0A", padx=16, pady=10)
            warn_frame.pack(fill="x", padx=20, pady=(14, 0))

            Label(warn_frame,
                  text="⚠️  SOLO para tablas copiadas directamente desde Excel",
                  font=("Segoe UI", 10, "bold"),
                  bg="#2A1E0A", fg=C_GOLD, justify="left").pack(anchor="w")
            Label(warn_frame,
                  text="Funciona cuando en Excel se seleccionó el rango → Copy → Paste en PowerPoint\n"
                       "→ se convirtió a PDF con LibreOffice o PDF Quick Disassembler.\n"
                       "NO funciona con tablas guardadas como PNG, JPG o captura de pantalla.",
                  font=("Segoe UI", 8), bg="#2A1E0A", fg="#C8A060",
                  justify="left").pack(anchor="w", pady=(4, 0))

            # ── Info de salida ────────────────────────────────────────────────────
            info = Frame(self, bg=C_CARD2, padx=16, pady=8)
            info.pack(fill="x", padx=20, pady=(8, 0))
            Label(info,
                  text="PDF  →  XLSX + CSV  (carpeta automática en el Desktop)",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_CARD2, fg=C_ACCENT).pack(anchor="w")
            Label(info,
                  text="• Escanea todas las páginas automáticamente\n"
                       "• Genera un archivo Excel con una hoja por tabla encontrada\n"
                       "• También genera un CSV por página como respaldo",
                  font=("Segoe UI", 8), bg=C_CARD2, fg=C_MUTED,
                  justify="left").pack(anchor="w", pady=(3, 0))

            # ── Opciones avanzadas ────────────────────────────────────────────────
            adv_frame = Frame(self, bg=C_PANEL)
            adv_frame.pack(fill="x", padx=20, pady=(8, 0))

            Label(adv_frame, text="Opciones avanzadas:",
                  font=("Segoe UI", 9, "bold"),
                  bg=C_PANEL, fg=C_MUTED).pack(anchor="w")

            row1 = Frame(adv_frame, bg=C_PANEL)
            row1.pack(anchor="w", pady=(4, 0))

            # Tolerancia de gap
            Label(row1, text="Separación columnas:",
                  bg=C_PANEL, fg=C_TEXT, font=("Segoe UI", 9)).pack(side="left")
            self.gap_var = StringVar(value="18")
            ttk.Radiobutton(row1, text="Estrecha (12)",
                             variable=self.gap_var, value="12").pack(side="left", padx=4)
            ttk.Radiobutton(row1, text="Normal (18) ✓",
                             variable=self.gap_var, value="18").pack(side="left", padx=4)
            ttk.Radiobutton(row1, text="Amplia (28)",
                             variable=self.gap_var, value="28").pack(side="left", padx=4)

            row2 = Frame(adv_frame, bg=C_PANEL)
            row2.pack(anchor="w", pady=(4, 0))
            Label(row2, text="Tamaño máx. fuente tabla:",
                  bg=C_PANEL, fg=C_TEXT, font=("Segoe UI", 9)).pack(side="left")
            self.font_var = StringVar(value="16")
            for v, l in [("12", "12pt"), ("16", "16pt ✓"), ("22", "22pt")]:
                ttk.Radiobutton(row2, text=l,
                                 variable=self.font_var, value=v).pack(side="left", padx=4)

            # ── Drop zone ─────────────────────────────────────────────────────────
            self.drop = DropZone(self, [".pdf"], "PDF con tablas de Excel")
            self.drop.pack(fill="x", padx=20, pady=12, ipady=4)

            # ── Botón ─────────────────────────────────────────────────────────────
            btn_row = Frame(self, bg=C_PANEL)
            btn_row.pack(pady=4)
            Button(btn_row, text="  📊  Recuperar Tablas Excel  ",
                   font=("Segoe UI", 12, "bold"),
                   bg=C_GOLD, fg="#1A1200", activebackground="#D4A800",
                   relief="flat", padx=20, pady=9,
                   cursor="hand2", command=self._run).pack()

            # ── Status ────────────────────────────────────────────────────────────
            self.status = StatusBar(self)
            self.status.pack(fill="x", padx=18, pady=(12, 0))

            # ── Resultado ─────────────────────────────────────────────────────────
            self.lbl_res = Label(self, text="",
                                  font=("Segoe UI", 9),
                                  bg=C_PANEL, fg=C_GREEN, wraplength=500)
            self.lbl_res.pack(pady=6)

            Button(self, text="📂  Abrir carpeta de salida",
                   font=("Segoe UI", 9), bg=C_CARD, fg=C_MUTED,
                   activebackground=C_BORDER,
                   relief="flat", padx=12, pady=4,
                   cursor="hand2",
                   command=self._open_out).pack()

        def _get_desktop_output(self, pdf_stem: str) -> Path:
            """Crea carpeta de salida automáticamente en el Desktop."""
            desktop = Path.home() / "Desktop"
            if not desktop.exists():
                desktop = Path.home()
            folder = desktop / f"TablaExcel_{pdf_stem}"
            folder.mkdir(parents=True, exist_ok=True)
            return folder

        def _run(self):
            src = self.drop.get()
            if not src:
                messagebox.showwarning("Sin archivo",
                                       "Selecciona un PDF primero.")
                return

            out_dir = self._get_desktop_output(Path(src).stem)
            self._out_dir = out_dir
            self.lbl_res.config(text="")

            gap_tol      = float(self.gap_var.get())
            max_font_size = float(self.font_var.get())

            def run():
                try:
                    result = tabla_excel_recover(
                        src_path=src,
                        out_dir=out_dir,
                        gap_tol=gap_tol,
                        max_font_size=max_font_size,
                        progress_cb=lambda p, m: self.after(
                            0, lambda p=p, m=m: self.status.set_status(p, m)),
                    )
                    if result["msg"] == "sin_tablas":
                        self.after(0, lambda: self.lbl_res.config(
                            text="⚠️  No se encontraron tablas compatibles.\n"
                                 "Verifica que la tabla haya sido copiada directamente desde Excel.",
                            fg=C_WARN))
                        return

                    n    = result["recovered"]
                    xlsx = result["xlsx"]
                    msg  = f"✅  {n} tabla(s) recuperada(s)  →  {out_dir.name}  (Desktop)"
                    if xlsx:
                        msg += f"\n📊  {xlsx.name}"
                    self.after(0, lambda: self.lbl_res.config(text=msg, fg=C_GREEN))

                except Exception as e:
                    self.after(0, lambda: self.lbl_res.config(
                        text=f"❌  Error: {e}", fg=C_WARN))
                    self.after(0, lambda: self.status.set_status(0, "Error"))

            threading.Thread(target=run, daemon=True).start()

        def _open_out(self):
            if self._out_dir and self._out_dir.exists():
                _open_folder(self._out_dir)
            else:
                folder = Path.home() / "Desktop"
                _open_folder(folder)


    # ══════════════════════════════════════════════════════════════════════════════
    def main():
        missing = []
        for mod, pkg in [("pypdf","pypdf"), ("docx","python-docx"),
                         ("pptx","python-pptx"), ("PIL","Pillow"),
                         ("pdf2image","pdf2image"), ("reportlab","reportlab"),
                         ("pdfplumber","pdfplumber"), ("fitz","PyMuPDF"),
                         ("openpyxl","openpyxl")]:
            try: __import__(mod)
            except ImportError: missing.append(pkg)
        if missing:
            print("⚠️  Falta instalar:  pip install", " ".join(missing))
        App().mainloop()

    if __name__ == "__main__":
        main()
