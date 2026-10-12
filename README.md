# PDF Quick Disassembler
**by GoTopo  •  v3.1 MVP  •  2026**

> Desarma, edita y reconstruye documentos PDF

---

## ¿Qué hace?

| FASE | Entrada | Salida | Descripción |
|------|---------|--------|-------------|
| **1** | `.docx` / `.pptx` | `.pdf` | Convierte documentos a PDF con texto embebido |
| **2** | `.pdf` | `.pptx` editable | Una diapositiva por página, con imagen de fondo y texto editable transparente |

**Flujo completo de edición:**
```
PDF original
    │
    ▼  FASE 2
PPTX editable  ←─── editar, reordenar, eliminar páginas en PowerPoint
    │
    ▼  FASE 1
Nuevo PDF listo
```

---

## Instalación

### Requisitos del sistema

**Poppler** (requerido para Fase 2):
- Descargar desde: https://github.com/oschwartz10612/poppler-windows/releases
- Extraer en `C:\Poppler\`
- El programa lo detecta automáticamente

**LibreOffice** (recomendado para Fase 1):
- Descargar desde: https://www.libreoffice.org/download/
- Sin LibreOffice: solo exporta texto (sin imágenes ni tipografía original)

### Dependencias Python

```bash
pip install pypdf python-pptx python-docx pdf2image Pillow reportlab pdfplumber
```

### OCR de tablas en imágenes (opcional, solo escritorio)

La Fase 3 puede leer tablas pegadas como imagen con la opción
«Leer también tablas en imágenes (OCR)». Usa PaddleOCR-VL y PP-OCRv5 en GPU NVIDIA,
instalados en un entorno aparte (`venv_ocr`) para no cargar la app ni la versión web:

```powershell
powershell -ExecutionPolicy Bypass -File install_ocr.ps1   # crea venv_ocr; reintenta si se corta la red
```

Si `venv_ocr` no existe la opción aparece desactivada. Otra ubicación: variable
`PQD_OCR_PYTHON` con la ruta al `python.exe` del entorno OCR. La primera ejecución
descarga los modelos (~2 GB).

### Ejecutar

```bash
python topo_vs_pdf.py
```

---

## Estructura del proyecto

```
PDF_Quick_Disassembler/
├── topo_vs_pdf.py      ← aplicación principal
├── requirements.txt    ← dependencias
├── README.md
└── .vscode/
    ├── launch.json     ← F5 para ejecutar en VS Code
    └── settings.json
```

---

*GoTopo  •  Democratizando el acceso a la justicia  •  Costa Rica  •  2026*
