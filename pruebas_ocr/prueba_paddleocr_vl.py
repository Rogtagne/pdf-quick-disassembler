"""Prueba de PaddleOCR-VL sobre un PDF o una imagen con tablas.

Uso (con el entorno venv_ocr):
    venv_ocr\\Scripts\\python.exe pruebas_ocr\\prueba_paddleocr_vl.py RUTA [--salida CARPETA]

Guarda por cada página un .md (texto + tablas) y un .json en la carpeta de salida,
y muestra el tiempo por página. La primera ejecución descarga los pesos del modelo (~2 GB).
"""
import argparse
import sys
import time
from pathlib import Path


def verificar_gpu():
    import paddle
    print(f"PaddlePaddle {paddle.__version__}")
    if not paddle.device.is_compiled_with_cuda():
        print("AVISO: Paddle no tiene soporte CUDA; se usará la CPU (mucho más lento).")
        return "cpu"
    n = paddle.device.cuda.device_count()
    print(f"CUDA disponible: {n} GPU(s) — {paddle.device.cuda.get_device_name(0)}")
    return "gpu:0"


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ruta", help="PDF o imagen (png/jpg) a leer")
    ap.add_argument("--salida", default=None,
                    help="Carpeta de resultados (por defecto: pruebas_ocr/resultados/<nombre>)")
    args = ap.parse_args()

    src = Path(args.ruta)
    if not src.exists():
        sys.exit(f"No existe: {src}")
    out = Path(args.salida) if args.salida else Path(__file__).parent / "resultados" / src.stem.strip()  # Windows no admite carpetas que terminen en espacio
    out.mkdir(parents=True, exist_ok=True)

    device = verificar_gpu()

    import paddleocr
    from paddleocr import PaddleOCRVL
    print(f"PaddleOCR {paddleocr.__version__}")

    t0 = time.perf_counter()
    pipeline = PaddleOCRVL(device=device)
    print(f"Modelo cargado en {time.perf_counter() - t0:.1f} s")

    t0 = time.perf_counter()
    n = 0
    for res in pipeline.predict(str(src)):
        n += 1
        res.save_to_markdown(save_path=str(out))
        res.save_to_json(save_path=str(out))
    total = time.perf_counter() - t0

    print(f"\n{n} página(s) en {total:.1f} s ({total / max(n, 1):.2f} s/página)")
    print(f"Resultados en: {out}")


if __name__ == "__main__":
    main()
