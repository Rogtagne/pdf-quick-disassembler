
C:\Users\Rogelio Montagne\OneDrive - CFZ\Documentos\GoTopo

C:\Users\Rogelio Montagne\OneDrive - CFZ\Documentos\GoTopo\topo_vs_pdf




#Activar
venv\Scripts\activate

#correr el script con 
cd topo_vs_pdf
python web_app.py

python topo_vs_pdf

https://pdf-quick-disassembler.onrender.com/


==> Available at your primary URL https://pdf-quick-disassembler.onrender.com

https://pdf-quick-disassembler.onrender.com

Para cambios desde visual code
=====================================================================

git add topo_vs_pdf.py
git commit -m "Fix: detectar Poppler en Linux /usr/bin para Render.com"
git push

nota git pull   en casos de error antes de hacer el commit
=====================================================================
La estrategia es: **un solo repositorio, dos modos de ejecución**. Los motores de conversión ya están escritos y probados — solo agregamos una capa web encima.Ahora construyo el servidor Flask:Ahora la interfaz web — misma estética GoTopo:Ahora los archivos de despliegue para Render.com:Proyecto completo listo. Ahora la guía de mantenimiento que pediste:

---

## Estructura final del proyecto

```
topo_vs_pdf/
├── topo_vs_pdf.py        ← MOTORES (Fase 1, Fase 2, Table Recovery) — no tocar
├── web_app.py            ← Servidor Flask — rutas y API
├── templates/
│   └── index.html        ← Interfaz web (HTML)
├── static/
│   ├── css/app.css       ← Estilos GoTopo
│   └── js/app.js         ← Lógica del navegador
├── Procfile              ← Instrucción de arranque para Render.com
├── render.yaml           ← Configuración automática de Render.com
└── requirements.txt      ← Todas las dependencias
```

---

## Cómo probarlo localmente antes de subir

```bash
cd topo_vs_pdf
python web_app.py
```
Abre `http://localhost:5000` — funciona igual que la app de escritorio.

---

## Despliegue en Render.com — paso a paso

**1.** Crear cuenta gratuita en `render.com`

**2.** Subir el proyecto a GitHub:
```bash
git init
git add .
git commit -m "PDF Quick Disassembler v3.4 Web"
git remote add origin https://github.com/TU_USUARIO/pdf-quick-disassembler.git
git push -u origin main
```

**3.** En Render.com: `New → Web Service → Connect GitHub → seleccionar el repo`

**4.** Render detecta el `render.yaml` automáticamente. Solo confirmar y hacer clic en **Deploy**. Poppler se instala solo en el servidor Linux.

**5.** En 3-5 minutos la app está en línea en una URL tipo `https://pdf-quick-disassembler.onrender.com`

---

## Mantenimiento en el tiempo

| Tarea | Archivo a editar |
|-------|-----------------|
| Cambiar colores o textos de la web | `static/css/app.css` o `templates/index.html` |
| Agregar una nueva función de conversión | `topo_vs_pdf.py` (motor) + nueva ruta en `web_app.py` + botón en `index.html` |
| Cambiar límite de tamaño de archivo | Variable `PQD_MAX_MB` en `render.yaml` |
| Actualizar una dependencia Python | `requirements.txt` → `git push` → Render redespliega automáticamente |
| Ver logs de errores en producción | Panel de Render.com → tu servicio → pestaña **Logs** |

Cada vez que hagas `git push`, Render redespliega automáticamente. No hay que hacer nada más.
 
Para borrar todas las carpetas __pycache__ de forma recursiva desde la raíz de tu proyecto, puedes usar una combinación de comandos de PowerShell.
Get-ChildItem -Path . -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -eq '__pycache__'} | Remove-Item -Recurse -Force
