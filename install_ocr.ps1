# Instala el entorno OCR opcional de la Fase 3 (venv_ocr): PaddlePaddle GPU (CUDA 12.9)
# + PaddleOCR (doc-parser) + PyMuPDF. Resistente a cortes de red:
#   - el wheel grande de Paddle (~1 GB) se baja con curl reanudable (lo descargado queda en disco);
#   - el resto viene de PyPI con pip, reintentando cada paso hasta 30 veces.
# Uso:  powershell -ExecutionPolicy Bypass -File install_ocr.ps1
# Avance en venv_ocr\_install\install.log; termina con "TERMINADO OK" o "TERMINADO CON ERROR".
$ErrorActionPreference = "Continue"
$root  = $PSScriptRoot
$venv  = "$root\venv_ocr"
$py    = "$venv\Scripts\python.exe"
$dir   = "$venv\_install"
$log   = "$dir\install.log"
$whl   = "$dir\wheels\paddlepaddle_gpu-3.4.0-cp312-cp312-win_amd64.whl"
$url   = "https://paddle-whl.cdn.bcebos.com/stable/cu129/paddlepaddle-gpu/paddlepaddle_gpu-3.4.0-cp312-cp312-win_amd64.whl"
$size  = 1054892336
$pipOpts = @("--timeout", "120", "--retries", "20", "--resume-retries", "50", "--progress-bar", "off")

if (-not (Test-Path $py)) { py -3.12 -m venv $venv }
New-Item -ItemType Directory -Force "$dir\wheels" | Out-Null

function Log($msg) { "[$(Get-Date -Format HH:mm:ss)] $msg" | Out-File $log -Append -Encoding utf8 }

function Download-Paddle {
    for ($i = 1; $i -le 200; $i++) {
        $have = if (Test-Path $whl) { (Get-Item $whl).Length } else { 0 }
        if ($have -eq $size) { Log "descarga paddle - completa"; return $true }
        Log ("descarga paddle - intento $i, llevamos {0:N0} de {1:N0} MB" -f ($have / 1MB), ($size / 1MB))
        curl.exe -sS -L -C - --retry 5 --retry-all-errors --retry-delay 10 --speed-limit 10000 --speed-time 60 -o $whl $url 2>&1 | Out-File $log -Append -Encoding utf8
        Start-Sleep -Seconds 5
    }
    Log "descarga paddle - FALLO tras 200 intentos"
    return $false
}

function Step($name, [string[]]$pipArgs) {
    for ($i = 1; $i -le 30; $i++) {
        Log "$name - intento $i"
        & $py -m pip install @pipArgs @pipOpts 2>&1 | Out-File $log -Append -Encoding utf8
        if ($LASTEXITCODE -eq 0) { Log "$name - OK"; return $true }
        Start-Sleep -Seconds 20
    }
    Log "$name - FALLO tras 30 intentos"
    return $false
}

function Fin($ok) {
    $estado = if ($ok) { "TERMINADO OK" } else { "TERMINADO CON ERROR" }
    "=== $estado $(Get-Date) ===" | Out-File $log -Append -Encoding utf8
    Write-Host "$estado  (detalle en $log)"
    exit ([int](-not $ok))
}

"=== Inicio $(Get-Date) ===" | Out-File $log -Append -Encoding utf8
if (-not (Download-Paddle)) { Fin $false }
if (-not (Step "paddlepaddle-gpu (dependencias)" @($whl, "--extra-index-url", "https://www.paddlepaddle.org.cn/packages/stable/cu129/"))) { Fin $false }
if (-not (Step "paddleocr + pymupdf" @("paddleocr[doc-parser]", "pymupdf"))) { Fin $false }
Fin $true
