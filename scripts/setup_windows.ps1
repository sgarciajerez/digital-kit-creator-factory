$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

if (-not (Test-Path '.venv\Scripts\python.exe')) {
    $pythonCandidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'),
        'C:\Program Files\Python312\python.exe'
    )
    $pythonExe = $pythonCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($pythonExe) {
        & $pythonExe -m venv .venv
    } else {
        py -3.12 -m venv .venv
    }
}

& .venv\Scripts\python.exe -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw 'No se pudo actualizar pip' }
& .venv\Scripts\python.exe -m pip install -r requirements.txt -c constraints-windows-py312.txt
if ($LASTEXITCODE -ne 0) { throw 'No se pudieron instalar las dependencias' }
& .venv\Scripts\python.exe scripts/check_environment.py --excel
if ($LASTEXITCODE -ne 0) { throw 'El entorno no está preparado para generación nativa' }
Write-Host 'Entorno preparado. Ejecuta .\scripts\run_app.ps1 para iniciar la interfaz.'
