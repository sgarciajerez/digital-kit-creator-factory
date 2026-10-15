$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot
if (-not (Test-Path '.venv\Scripts\python.exe')) {
    throw 'No existe .venv. Ejecuta primero: py -3.12 -m venv .venv'
}
& .venv\Scripts\python.exe -m streamlit run .\scripts\app.py
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

