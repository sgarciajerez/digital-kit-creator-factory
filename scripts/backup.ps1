param(
    [string]$Destination = (Join-Path (Split-Path -Parent $PSScriptRoot) ("backups\factory-" + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.zip'))
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

$destinationDirectory = Split-Path -Parent $Destination
New-Item -ItemType Directory -Force -Path $destinationDirectory | Out-Null

$items = Get-ChildItem -Force | Where-Object {
    $_.Name -notin @('.env', '.venv', '.git', '.pytest_cache', '__pycache__', 'backups', 'runtime')
}
Compress-Archive -Path $items.FullName -DestinationPath $Destination -Force
Write-Host "Copia creada: $Destination"
