$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Local .venv is missing. Run .\scripts\setup_windows.ps1 first."
}

$OutputDir = ".\outputs\release_diagnostics"
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$Destination = Join-Path $OutputDir "environment_packages.txt"
uv pip freeze --python ".\.venv\Scripts\python.exe" | Out-File -FilePath $Destination -Encoding utf8
Write-Host "Environment package snapshot written to: $Destination"
