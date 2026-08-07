$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    throw "uv is not available in PATH. This project setup uses uv."
}

if (-not (Test-Path ".venv")) {
    uv venv .venv --python 3.11
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

uv pip install --python ".\.venv\Scripts\python.exe" -e ".[dev]"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Environment ready."
Write-Host "Activate it with: .\.venv\Scripts\Activate.ps1"
