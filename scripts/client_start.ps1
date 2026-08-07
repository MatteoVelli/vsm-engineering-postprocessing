$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

Write-Host ""
Write-Host "VSM Engineering Post-Processing" -ForegroundColor Cyan
Write-Host "Client launcher" -ForegroundColor Cyan
Write-Host ""

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: uv is not available in PATH." -ForegroundColor Red
    Write-Host "Install uv or ask the project maintainer for the packaged runtime instructions."
    exit 2
}

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    Write-Host "Creating local Python 3.11 environment..."
    uv venv .venv --python 3.11
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Write-Host "Checking/installing application dependencies..."
uv pip install --python ".\.venv\Scripts\python.exe" -e "."
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host ""
Write-Host "Running system health check..."
& ".\.venv\Scripts\python.exe" -m vsm_postprocessing.doctor_cli --project-root "."
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "The health check found a blocking problem. The UI will not start." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Starting VSM user interface..." -ForegroundColor Green
& ".\.venv\Scripts\python.exe" -m streamlit run ".\src\vsm_postprocessing\ui_app.py" --server.headless false
exit $LASTEXITCODE
