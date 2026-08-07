$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Virtual environment not found. Run .\scripts\setup_windows.ps1 first."
}

& ".\.venv\Scripts\python.exe" -m vsm_postprocessing.doctor_cli --project-root "."
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
