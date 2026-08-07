$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Virtual environment not found. Run .\scripts\setup_windows.ps1 first."
}

& ".\.venv\Scripts\python.exe" -m streamlit run ".\src\vsm_postprocessing\ui_app.py" --server.headless false
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
