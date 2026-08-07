$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Virtual environment not found. Run .\scripts\setup_windows.ps1 first."
}

& ".\.venv\Scripts\python.exe" -m vsm_postprocessing.pipeline_cli ".\config\end_to_end_example.yaml"
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
