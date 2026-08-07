$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

if (-not (Test-Path ".\.venv\Scripts\python.exe")) {
    throw "Local .venv is missing. Run .\scripts\setup_windows.ps1 first."
}

& ".\.venv\Scripts\python.exe" -m vsm_postprocessing.release_cli --project-root "." --output-dir ".\dist"
exit $LASTEXITCODE
