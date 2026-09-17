param(
    [string]$ProjectRoot
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
}

try {
    if ($ProjectRoot.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ge 0) {
        throw "The path contains invalid characters, including a possible literal quote."
    }
    $ProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
    if (-not (Test-Path -LiteralPath $ProjectRoot -PathType Container)) {
        throw "The project directory does not exist."
    }
} catch {
    Write-Host "Invalid ProjectRoot: $ProjectRoot" -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host "Use START_VSM_TOOL.bat or SETUP_VSM_TOOL.bat from the extracted package."
    exit 2
}
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SetupScript = Join-Path $ProjectRoot "scripts\client_setup.ps1"
$LogDir = Join-Path $ProjectRoot "logs"
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

Set-Location -LiteralPath $ProjectRoot

Write-Host ""
Write-Host "Running system health check..."

if (-not (Test-Path $VenvPython)) {
    Write-Host ""
    Write-Host "VSM Engineering Post-Processing is not set up yet." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please double-click SETUP_VSM_TOOL.bat first."
    Write-Host "Setup creates the package-local Python environment and installs the required dependencies."
    exit 2
}

& $VenvPython -m vsm_postprocessing.doctor_cli --project-root $ProjectRoot
if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "The health check found a blocking problem. The UI will not start." -ForegroundColor Red
    Write-Host "Run SETUP_VSM_TOOL.bat again. Technical details are in the Tool\logs folder."
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "Starting VSM user interface..." -ForegroundColor Green
$AppPath = Join-Path $ProjectRoot "src\vsm_postprocessing\ui_app.py"
& $VenvPython -m streamlit run $AppPath --server.headless false
exit $LASTEXITCODE
