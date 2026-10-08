# Client Quick Start

## First-Time Setup

1. Extract the complete ZIP into a writable Windows folder; do not run it inside the ZIP.
2. Open the extracted folder.
3. Double-click `SETUP_VSM_TOOL.bat`.
4. Follow the on-screen instructions.
5. When setup reports SUCCESS, double-click `START_VSM_TOOL.bat`.

Setup automatically checks Python, offers to install the supported official CPython 3.11.9 runtime if a validated Python 3.11.x runtime is missing, creates the package-local `.venv`, installs Streamlit and all required libraries, and validates the installation. No manual `pip`, `python -m venv`, Streamlit, PowerShell, YAML editing or PATH configuration is required for normal use.

The first setup requires internet access if Python or Python packages are not already available locally.

## Normal Use

1. Double-click `START_VSM_TOOL.bat`.
2. Select **Profile-driven VSM report**.
3. Upload the source VSM `.xlsx` or `.csv` results file.
4. Choose **Diesel**, **Hybrid** or **Electric**, and confirm the detected machine name.
5. Click **STEP 3 - Validate**. Resolve any missing required channels or invalid data before continuing; the UI lists the affected channels and blocks report generation.
6. Click **STEP 4 - Generate Engineering Report**.
7. Download the Excel report and, when needed, the PowerPoint report from **STEP 5 - Access Output**. The profile workflow generates both reports; using the PowerPoint is optional.

The launcher opens the local Streamlit interface in your browser. If the browser does not open automatically, use `http://localhost:8501`. Keep the launcher running while using the application; close it when finished.

## Workflows

**Profile-driven VSM report** generates the client Excel report and optional PowerPoint directly from the uploaded source data and selected reporting profile.

**Custom Analysis** analyzes the uploaded file exactly as supplied and keeps the manual channel/math/statistics/plot/report controls for engineering investigation.

## Where Runs Are Stored

UI runs are isolated under:

```text
outputs/ui_runs/<timestamp>/
```

Profile reports are saved inside the run's `profile_excel_report/` and `profile_powerpoint_report/` folders. You can also save copies anywhere using the browser download buttons. Custom Analysis pipeline runs include their manifests, summaries and logs.

## Equations in Excel

Select a mathematical data cell to inspect its live Excel formula when one is available. Every mathematical column also has an equation note on its main-sheet heading; hover over the heading or show its note in Excel. The mapping worksheet retains the full equation inventory.

Pointwise arithmetic formulas and the upper RMS formulas recalculate in Excel. Integration and cumulative channels retain validated Python values. After changing simulation input data, regenerate the report from the updated source file to refresh those channels and the complete report consistently.

## Setup Problems

If Windows blocks the downloaded archive, right-click the ZIP, select **Properties**, use **Unblock** if offered, and extract again. Follow your organisation's application policy; do not disable security software. Run setup from a writable folder and keep the package files together.

First-time setup needs access to the official Python download and Python package servers. If dependency installation fails, inspect `logs/setup.log`, check internet/proxy permissions, and rerun setup. If downloads are blocked by company policy, ask for an approved Python 3.11 installation or an offline dependency installation. See `docs/TROUBLESHOOTING.md` for details.

## Health Check

From PowerShell:

```powershell
.\.venv\Scripts\python.exe -m vsm_postprocessing.doctor_cli --project-root .
```

`PASS` means there are no blocking installation/configuration problems. `WARN` items are informational/non-blocking. `FAIL` must be corrected before production use.
