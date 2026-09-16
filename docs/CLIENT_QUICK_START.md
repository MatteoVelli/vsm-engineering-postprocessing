# Client Quick Start

## First-Time Setup

1. Extract the ZIP file.
2. Open the extracted folder.
3. Double-click `SETUP_VSM_TOOL.bat`.
4. Follow the on-screen instructions.
5. When setup reports SUCCESS, double-click `START_VSM_TOOL.bat`.

Setup automatically checks Python, offers to install the supported official CPython 3.11.9 runtime if a validated Python 3.11.x runtime is missing, creates the package-local `.venv`, installs Streamlit and all required libraries, and validates the installation. No manual `pip`, `python -m venv`, Streamlit, PowerShell, YAML editing or PATH configuration is required for normal use.

The first setup requires internet access if Python or Python packages are not already available locally.

## Normal Use

1. Double-click `START_VSM_TOOL.bat`.
2. Select **Profile-driven VSM report**.
3. Choose the Electric or Hybrid reporting profile and confirm the detected machine name.
4. Upload the source VSM `.xlsx` or `.csv` results file.
5. Click **Generate Engineering Report**.
6. Download or open the final Excel/PowerPoint reports from the UI.

## Workflows

**Profile-driven VSM report** generates the client Excel report and optional PowerPoint directly from the uploaded source data and selected reporting profile.

**Custom Analysis** analyzes the uploaded file exactly as supplied and keeps the manual channel/math/statistics/plot/report controls for engineering investigation.

## Where Runs Are Stored

UI runs are isolated under:

```text
outputs/ui_runs/<timestamp>/
```

Each run contains runtime configuration, intermediate deterministic outputs, `pipeline_manifest.json`, `pipeline_summary.txt`, and `pipeline.log`.

## Health Check

From PowerShell:

```powershell
.\.venv\Scripts\python.exe -m vsm_postprocessing.doctor_cli --project-root .
```

`PASS` means there are no blocking installation/configuration problems. `WARN` items are informational/non-blocking. `FAIL` must be corrected before production use.
