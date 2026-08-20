# Client Quick Start

## Normal use

1. Extract the complete project folder to a writable local location.
2. Double-click `START_VSM_TOOL.bat`.
3. On first launch, the tool creates/updates its local `.venv`, performs a health check, and opens the UI in the default browser.
4. Select **Profile-driven VSM report**.
5. Choose the Electric or Hybrid reporting profile and confirm the detected machine name.
6. Upload the source VSM `.xlsx` or `.csv` results file.
7. Click **Generate Engineering Report**.
8. Download or open the final Excel/PowerPoint reports from the UI.

No Python or YAML editing is required for normal use. The first setup/update may require internet access so `uv` can obtain missing Python packages.

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
.\scripts\run_doctor.ps1
```

`PASS` means there are no blocking installation/configuration problems. `WARN` items are informational/non-blocking. `FAIL` must be corrected before production use.
