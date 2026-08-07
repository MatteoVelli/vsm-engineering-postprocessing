# Client Quick Start

## Normal use

1. Extract the complete project folder to a writable local location.
2. Double-click `START_VSM_TOOL.bat`.
3. On first launch, the tool creates/updates its local `.venv`, performs a health check, and opens the UI in the default browser.
4. Upload a VSM `.xlsx` or `.csv` file.
5. Choose export channels, math channels, statistics, plots, Excel channels and optional PowerPoint.
6. Click **Run complete pipeline**.
7. Download or open the final Excel/PowerPoint reports from the UI.

No Python or YAML editing is required for normal use. The first setup/update may require internet access so `uv` can obtain missing Python packages.

## Where runs are stored

UI runs are isolated under:

```text
outputs/ui_runs/<timestamp>/
```

Each run contains runtime configuration, intermediate deterministic outputs, `pipeline_manifest.json`, `pipeline_summary.txt`, and `pipeline.log`.

## Health check

From PowerShell:

```powershell
.\scripts\run_doctor.ps1
```

`PASS` means there are no blocking installation/configuration problems. `WARN` items are informational/non-blocking. `FAIL` must be corrected before production use.

## Engineering traceability

For every run, retain the complete run directory. The manifest records source/configuration hashes, software/Python/platform details, stage status and execution time.
