# Client Quick Start

## Normal use

1. Extract the complete project folder to a writable local location.
2. Double-click `START_VSM_TOOL.bat`.
3. On first launch, the tool creates/updates its local `.venv`, performs a health check, and opens the UI in the default browser.
4. Leave the default workflow set to **Engineering Report**.
5. Confirm the selected scenario is **Caiman SP Hybrid - 6 Field Cycles + Road Transfer**.
6. Upload the source VSM `.xlsx` or `.csv` results file.
7. Click **Generate Engineering Report**.
8. Download or open the final Excel/PowerPoint reports from the UI.

No Python or YAML editing is required for normal use. The first setup/update may require internet access so `uv` can obtain missing Python packages.

## Workflows

**Engineering Report** generates the configured complete Caiman SP Hybrid mission report. It uses the uploaded VSM source results plus packaged scenario assets; no reference workbook upload is required.

**Custom Analysis** analyzes the uploaded file exactly as supplied and keeps the manual channel/math/statistics/plot/report controls for engineering investigation.

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
