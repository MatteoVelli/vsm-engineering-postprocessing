# Troubleshooting

## UI does not start

Run:

```powershell
.\scripts\run_doctor.ps1
```

If the virtual environment is missing, run `START_VSM_TOOL.bat` or `scripts/setup_windows.ps1`.

## First launch cannot download dependencies

The source-code distribution uses `uv` to create/update the local environment. If required packages are not already cached, the first setup needs network access. A fully standalone/offline executable is a separate packaging phase.

## `uv` is not recognised

The one-click launcher requires `uv` on PATH to create/update the local environment. If `.venv` already exists, the Python commands can still run directly; otherwise install `uv` or contact the maintainer.

## Pipeline fails

Open the run directory and inspect, in this order:

1. `pipeline_summary.txt`
2. `pipeline.log`
3. `pipeline_manifest.json`
4. the output directory of the failed stage

The pipeline stops at the first failed stage and retains all completed-stage outputs.

## Excel/PowerPoint cannot be overwritten

Close the previously generated report if it is open in Excel or PowerPoint, then run again.

## A channel/statistic/plot is unavailable in the UI

The UI filters options against the channels actually present in the uploaded source and the dependencies of configured math channels. This prevents invalid selections. Check the source channel catalogue if the expected signal is missing.

## Output cleaning

`clean_before_run: true` removes only known VSM-generated stage folders and pipeline metadata. Unknown/user-created files in the output root are deliberately preserved.
