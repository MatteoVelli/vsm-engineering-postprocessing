# Troubleshooting

## UI does not start

Run:

```powershell
.\.venv\Scripts\python.exe -m vsm_postprocessing.doctor_cli --project-root .
```

If the virtual environment is missing, double-click `SETUP_VSM_TOOL.bat`.

## First-Time Setup

Extract the ZIP, open the extracted folder, and double-click `SETUP_VSM_TOOL.bat`.
Setup checks for a validated Python 3.11.x runtime, offers to install the
supported official CPython 3.11.9 runtime if one is missing, creates `.venv`,
installs dependencies and runs the health check. No PATH changes are required.

## Python is not found

If setup says no compatible VSM runtime was found, choose `Y` when prompted to
let setup download and install the supported CPython 3.11.9 runtime for this
extracted tool folder. This uses the official installer from `python.org`,
installs without administrator rights in the normal case, includes pip, and does
not rely on the Microsoft Store or PATH changes.

Manual fallback:

1. Install 64-bit Python 3.11.x from `https://www.python.org/downloads/windows/`.
2. Rerun `SETUP_VSM_TOOL.bat`.
3. Do not install Streamlit or other packages manually; setup installs the application dependencies into `.venv`.

## No internet or company download block

The first setup needs internet access if Python or packages are not already
cached. If setup cannot download Python or dependencies, check the internet
connection and rerun `SETUP_VSM_TOOL.bat`. If a company network blocks software
downloads, use the manual Python fallback above or ask the maintainer for an
offline wheelhouse package.

## Virtual environment is missing

Run `SETUP_VSM_TOOL.bat`. `START_VSM_TOOL.bat` intentionally does not create the
environment; it shows a clear message asking you to run setup first.

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
