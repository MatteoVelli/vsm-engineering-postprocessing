# Milestone 09 — Simple User UI

## Goal
Provide a first client-usable interface that configures and runs the already validated deterministic processing pipeline without requiring edits to Python or YAML.

## Architecture
The UI is a local Streamlit application. It does not perform engineering calculations itself. It:

1. accepts a CSV/XLSX VSM file;
2. runs the validated inspection layer;
3. presents stable channel IDs with display names and units;
4. lets the user separately select export channels, math channels, statistics, KPI values, plots and Excel report channels;
5. creates a versioned runtime configuration bundle;
6. calls the existing `run_pipeline()` orchestrator;
7. shows stage status and provides the final Excel workbook for download/opening.

Math dependencies required by statistics, plots or report columns are automatically added. This prevents an apparently valid UI selection from producing a broken calculation graph.

## Launch

```powershell
.\scripts\setup_windows.ps1
.\.venv\Scripts\Activate.ps1
.\scripts\run_ui.ps1
```

The browser opens on the local Streamlit server.

## Persistence
`Save selections` writes `config/ui_saved_profile.yaml`. The profile stores only reusable channel/statistics/plot selections, not source data.

## Outputs
Each run is isolated under:

```text
outputs/ui_runs/<timestamp>/
├── runtime_config/
└── results/
    ├── 01_inspection/
    ├── 02_channel_selection/
    ├── 03_math_channels/
    ├── 04_statistics/
    ├── 05_plots/
    ├── 06_excel_report/
    ├── pipeline_manifest.json
    └── pipeline_summary.txt
```

## Scope boundary
This is intentionally a simple configuration UI. It does not yet provide a formula editor for arbitrary new math channels, custom plot construction from scratch, PowerPoint generation, or AI-assisted recommendations.
