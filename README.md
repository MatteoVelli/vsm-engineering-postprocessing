# VSM Engineering Data Post-Processing

Professional deterministic Python tool for VSM simulation/test post-processing and engineering reporting.

## Current release: 1.2.10 - Profile Reporting

The validated workflow is:

```text
VSM CSV/XLSX
    ->
Inspection
    ->
Profile channel mapping
    ->
Math channels
    ->
Statistics and KPIs
    ->
Plots
    ->
Excel report
    ->
Optional PowerPoint report
```

The production client workflow uses the RoboSprayer Electric and Hybrid reporting profiles in `config/report_profiles/`. Those profiles are synced to Sergio's latest Electric_03 and Hybrid_04 Excel templates, including optional `Track_Height` / `Road Height` support when the source data includes it.

A local Streamlit UI exposes the workflow without requiring Python or YAML editing.

## Client start (recommended)

On Windows, double-click:

```text
START_VSM_TOOL.bat
```

The launcher:

1. checks that `uv` is available;
2. creates a local Python 3.11 `.venv` when required;
3. installs/updates the application dependencies;
4. runs the client-readiness health check;
5. opens the UI in the default browser.

See `docs/CLIENT_QUICK_START.md` for the client workflow.

## Developer setup

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\setup_windows.ps1
.\.venv\Scripts\Activate.ps1
```

Run the health check:

```powershell
.\scripts\run_doctor.ps1
```

Launch the UI:

```powershell
.\scripts\run_ui.ps1
```

Run the configured pipeline:

```powershell
.\scripts\run_end_to_end.ps1
```

## Traceability and diagnostics

Each complete pipeline run writes:

```text
pipeline_manifest.json
pipeline_summary.txt
pipeline.log
```

The manifest records source/config hashes, software version, Python/runtime details, timestamps, per-stage durations, stage outputs/metrics/status, and final report paths.

A failed run retains completed-stage outputs and diagnostic metadata. Unexpected Python/runtime exceptions are also captured in the pipeline diagnostics before the error is returned to the user.

## Safe output handling

When `clean_before_run: true`, the pipeline removes only known generated VSM artifacts such as `01_inspection`, `02_channel_selection`, etc. It does not delete the whole output directory, so unrelated/user-created files at the output root are preserved.

## Repository structure

```text
vsm-post-processing/
  .streamlit/
  config/
  docs/
  reference_files/
  outputs/
  scripts/
  src/vsm_postprocessing/
  tests/
  START_VSM_TOOL.bat
  CHANGELOG.md
  pyproject.toml
  README.md
```

## Individual engineering commands

```powershell
.\scripts\run_source_inspection.ps1
.\scripts\run_channel_selection.ps1
.\scripts\run_math_channels.ps1
.\scripts\run_statistics.ps1
.\scripts\run_plotting.ps1
.\scripts\run_excel_report.ps1
.\scripts\run_powerpoint_report.ps1
```

## Regression tests

```powershell
pytest
```

The acceptance tests use reference files when they are present in `reference_files/`. Client/reference data are intentionally excluded from release ZIPs and Git.

Build the deterministic client release package with:

```powershell
.\scripts\build_release.ps1
```

This creates a clean client ZIP plus a SHA-256 checksum in `dist/`. The ZIP contains its own `RELEASE_MANIFEST.json` with per-file hashes. To record the exact packages installed on a validated machine, run `scripts/snapshot_environment.ps1`.

## Engineering safeguards

- stable semantic profile mappings instead of ambiguous display-name matching;
- explicit RAW/MATH provenance;
- deterministic formulas/statistics;
- strict dependency and numeric validation;
- no hidden smoothing/resampling in plots;
- reporting layers do not replace engineering calculations;
- source/config hashes and run manifests retained for traceability;
- client-readable stage diagnostics and run log;
- safe output cleanup that preserves unrelated files;
- AI is not part of engineering calculation or selection logic.
