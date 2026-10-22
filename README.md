# VSM Engineering Data Post-Processing

Professional deterministic Python tool for VSM simulation/test post-processing and engineering reporting.

## Current release: 1.3.0 - Dynamic Vehicle Reporting

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

The production client workflow supports Diesel, Hybrid and Electric. After upload,
the UI defaults to Sergio's 15 October 2026 definitions in
`config/report_profiles/sergio_2026_10/`, aligned with Diesel `_04` and
Electric/Hybrid `_08`. The original profiles remain explicitly available as
legacy reporting. Powertrain selection and source channels determine compatibility;
machine names remain editable report metadata.

Electric and Hybrid report configuration includes an optional **Battery pack mass (kg)**
field. Enter a finite positive mass to display it as the seventh card on Battery &
Electrical Energy System. Blank input displays N/A; no mass is inferred from the
machine filename. Manual input is recorded as user-provided in report metadata.
The Agro Discharge Rate card uses net initial-minus-final remaining chemical mass
[kg] divided by final cumulative distance [km], without requiring chemical density.
Nonzero distance starts and possible refills are reported in the diagnostics.
The existing volumetric diagnostics remain available separately in Excel/metadata.

Revised reports add measured gearbox torque and high-voltage auxiliaries,
soil-resistance power/energy plots, configurable Diesel coolant temperature and
live Excel integration/statistic formulas. Energy uses observed-time trapezoidal
integration. Missing optional measurements disable their dependent sections;
measured zeros remain zeros. Channels, statistics, charts, analysis intervals and
existing PowerPoint plot slots are selectable. See the
[implementation and validation report](docs/SERGIO_2026_10_15_IMPLEMENTATION.md)
for exact conventions, tests and unresolved upstream measurements.

The October client update retains application version 1.3.0 and adds live formulas
for all applicable configured Excel KPIs. The existing Diesel thermal slide shows
independent oil/water plots and four MIN/MAX cards; its compact Excel header is
24 pt. Diesel and Electric use updated real VSM CSVs for release validation.
The available Caiman SP Hybrid reference also passes report-generation checks;
its negative inferred chemical contents correctly prevent a full-cycle rate.
A newer client Hybrid export beyond that reference is not claimed as tested.
See [client release notes](docs/FINAL_RELEASE_NOTES.md) for the validation scope
and operating instructions.

The 16 October plot update adds four separate soil-resistance plots versus time
to current Electric and Hybrid Excel and PowerPoint reports, reusing existing
power and cumulative-energy calculations. Existing combined charts remain.
Two Astauto pages precede the final summary; complete Electric reports now have
25 Excel charts and 15 PowerPoint slides after the final Sergio refinements. The portable folder and archive are
named `Astauto_VSM_October_22102026` for the final 22 October delivery.
The two soil pages also include six engineering KPI cards each, using the Wheel
Steering Angles card style and existing signed power/energy/distance conventions.
The soil cards reuse existing Excel calculations and layouts.

The final 17 October refinements display the exact input basename on every
profile cover, add a tyre power/energy detail page before Electric's soil
pages, and exclude Engine Load Torque from Diesel reports while retaining
measured Engine Torque and throttle. Following Sergio's clarification, tyre
consumption and the detail page now reuse rolling-resistance power and energy,
which already include bulldozing and compaction. The four redundant additive
channels are removed; separate soil diagnostics remain. See the
[18 October correction audit](docs/TYRE_ROLLING_CORRECTION_2026_10_18.txt).

A local Streamlit UI exposes the workflow without requiring Python or YAML editing.

## Client start (recommended)

On Windows, first double-click:

```text
SETUP_VSM_TOOL.bat
```

Setup detects a validated Python 3.11.x runtime, offers to install the supported
official CPython 3.11.9 runtime if one is missing, creates the package-local
`.venv`, installs runtime dependencies, and runs the health check.

After setup reports success, double-click:

```text
START_VSM_TOOL.bat
```

The launcher:

1. uses the package-local `.venv`;
2. runs the client-readiness health check;
3. opens the UI in the default browser.

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
  SETUP_VSM_TOOL.bat
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
