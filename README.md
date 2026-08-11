# VSM Engineering Data Post-Processing

Professional deterministic Python tool for VSM simulation/test post-processing and engineering reporting.

## Current release: 1.2.5 - Full Duty-Cycle Pipeline Integration

The validated workflow is:

```text
VSM CSV/XLSX
    ↓
Inspection
    ↓
Channel selection
    ↓
Math channels
    ↓
Statistics
    ↓
Plots
    ↓
Excel report
    ↓
Optional PowerPoint report
```

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

Run the complete configured pipeline:

```powershell
.\scripts\run_end_to_end.ps1
```

## Traceability and diagnostics

Each complete pipeline run now writes:

```text
pipeline_manifest.json
pipeline_summary.txt
pipeline.log
```

The manifest records:

- source SHA-256;
- configuration SHA-256 values;
- software version;
- Python version and executable;
- operating-system/platform information;
- UTC start/end timestamps;
- total and per-stage durations;
- stage outputs/metrics/status;
- final Excel and optional PowerPoint paths.

A failed run retains completed-stage outputs and diagnostic metadata. Unexpected Python/runtime exceptions are also captured in the pipeline diagnostics before the error is returned to the user.

## Safe output handling

When `clean_before_run: true`, the pipeline removes only known generated VSM artifacts such as `01_inspection`, `02_channel_selection`, etc. It **does not delete the whole output directory**, so unrelated/user-created files at the output root are preserved.

## Repository structure

```text
vsm-post-processing/
├── .streamlit/
├── config/
├── docs/
├── reference_files/
├── outputs/
├── scripts/
├── src/vsm_postprocessing/
├── tests/
├── START_VSM_TOOL.bat
├── CHANGELOG.md
├── pyproject.toml
└── README.md
```

## Individual engineering commands

```powershell
.\scripts\run_source_inspection.ps1
.\scripts\run_channel_selection.ps1
.\scripts\run_math_channels.ps1
.\scripts\run_statistics.ps1
.\scripts\run_statistics_reference_validation.ps1
.\scripts\run_plotting.ps1
.\scripts\run_plotting_reference_report.ps1
.\scripts\run_excel_report.ps1
.\scripts\run_powerpoint_report.ps1
```

## Regression tests

```powershell
pytest
```

The acceptance tests use Sergio's original reference files when they are present in `reference_files/`. Client/reference data are intentionally excluded from release ZIPs and Git.

Build the deterministic client release package with:

```powershell
.\scripts\build_release.ps1
```

This creates a clean client ZIP plus a SHA-256 checksum in `dist/`. The ZIP contains its own `RELEASE_MANIFEST.json` with per-file hashes. To record the exact packages installed on a validated machine, run `scripts/snapshot_environment.ps1`.

## Engineering safeguards

- stable `channel_id` values instead of ambiguous display-name matching;
- explicit RAW/MATH provenance;
- deterministic formulas/statistics;
- strict dependency and numeric validation;
- no hidden smoothing/resampling in plots;
- reporting layers do not replace engineering calculations;
- source/config hashes and run manifests retained for traceability;
- client-readable stage diagnostics and run log;
- safe output cleanup that preserves unrelated files;
- AI is not part of engineering calculation or selection logic.

## Milestone status

- Phase 1 - reverse engineering: complete
- Milestone 01 - data import: complete
- Milestone 02 - channel management: complete
- Milestone 03 - math channels: complete
- Milestone 04 - statistics: complete
- Milestone 05 - plotting: complete
- Milestone 06 - Excel report generation: complete
- Milestone 07 - report layout/template fidelity: complete
- Milestone 08 - end-to-end processing pipeline: complete
- Milestone 09 - simple user UI: complete
- Milestone 10 - deterministic PowerPoint report: complete
- Milestone 11 - hardening/client readiness: complete
- Milestone 12 - final release/package: complete
- Milestone 13A - Sergio reference duty-cycle reverse engineering: complete
- Milestone 13A.1 - reference conflict resolution & formal KPI specification: complete
- Milestone 13B.0 - duty-cycle composition/provenance foundation: complete
- Milestone 13B.1 - deterministic native/synthetic prefix P01-P04: complete
- Milestone 13B.2 - external phase Profile Provider + complete 17,418-row fidelity composition: complete
- Milestone 13B.3 - full duty-cycle integration into the normal reporting pipeline: complete
- Milestone 13B - independent-profile replacement of the Sergio fidelity bridge: future optional refinement
- Milestone 13C - full Sergio report fidelity: not started

## Milestone 13A - Sergio reference duty-cycle reconstruction

The repository now includes the reverse-engineering evidence required before implementing the optional duty-cycle composer. The analysis is intentionally separated from the validated deterministic core. Milestone 13A.1 adds specification clarifications and safe source-report fidelity fixes without implementing the mission composer.

Files are in `docs/milestone_13A/`:

- `duty_cycle_reverse_engineering_report.md`
- `duty_cycle_phase_inventory.csv`
- `cumulative_channel_logic.csv`
- `source_to_report_row_mapping.csv`
- `unresolved_ambiguities.md`
- `proposed_duty_cycle_spec.yaml`

The analysis accounts for all 17,418 Sergio reference samples as a 12-phase mission with six field-work sections, one separate road-travel section, and five 900 s agrochemical-loading/opportunity-charging sections. The future composer must remain optional and must not modify or replace the existing validated deterministic calculations.


## Milestone 13A.1 - reference conflict resolution

`docs/milestone_13A_1/` freezes the reference hierarchy and definitions required before implementation:

- Sergio Excel is the primary full-mission numerical reference;
- Sergio PowerPoint is presentation/scenario-intent evidence where it does not conflict with Excel;
- source-only reports are explicitly labelled as a 74 ha field cycle rather than the complete duty cycle;
- the top RMS pair is Battery Power RMS + Battery Heatflow RMS;
- Sergio's `Max Battery Power` is defined as discharge magnitude `-MIN(Battery Power)`;
- default drive range-extender control is 40–80% SOC, with a 5–80% final-cycle override;
- 100.0 kWh mission scaling is separated from the 100.77 kWh physical pack reference;
- sample count, timestamp span, event duration and integration semantics are distinct;
- unexplained legacy formulas such as `Fuel Energy = 80 * 12` remain compatibility-only.

The missing road and generator-on field profile provenance remains an explicit input dependency for exact 70-channel reconstruction.


## Milestone 13B.2 - external phase Profile Provider

The duty-cycle subsystem can now materialise the complete 17,418-sample Sergio mission **when an explicit validated provider is supplied for the four phases whose dynamics are absent from the source field workbook**.

Native/synthetic composition remains responsible for P01-P04, P07, P09, P11 and P12. The external provider supplies P05, P06, P08 and P10. The Sergio fidelity provider is configured in:

```text
config/duty_cycle_profiles_sergio_reference.yaml
```

The current Sergio provider is deliberately a **reference-fidelity bridge**: it replays only the missing phase rows from the authoritative Sergio Excel workbook, with filename and SHA-256 validation. It is not a guessed generator overlay or a replacement for independent VSM road/range-extender simulations.

The provider interface is separate from the composer so independent phase workbooks can replace the reference provider later. Full row- and phase-level provenance records which samples came from the external provider.

Example developer command:

```powershell
python -m vsm_postprocessing.duty_cycle_cli config/duty_cycle_sergio_reference.yaml `
  --source reference_files/Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx `
  --profile-config config/duty_cycle_profiles_sergio_reference.yaml `
  --profile-workbook reference_files/Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx `
  --materialize-full
```

This produces a complete 17,418-row numerical CSV plus the row-level composition/provenance plan.

## Milestone 13B.3 - full duty-cycle pipeline integration

The normal end-to-end pipeline now accepts an optional `duty_cycle:` block. When present, the raw source workbook is inspected first, the 12-phase mission is composed with explicit profile provenance, and a clean 70-channel numeric CSV is generated as the downstream processing input. The existing deterministic channel-selection, math, statistics, plotting, Excel and PowerPoint engines then process that 17,418-sample mission without special-case calculation logic.

Reference integration config:

```text
config/end_to_end_sergio_duty_cycle.yaml
```

Direct PowerShell command (no `.ps1` execution-policy dependency):

```powershell
python -m vsm_postprocessing.pipeline_cli config/end_to_end_sergio_duty_cycle.yaml
```

The pipeline output includes `02_duty_cycle/duty_cycle_dataset.csv`, row-level provenance, external-profile provenance, a duty-cycle summary, and the normal Excel/PowerPoint deliverables. The original source-cycle pipeline remains unchanged when the `duty_cycle:` block is omitted.

The resulting Excel/PowerPoint are still the compact v1.x engineering-report layouts; reproducing Sergio's full 70-channel/KPI/chart workbook presentation is Milestone 13C.

## Scope boundary

A general-purpose formula builder, free-form plot designer and AI-assisted report recommendations remain outside the deterministic core. AI, if investigated later, must never replace deterministic engineering calculations.
