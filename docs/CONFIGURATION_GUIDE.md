# Configuration Guide

The user interface is the recommended way to configure normal runs. YAML files remain available for engineering traceability, automation and advanced maintenance.

## Configuration separation

The application deliberately keeps the main engineering choices independent:

- `channel_selection_example.yaml` - source channels to export;
- `math_channels_example.yaml` - deterministic calculated channels;
- `statistics_example.yaml` - general RMS/MAX/MIN/LAST/SUM calculations;
- `statistics_excel_report.yaml` - statistics placed in the Excel report;
- `plotting_example.yaml` - plots generated from raw/math channels;
- `excel_report_example.yaml` - Excel channel/layout choices;
- `powerpoint_report_example.yaml` - optional PowerPoint slide content;
- `end_to_end_example.yaml` - orchestration of the complete workflow.

## Channel identity

Always use the stable `channel_id` emitted by the source inspection catalogue. Display names are not guaranteed to be unique. This is essential for datasets such as the Sergio reference workbook, which contains duplicate display names with different units.

## Math channels

Math-channel definitions are deterministic. Each definition declares an ID, display name, unit, expression and dependencies. Missing dependencies, dependency cycles, non-finite outputs and invalid formula operations are rejected.

## Statistics

Supported deterministic operations are:

- RMS;
- time-weighted RMS;
- MAX;
- MIN;
- LAST;
- SUM.

NaN handling is explicit in configuration. No AI is used to calculate engineering statistics.

## Plots

Plot definitions specify the X channel and one or more Y series. Raw and calculated channels can be mixed. Plot generation performs no hidden smoothing, interpolation or resampling.

## Reports

Excel and PowerPoint are presentation layers. They consume outputs from the deterministic calculation layers and do not independently recompute engineering values.

## Runtime UI profiles

The Streamlit UI creates isolated runtime configurations for each run under `outputs/ui_runs/<timestamp>/`. Retain the complete run directory when traceability is required.

## Duty-cycle external profile provider (v1.2.4)

The duty-cycle composer keeps missing numerical profiles separate from the scenario definition.

Scenario configuration:

```text
config/duty_cycle_sergio_reference.yaml
```

Sergio reference-fidelity provider configuration:

```text
config/duty_cycle_profiles_sergio_reference.yaml
```

The provider configuration defines:

- provider ID/type;
- supported phase IDs;
- explicit source row ranges per phase;
- workbook import bounds;
- channel-alignment policy;
- value policy;
- expected filename and SHA-256.

For Sergio's reference report, column headers differ from the supplied source workbook. The configured 70-column positional alignment is therefore accepted only together with the exact reference workbook SHA-256.

Example full materialisation:

```powershell
python -m vsm_postprocessing.duty_cycle_cli config/duty_cycle_sergio_reference.yaml `
  --source reference_files/Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx `
  --profile-config config/duty_cycle_profiles_sergio_reference.yaml `
  --profile-workbook reference_files/Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx `
  --materialize-full `
  --full-output outputs/duty_cycle_full_composition.csv
```

Without an external provider, `compose_supported_prefix(...)` remains available and intentionally stops before unresolved P05.


## Full duty-cycle end-to-end pipeline (v1.2.5)

To run the normal reporting pipeline on a composed mission, add an optional root-level `duty_cycle` block:

```yaml
duty_cycle:
  scenario: duty_cycle_sergio_reference.yaml
  profile_provider: duty_cycle_profiles_sergio_reference.yaml
  profile_workbook: ../reference_files/Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx
```

The `input.file` remains the raw/source VSM workbook. The pipeline first validates that source, then composes the mission and writes a clean numeric `duty_cycle_dataset.csv` with the same 70 stable channel IDs. Every downstream deterministic processing stage uses that composed dataset.

Reference configuration:

```text
config/end_to_end_sergio_duty_cycle.yaml
```

Run directly from an activated virtual environment:

```powershell
python -m vsm_postprocessing.pipeline_cli config/end_to_end_sergio_duty_cycle.yaml
```

If the `duty_cycle` block is omitted, the original source-file pipeline behavior is unchanged.
