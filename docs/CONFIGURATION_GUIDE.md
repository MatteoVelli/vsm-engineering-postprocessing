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
