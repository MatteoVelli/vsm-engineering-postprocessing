# Configuration Guide

The user interface is the recommended way to configure normal runs. YAML files remain available for engineering traceability, automation and advanced maintenance.

## Reporting Profiles

Production client reports are driven by:

- `config/report_profiles/robosprayer_electric.yaml`
- `config/report_profiles/robosprayer_hybrid.yaml`

The profiles define raw channel mappings, deterministic math channels, statistics, KPIs, plots, Excel layout, and PowerPoint template selection. Raw channels may be marked `required: false`; missing optional channels are omitted from generated reports without invalidating the source file.

The current Electric and Hybrid profiles are synced to Sergio's Electric_03 and Hybrid_04 templates and include optional `Track_Height` mapped to visible report label `Road Height`.

## Configuration Separation

The generic pipeline keeps the main engineering choices independent:

- `channel_selection_example.yaml` - source channels to export;
- `math_channels_example.yaml` - deterministic calculated channels;
- `statistics_example.yaml` - general RMS/MAX/MIN/LAST/SUM calculations;
- `statistics_excel_report.yaml` - statistics placed in the Excel report;
- `plotting_example.yaml` - plots generated from raw/math channels;
- `excel_report_example.yaml` - Excel channel/layout choices;
- `powerpoint_report_example.yaml` - optional PowerPoint slide content;
- `end_to_end_example.yaml` - orchestration of the complete workflow.

## Channel Identity

Always use the stable `channel_id` emitted by source inspection or the semantic name defined in a reporting profile. Display names are not guaranteed to be unique.

## Math Channels

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

The previous composed mission scenario/provider workflow has been removed. Source files are processed directly through the selected profile or generic pipeline configuration.
