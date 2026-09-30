# Configuration Guide

The user interface is the recommended way to configure normal runs. YAML files remain available for engineering traceability, automation and advanced maintenance.

## Reporting Profiles

Production client reports are driven by:

- `config/report_profiles/robosprayer_electric.yaml`
- `config/report_profiles/robosprayer_hybrid.yaml`
- `config/report_profiles/full_size_sprayer_diesel.yaml` (UI label: `Diesel`)

The profiles define raw channel mappings, deterministic math channels, statistics, KPIs, plots, Excel layout, and PowerPoint template selection. Raw channels may be marked `required: false`; missing optional channels are omitted from generated reports without invalidating the source file.

The current Electric and Hybrid profiles are synced to Sergio's Electric_05 and Hybrid_06 templates and include optional `Track_Height` mapped to visible report label `Road Height`.

Both profiles require all four raw `DriveShaft_Torque_*` channels. Wheel torque
is never synthesized. Tyre Total Energy Consumption includes rolling,
compaction, and bulldozing resistance, integrated using actual sample timing.

An extending profile can set `channels.plot_overrides` to a mapping of existing
semantic channel names to boolean selection flags. Hybrid uses this to retain
the four wheel-load flags while inheriting Electric's common calculations.
Unknown names and non-boolean values are rejected. These flags describe the
channel-selection sheet; explicit report charts remain defined under `plots`.

## Configuration Separation

### Full Size Sprayer Diesel

Diesel is an independent profile, with no Electric/Hybrid inheritance. Required
VSM inputs are `Track_Time [s]`, `Track_Distance [m]`, `Chassis_Speed [kph]`,
`Track_Gradient [%]`, `Engine_Speed [rpm]`, `Engine_Torque [Nm]`, and
`Engine_FuelConsumption_absolut [kg]`. Source load torque is Nm, not percent.

The 18 optional channels cover height, volumetric fuel flow, source specific
fuel consumption, load/throttle, oil/coolant temperature, three accelerations,
four tyre vertical loads and four driveshaft torques. Absent optional channels
omit dependent results; present invalid channels fail validation. Source data
must be finite, time strictly increasing with at least two samples, and fuel
mass/distance counters nonnegative and nondecreasing. Counter offsets are allowed.

Three MATH channels are calculated internally: signed engine power [kW], net
engine mechanical energy [kWh], and optional integrated fuel volume [L]. Energy
and volume use trapezoidal integration over elapsed seconds and start at zero.
No fuel density is assumed. MATH channels are never source requirements.

The profile configures 52 statistics, five KPIs, ten plots and ten PowerPoint
pages. Counts reduce when optional data is absent. The representative Test_03
CSV yields 50 statistics (height is absent), five KPIs and all ten plots/pages.
With only required channels, Excel has four native charts and PowerPoint seven
pages. All-zero plot traces may be omitted from PowerPoint while remaining in
Excel. No battery/generator, wheel-power, auxiliary-energy or tyre-loss KPI is
inferred for Diesel.

Use Engineering Report, upload CSV/XLSX, select Diesel, validate, set the machine
name and generate both reports. Changed input/profile or a new validation attempt
clears validation/download state. A changed machine name hides old downloads.
Output filenames preserve the source stem; machine overrides change report titles.
Both downloads and plot previews are available. Streamlit 1.49.1 is the minimum
supported client runtime for the existing `width` API calls.

See [Diesel final validation](DIESEL_FINAL_VALIDATION.md) for formulas, reference
provenance, independent numerical checks and the release assessment.

## Generic Pipeline Configuration

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
