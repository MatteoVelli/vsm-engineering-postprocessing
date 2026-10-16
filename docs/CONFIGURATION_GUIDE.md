# Configuration Guide

The user interface is the recommended way to configure normal runs. YAML files remain available for engineering traceability, automation and advanced maintenance.

## Reporting Profiles

Production client reports are driven by:

- `config/report_profiles/sergio_2026_10/robosprayer_electric.yaml`
- `config/report_profiles/sergio_2026_10/robosprayer_hybrid.yaml`
- `config/report_profiles/sergio_2026_10/full_size_sprayer_diesel.yaml`

The profiles define raw channel mappings, deterministic math channels, statistics, KPIs, plots, Excel layout, and PowerPoint template selection. Raw channels may be marked `required: false`; missing optional channels are omitted from generated reports without invalidating the source file.

The UI defaults to Sergio's 15 October 2026 revision. The original profiles one
directory above remain available through the explicit Legacy revision. Current
profiles extend those definitions with validated `overrides` maps for `raw`,
`math`, `statistics`, `kpis` and `plots`. Override keys identify existing semantic
names/IDs; null removes a definition. Unknown IDs/fields are rejected. Raw/MATH
names must remain distinct after overrides. `profile.excel_template` identifies
the optional workbook used for column-order verification.

Current Electric/Hybrid profiles use the four exact raw
`KERS_MotGenUnit_GearboxOutputTorque_1..4` channels. Missing new sources disable
dependent wheel calculations, without a driveshaft fallback. Legacy profiles
retain their four `DriveShaft_Torque_*` sources. Current auxiliary calculations
require the measured high-voltage source. All source names, units and optional
dependencies are specified in YAML. See
[the current implementation report](SERGIO_2026_10_15_IMPLEMENTATION.md) for
integration, formula, source-availability and selection conventions.

### Confirmed source-unit interpretations

An optional `unit_interpretation` on a raw profile channel declares an explicitly
confirmed engineering unit without converting its numeric values. The current
Electric and Hybrid profiles contain this declaration only for
`User_defined_Data_High_Voltage_Auxiliary_signal1_0`:

```yaml
unit: kW
unit_interpretation:
  source_unit: '-'
  engineering_unit: kW
  authority: 'Sergio, Astauto, 16 October 2026: confirmed power in kW; factor 1, offset 0.'
```

It applies only when both the actual source name and source-reported unit match
the declaration exactly. Aliases, normalized name matches, unspecified units and
other units do not activate it. An explicit source `kW` is accepted normally.
The engineering unit must match the channel's expected `unit`; authority is
mandatory and conversion/scaling fields are rejected. There is no global `-`
unit rule. Missing or invalid sources remain unavailable as before.

Imported `ChannelInfo.unit` and raw numeric values remain unchanged. Resolved
channels expose `effective_unit` and a separate engineering metadata view for
statistics and plots. Applied declarations appear in source diagnostics, the
Excel mapping-sheet notes, workbook metadata, and the report manifest's
`source_unit_interpretations` records. Each record preserves the reported unit,
effective unit, exact source identity, authority and unchanged-value status.
The existing signed, timestamp-based trapezoidal auxiliary-energy integration
is unchanged. See [HV engineering validation](HV_AUXILIARY_CONFIRMATION_2026_10_16.md).

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

Optional channels cover height, fuel flow/specific consumption, controls,
oil/coolant temperature, accelerations, tyre forces/powers, wheel speeds/torques
and mechanical auxiliaries. In current profiles, invalid optional sources are
reported as unavailable and dependent sections omitted. Required data must be
finite, time strictly increasing with at least two samples, and fuel mass/distance
counters nonnegative and nondecreasing. Counter offsets are allowed.

The current Diesel profile has 48 raw definitions, 20 MATH definitions, 55
statistics, five KPIs and 20 plots. These include signed engine/wheel/auxiliary
power, net mechanical energy, integrated fuel volume, agrochemical mass and tyre
losses. Integration starts at zero and uses actual elapsed time. No fuel density
is assumed. MATH channels are calculated, never requested as raw source inputs.

Counts reduce with unavailable inputs or user selections. The ten configured
Diesel PowerPoint pages retain approved layouts; optional pages can be omitted
when unavailable. The separate water-temperature plot and its source mapping
are selectable. No battery/generator metric is inferred for Diesel.

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
