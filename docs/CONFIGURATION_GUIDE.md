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

Current October Electric/Hybrid profiles append four single-series soil plots
using `x: track_time`, `x_label: Time [s]`, power in kW and cumulative energy in
kWh. They reference the existing MATH channels directly; accumulated energy is
not integrated again. Their orders 100-103 append them after the existing charts.

An extending profile can append styled PowerPoint plot-pair pages while keeping
its inherited presentation:

```yaml
overrides:
  additional_presentation_slides:
  - slide_id: soil_bulldozing
    mode: additional_plot_pair
    title: Tyre-Soil Bulldozing
    subtitle: Resistance power and accumulated energy versus time
    plots:
    - total_bulldozing_resistance_power
    - tyre_bulldozing_resistance_energy_accumulated
```

These additions require an inherited presentation, unique slide IDs, a title,
one or two existing plot IDs and either zero or six statistics. The current
soil pages configure six presentation-only `soil_bulldozing_*` or
`soil_compaction_*` metric IDs; their definitions reuse validated report channels,
the shared signed maximum/sample RMS/time-average functions and `distance_km_last`.
No new Excel statistic or formula is introduced. They use the approved Astauto
background, header, logo, footer and plot-pair dimensions and appear before the
final summary. Unavailable dependencies produce labelled notices. Existing
PowerPoint plot selection applies to the new slots, including deselection.

## Reports

Excel and PowerPoint are presentation layers. They consume outputs from the deterministic calculation layers and do not independently recompute engineering values.

Raw definitions accept `for_report: false` (default: true) to suppress an Excel
report column while preserving ingestion, semantic mapping and calculation
availability. Diesel uses this for `engine_load`; its dedicated statistic and
plot are also removed from the profile. `engine_torque` is unchanged. The UI
omits suppressed channels from its Excel export chooser.

Every profile cover displays `ReportMetadata.source_filename` exactly, including
its extension, in a 13 pt text area with native word splitting. Filename values
are descriptive metadata and are not used as measured KPI values.

The current Electric profile places `tyre_total` before its two soil pages in
`additional_presentation_slides`. Its six `tyre_total_*` cards reuse the same
presentation metric provider as the soil pages. Two time plots (orders 104-105)
reference `total_rolling_resistance_power` and
`tyre_rolling_resistance_energy_accumulated`; they also appear as native Excel
charts. Rolling resistance already includes bulldozing and compaction (Sergio's
clarification). The slide identifies these as rolling resistance
results and preserves its six metric IDs and layout. The redundant four
`tyre_total_resistance_*` MATH channels are removed, including when reading
historical template selections. `tyre_total_energy_accumulated_last` retains its
ID but now targets accumulated rolling energy; the existing tyres-plus-auxiliary
KPI inherits that source. Soil diagnostics remain independent and missing soil
inputs do not disable rolling consumption. Integration conventions are unchanged.

The previous composed mission scenario/provider workflow has been removed. Source files are processed directly through the selected profile or generic pipeline configuration.

## Machine specifications and qualified agro reporting

Existing profile metadata contains reusable `machine_configurations` records for
compatible machine identity and physical source interpretation. Vehicle mass alone
does not identify a battery pack. Electric and Hybrid use the explicit optional
`reporting_inputs.battery_mass_kg` value supplied for the current report. The UI
field is **Battery pack mass (kg)**. It accepts a finite positive value, records it
as `user_provided`, and displays it in the seventh battery-system PowerPoint card.
Blank input displays N/A. No preset pack mass is inferred; Diesel needs no pack.
An explicitly supplied status/provenance remains available through configuration.

The primary agro discharge KPI is `agro_discharge_rate_kg_per_km`: initial minus
final remaining mass [kg] divided by final cumulative distance [km]. It reuses the
existing mass channel, with -force/9.81 for Electric/Hybrid and -force/9.80665 for
Diesel. Existing kg sources are not converted again. It needs no density. Missing,
invalid or nonphysical data produces N/A and a warning. A nonzero distance start
or possible refill is warned about without changing the specified net formula.
Current templates are Electric 08, Hybrid 08 and Diesel 04. Semantic channel
mapping and dynamic sample extents are retained; worksheet names do not determine
powertrain. Legacy profiles remain available with their established templates.

`reporting_inputs` also supports advanced volumetric diagnostics and density
provenance. The `agro_density_kg_per_L` public field and `density_kg_per_L` internal
alias are compatible; supply only one. Density must be positive, finite and
documented. Trusted cumulative chemical volume takes priority over integrated
volumetric flow, then supported remaining mass with density. Fuel channels are
not chemical volume sources. Missing density leaves the separate L/m diagnostic
unavailable. Its existing refill, discontinuity and physical-data safeguards are
unchanged. Agro Capacity is maximum observed inferred remaining mass, not nominal
tank capacity. The accepted legacy Diesel diagnostic gravity difference is retained.
Detailed source, status and warnings are in Statistics and report metadata.
