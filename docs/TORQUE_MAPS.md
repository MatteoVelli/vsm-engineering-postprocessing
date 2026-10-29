# Torque-speed operating maps

The current Electric, Hybrid and Diesel reporting profiles add a **Torque Maps**
Excel worksheet and **Torque–Speed Operating Maps** PowerPoint page immediately
before Simulation Summary. Explicit legacy profiles retain their original
inventories. No manufacturer envelopes, drivetrain conversions or new engineering
math channels are introduced.

## Measurements and availability

| Profile | Left map | Right map |
|---|---|---|
| Electric | `ElectricSystem_EM1..4_Torque` paired with the corresponding `EM1..4_Speed` | Existing `KERS_MotGenUnit_GearboxOutputTorque_1..4` paired with `Wheel_RotationalSpeed_FL/FR/RL/RR` |
| Hybrid | Same motor pairs; the available Caiman reference has active EM3/EM4 | Configurable gearbox-output pairs plus explicit rear `DriveShaft_Torque_RL/RR` measurements, each paired with its corresponding wheel RPM |
| Diesel | `Engine_Torque` paired with `Engine_Speed` | Existing `DriveShaft_Torque_FL/FR/RL/RR` and corresponding wheel RPM |

Legends identify motor shafts, gearbox outputs and driveshafts. Caiman's added
rear driveshaft inputs are used only by the new maps, have `for_report: false`,
and do not replace the existing gearbox sources in wheel-power/torque equations.
No equivalence to gearbox output or tyre-contact torque is asserted. No source
falls back to `Wheel_CounterTorque`.

Newer Electric UoL exports contain the approved gearbox signals. Older Crop Field
and Gradient/Discharge exports lack them and display an unavailable right map.
The available Caiman reference supports two active motor and two rear driveshaft
series; the absent gearbox series are recorded separately as unavailable.

Signed RPM and torque are retained without smoothing, clipping, sorting or
resampling. Every included component retains every sample in the selected
interval. A missing, invalid-unit, non-finite or misaligned pair is unavailable
as a whole; individual samples are never silently filtered. A component whose
RPM and torque are both exactly zero throughout the selected interval is
explicitly inactive. Constant nonzero pairs and moving zero-torque pairs remain
valid. These decisions do not use the legacy `is_active` flag.

Missing optional measurements do not invalidate an otherwise valid report.
Existing mandatory source requirements and numerical validation remain in force.
Negative electric motor mechanical power is consistent with generating operation;
these maps alone do not quantify recovered battery energy. Negative engine torque
is retained as signed combustion-engine operation/overrun.

## Configuration and selections

The existing current-profile YAML files under
`config/report_profiles/sergio_2026_10/` define both plots and the additional page.
Each series can specify its own `x`, component, label and measurement location:

```yaml
- plot_id: torque_map_powertrain
  title: Motor Torque vs Motor Speed
  x: electricsystem_em1_speed
  x_label: Motor Speed [rpm]
  primary_y_label: Torque [Nm]
  render_mode: scatter
  report_group: torque_maps
  series:
  - semantic_name: electricsystem_em1_torque
    x: electricsystem_em1_speed
    component: EM1
    label: EM1 / FL
    measurement_location: electric motor shaft
    required: false
```

The shared plotting engine also supports an optional `x_channel_id` on generic
series and `render_mode: scatter` on generic plot definitions. Omitted rendering
mode retains the existing line behaviour. Simulation maps use markers without
connecting lines; each native Excel series has an independent X/Y reference.

The Streamlit expander **Channels, statistics, plots and optional inputs** adds
two component selectors. Existing Excel plot inclusion and PowerPoint slide-slot
selectors control report inclusion. Excluding sources from the main worksheet
does not remove the new worksheet's paired chart data. Deselecting both map plots
from both destinations omits the feature.

Programmatic selection uses the existing `ProfileReportSelection`, with an
optional `plot_series` mapping from plot ID to selected Y semantic identifiers.
Unknown and duplicate component selections fail explicitly. Source changes remain
explicit YAML mappings; the UI does not guess equivalent torque measurements.

## Excel and PowerPoint output

The Torque Maps worksheet contains two native editable marker-only XY charts
when both maps are included. Unavailable maps retain a labelled empty native
chart and an explicit reason. Component source tables begin below the charts:

- Component and measurement location.
- Actual RPM and torque source names.
- Availability/activity classification and deterministic reason.
- Original CSV/XLSX source row, original sample index and time.
- Exact paired RPM and torque values.

The worksheet has no frozen rows/columns or split panes and scrolls freely in
both directions. Its zoom, styling, chart positions and source tables are retained.

Inactive components retain their zero source tables. Invalid/unavailable pairs
receive source descriptions and reasons without fabricated numeric values.
An unavailable native chart retains marker-free references to blank cells, so
there is no invented operating point. Native Excel can classify that empty
scatter as XY subtype 75; populated operating maps use marker-only subtype -4169.
`profile_plot_manifest.json` records scenario basename, source hash, selected
source row range, per-series X/Y sources, units, component and measurement
location. No external local paths are added to visible PowerPoint content.

PowerPoint uses deterministic Matplotlib images through the existing Astauto
additional plot-pair renderer. The Diesel builder explicitly supports the new
page while retaining its fixed original template pages. Its caption is copied
before Diesel-specific shape replacements can change template shape indices.
All page numbers include the new total. Default complete reports contain
16 Electric, 15 Hybrid and 11 Diesel slides; ordinary optional Diesel-page
omissions still apply.

Existing main worksheet cells/formulas, charts and KPI definitions are preserved.
The overview's rendered-plot count and deck numbering update to reflect the new
content. Reference templates and source datasets are unchanged. The application
version is 1.4.0 for the 29 October 2026 client release.

The subtitle is exactly the same across all three current profiles:
"VSM simulation operating points showing the relationship between rotational speed and torque."

## Validation

`tests/test_torque_maps.py` covers independent signed XY pairing, marker styling,
missing X/Y, incompatible units, non-finite/misaligned pairs, inactive versus
constant data, counter-torque rejection, source policy, component selection,
analysis intervals, main-export deselection, native chart references, caption
geometry, slide ordering/numbering and explicit feature omission.

Existing regression expectations change only where the new default page or map
previews change an intentional count. The Hybrid fallback test avoids duplicate
raw driveshaft headers introduced by its synthetic substitutions; its assertion
that existing wheel-power equations cannot fall back remains intact.

Run the complete suite with a fresh, short temporary directory:

```powershell
.venv/Scripts/python.exe -B -m pytest -p no:cacheprovider --basetemp=outputs/tqnew -q
```

Generate and independently inspect all available raw references:

```powershell
.venv/Scripts/python.exe -B scripts/validate_torque_maps.py
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/validate_torque_maps_office.ps1
```

The Python validator checks source values, each chart reference, markers, slide
geometry/order and reference-file integrity. When the assessment's local baseline
snapshot is present, it also compares raw/math hashes, every statistic/KPI,
existing main worksheet/chart XML and original slide XML. Only page totals and
the rendered-plot count are normalized in slide comparisons.

The Office validator opens disposable short-path copies, exports chart/slide
previews and checks native chart types and text bounds. It tests editable source
references through a temporary edit. Hidden read-only Excel can retain its
display cache; the validator verifies the persisted edit by saving/reopening a
disposable copy when required. Original generated reports are not edited.

Local evidence and report paths are under `outputs/torque_maps_2026_10_28/`:
`validation.json`, `report_paths.txt`, final pytest XML/logs and `office_preview/`.
These generated artifacts are excluded from source control.

## Further work

Optional external manufacturer curves and multiple uploaded scenarios are not
implemented. Future curves need independent RPM arrays, declared units,
component/shaft location, peak/continuous/full-load classification and source
revision/operating conditions. Wheel-side conversion requires validated ratios,
topology and losses. The existing screenshot is not a numerical curve source.
Observed operating maxima must never become fabricated capability envelopes.
