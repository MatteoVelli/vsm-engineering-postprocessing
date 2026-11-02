# Torque-speed operating maps

The current Electric, Hybrid and Diesel reporting profiles add a **Torque Maps**
Excel worksheet and **Torque–Speed Operating Maps** PowerPoint page immediately
before Simulation Summary. Explicit legacy profiles retain their original
inventories. No manufacturer envelopes, drivetrain conversions or new engineering
math channels are generated. Both charts support manual overlays through
Excel's native Select Data command; no dedicated input worksheet is needed.

## Measurements and availability

| Profile | Left map | Right map |
|---|---|---|
| Electric | `ElectricSystem_EM1..4_Torque` paired with the corresponding `EM1..4_Speed` | Existing `KERS_MotGenUnit_GearboxOutputTorque_1..4` paired with `Wheel_RotationalSpeed_FL/FR/RL/RR` |
| Hybrid | Same motor pairs; the available Caiman reference has active EM3/EM4 | Configurable gearbox-output pairs plus explicit rear `DriveShaft_Torque_RL/RR` measurements, each paired with its corresponding wheel RPM |
| Diesel | `Engine_Torque` paired with `Engine_Speed` | Existing `DriveShaft_Torque_FL/FR/RL/RR` and corresponding wheel RPM |

Electric/Hybrid gearbox series now display `FL Wheel`, `FR Wheel`, `RL Wheel`
and `RR Wheel` in both Excel and PowerPoint. These are display labels only:
`KERS_MotGenUnit_GearboxOutputTorque_1..4`, their corner assignments, units and
numeric arrays are unchanged. Each series retains `measurement_location: motor
gearbox output`; full RPM/torque source names and locations remain in the Torque
Maps source tables and manifests. Separate driveshaft legends retain their
distinct identities. Caiman's added
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
chart and an explicit reason. Component provenance begins below the charts:

- Component and measurement location.
- Actual RPM and torque source names.
- Availability/activity classification and deterministic reason.
- Each component has a local numerical table with source row, sample index,
  time, speed [rpm] and torque [Nm]. Headers are in row 39; samples start at
  row 40, in six-column blocks with a spacer between components.
- Native X/Y references point directly to the corresponding local speed and
  torque columns, independently of main-report channel export selections.
- Local values are deterministic copies of the processed VSM samples. Editing
  a local map cell updates its chart without changing main-report source data,
  mathematical channels, statistics or KPIs.

The worksheet has no frozen rows/columns or split panes and scrolls freely in
both directions. Its zoom, styling, chart positions and source tables are retained.
Both charts have native numeric XY axes (RPM horizontally, torque vertically),
three-point markers and the same deterministic series colours as PowerPoint.
Their minimum, maximum and major interval are automatic in Excel so an added
manufacturer curve can expand the axes without clipping. All signed simulation
samples are retained; other report charts keep their existing axis settings.
The workbook uses automatic calculation and full recalculation on opening.
`forceFullCalc` remains false: enabling it before loading prevents live chart
source updates in the installed Excel. No equations or cached numerical
results change.

## Adding manufacturer curves in Excel

1. Enter the manufacturer's RPM and torque values in any unused cells, such as
   columns A/B below the existing content on **Torque Maps**. Keep the original
   report data and provenance intact. No extra worksheet is required.
2. Right-click the matching chart and choose **Select Data > Add**. Enter a
   series name, select the populated RPM cells for X values, and select the
   matching torque cells for Y values. Existing series can be inspected or
   changed with **Select Data > Edit**. Referenced cell edits update the chart.
3. Format the added curve with straight lines if desired. Keep simulation
   series as operating-point markers. Further curves can use separate XY
   series, each with its own numerical speed array. The axes expand automatically.
4. Use motor-shaft data for the motor/EDU chart. The wheel chart requires values
   at the corresponding wheel-side measurement location in rpm and Nm. Use
   only independently verified gearing and loss assumptions; the application
   performs no automatic conversion. Retain the manufacturer's source, rating
   (peak/continuous), revision and operating conditions with the curve data.
5. Copy the finished chart into PowerPoint using a native chart paste option to
   retain editability. Choose **Keep Source Formatting & Embed Workbook** when
   available; a linked native chart retains its dependency on the Excel file.
   Picture paste retains only the appearance. Copying into a separate validation
   deck does not alter the approved PowerPoint.

Motor-shaft and wheel-side curves remain distinct. Observed ratios from one
simulation are not verified manufacturer capability parameters. No curves,
conversion factors or additional worksheets are generated.

Inactive components retain their zero local source tables. Invalid/unavailable pairs
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
geometry, slide ordering/numbering, complete local source tables regardless of
main export selection, independent native X/Y references, manual XY additions
and explicit feature omission.

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
references through a temporary edit. API writes can retain a display cache;
the validator verifies persisted edits by saving/reopening a disposable copy
when required. With `-ExerciseMotorOverlays`, it also uses focus-checked
keyboard entries in a private visible Excel window to verify live X and Y
point updates and automatic scaling on both charts, then restores the values. Original generated reports are not edited. With `-ExerciseMotorOverlays`, it
also adds curves from ordinary unused cells, verifies automatic scaling and
preservation of the simulation series, and copies both charts into a separate
disposable PowerPoint deck. `scripts/validate_final_torque_refinements.py`
regenerates Excel only, compares the approved workbook and hashes the approved
PowerPoint without regenerating it.

Latest Electric 03 baseline-restoration evidence is under `outputs/localmaps03/`:
`validation.json` records source pairing and unchanged engineering values;
`office_preview/checks.json` records live X/Y edits, overlays and editable chart
pastes. The clean workbook is in `case_0/profile_excel_report/`. Final focused
and complete-suite logs/XML are in
`outputs/local_torque_restore_2026_11_02/`. The actual baseline/current workbook
inventory and protected reference hashes are also in that test directory.
`validate_torque_maps_restoration.py` regenerates Excel from the Electric 03 CSV
with the current workbook's reporting inputs, requires byte-identical baseline
Torque Maps worksheet/chart XML, and requires unchanged main-report worksheet
and unrelated chart XML. Earlier assessment evidence remains
in `outputs/torque_maps_2026_10_28/`. Generated artifacts are excluded from
source control.

## Further work

Automatic manufacturer-curve import/conversion and multiple uploaded scenarios
are not implemented. Manual Excel overlays are supported as described above.
Curves need independent RPM arrays, declared units,
component/shaft location, peak/continuous/full-load classification and source
revision/operating conditions. Wheel-side conversion requires validated ratios,
topology and losses. The existing screenshot is not a numerical curve source.
Observed operating maxima must never become fabricated capability envelopes.
