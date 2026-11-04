# Motor Torque and Power — review evidence, 4 November 2026

Implemented in the permanent Electric and Hybrid generator paths. Fresh actual-UI
Electric 03 output has **30 native Excel charts and 17 PowerPoint slides**.
The 29 original chart definitions remain byte-identical. No original export was
overwritten; no commit, push, tag or client ZIP was created.

## Generated deliverables

- [Fresh UI Excel](../outputs/motor_power_2026_11_04/browser_reports/20261104_170710_562617/profile_excel_report/RoboSprayer_8500Kg_Electric_03_UoL_DC_20kph_DriverPC_25pcGrad_FL-RL_Rough.xlsx)
- [Fresh UI PowerPoint](../outputs/motor_power_2026_11_04/browser_reports/20261104_170710_562617/profile_powerpoint_report/RoboSprayer_8500Kg_Electric_03_UoL_DC_20kph_DriverPC_25pcGrad_FL-RL_Rough.pptx)
- [Native PowerPoint review PDF](../outputs/motor_power_2026_11_04/native_review/Electric_03_review.pdf)
- [Native slide 15 preview](../outputs/motor_power_2026_11_04/native_review/Motor_Torque_Power_slide_15.png)
- [Native Excel chart preview](../outputs/motor_power_2026_11_04/native_review/torque_maps_chart_3.png)
- [Native Torque Maps chart-region PDF](../outputs/motor_power_2026_11_04/native_review/Torque_Maps_chart_region.pdf)

These XLSX/PPTX files were produced by the actual Streamlit generation button from
the Electric 03 CSV upload. The browser entered the reference’s 23 peak/continuous
points, ratio 17.5, efficiency 0.97 and recorded battery mass 1,700 kg, then
validated the inputs and generated both outputs. The wrapper only redirects
uploads/output to this review folder; it does not inject backend state or patch
generated reports. The supported installed runtime uses Python 3.11.9 and
Streamlit 1.49.1 with the current workspace source.

## Cause and final behavior

The earlier 16-slide export included operating maps and torque-limit overlays,
but no power columns or four-curve capability chart. The first-pass PPT helper
also depended on a separately configured workbook that already contained that
chart. That dependency has been removed. The pipeline now writes live formulas
from the current run’s editable inputs and reads that same generated workbook
for PowerPoint.

On `Torque Maps`, peak inputs remain AX5:AX27 / AY5:AY27 and continuous inputs
BI5:BI27 / BJ5:BJ27 for this case. New outputs are BP5:BP27 and BQ5:BQ27:

```text
=AY5*AX5*2*PI()/60000
=BJ5*BI5*2*PI()/60000
```

The chart is at AA6, above the original tables and clear of the two operating
maps. Output columns and a fallback chart region are selected in free space.
Native Excel needs a complete secondary XY axis pair; the extra speed axis is
hidden. Both groups use the same two RPM arrays, keeping their automatic bounds
aligned. The visible axes are motor speed [rpm], torque [Nm] on the left and
power [kW] on the right. Inputs stay blue and calculated power is black.

The new PPT slide is 15, `Motor Torque & Power Characteristics`; operating maps
move to 16 and summary remains last at 17. `MOTOR BRAND:` and
`MOTOR TYPE / MODEL:` are native editable text boxes with intentionally blank
values. The plot is an image of the exact workbook arrays. Excel retains its
native editable chart; regenerate PowerPoint after changing inputs.

Missing/invalid curves produce a visible omission warning. Peak-only overlays
and continuous-only analysis still work. There are no default manufacturer
curves, external-workbook fallbacks, or Electric-to-Hybrid data substitution.
Each curve may have its own RPM grid. Final drive and efficiency never enter
the motor-shaft power formula.

## Numerical and preservation acceptance

| Check | Result |
|---|---|
| Points per input curve | 23 |
| RPM | 0–3500, strictly increasing, matching grids for this case |
| Peak / continuous maximum torque | 1068.571055 / 696.988611 Nm |
| Power at zero RPM | 0 / 0 kW |
| Peak maximum power | 55.000000141031 kW |
| Continuous maximum power | 44.510100769599 kW |
| Final seven power points | Within 0.000002 kW of 55 / 44.5101 kW |
| Original native chart XML | All 29 byte-identical; one chart added |
| Original numeric cells and formula caches | Unchanged across all original sheets |
| VSM samples | 2161 unchanged |
| Existing PPT text/KPI values | All 16 slides unchanged apart from consumption titles and page counts |

Spot checks at 175 / 525 / 1750 RPM give peak powers
19.582587361 / 54.999999976 / 55.000000031 kW and continuous powers
12.772983416 / 38.318950249 / 44.510100660 kW. All 46 power points were checked.

The only non-metadata original cell change is the Motor Duty Summary source-path
label, which points to the UI upload copy. SHA-256 verifies that copy is identical
to the original CSV. The existing consumption labels occur once; KPI values
are retained. No engineering channels or telemetry calculations were added.

## Verification

- Full regression suite: **1110 passed, 7 skipped, 0 failures, 0 errors (1117 collected).**
- Final focused implementation tests: **46 passed, 0 failures, 0 errors**.
  These reran after the native secondary-axis correction; they overlap the full suite.
- Native Excel: four series in the correct axis groups, common visible speed
  scale, all XY points equal worksheet cells, and automatic recalculation after
  editing AX6, AY6, BI6 and BJ6. Ratio/efficiency edits leave shaft power unchanged.
  Verification used a disposable copy and did not reopen it between edits.
- Native Office: both exports open successfully, all slide-15 text fits, and the
  new Excel chart covers neither existing charts nor data tables. All three native
  Torque Maps chart renders and the new PPT slide were visually inspected.
- Electric/Hybrid isolation: integration tests use explicitly synthetic, distinct
  ratings and unequal grids. Both add exactly one slide and retain other slides.
  No actual Hybrid manufacturer curve was supplied, so real-data acceptance is
  limited to Electric 03.
- The attachment contains text rather than Sergio’s screenshot image. The
  rendered plateau shape matches the supplied numerical specification; a direct
  image-to-image comparison could not be made.

The ZIP-producing tests are intentionally skipped because client ZIP creation
was prohibited. The other skip is the existing Diesel soil-applicability case.
Synthetic two-point tests intentionally emit the fewer-than-three-checks warning;
the actual 23-point report has **no motor capability validation warnings**.

One existing visible data warning remains: **density is required for mass-to-volume
conversion**. Existing informational provenance about battery energy/SOC and tank
load interpretation remains in the manifest. No metadata warnings are present.

Evidence: [artifact audit](../outputs/motor_power_2026_11_04/artifact_validation.json),
[cell/KPI preservation](../outputs/motor_power_2026_11_04/preservation_validation.json),
[Office checks](../outputs/motor_power_2026_11_04/native_review/office_validation.json),
[browser run](../outputs/motor_power_2026_11_04/browser_validation.json),
[full JUnit](../outputs/motor_power_2026_11_04/full_suite.xml),
[focused JUnit](../outputs/motor_power_2026_11_04/final_focused.xml).

## Source files changed for this request

| File | Purpose |
|---|---|
| [motor_capability_excel.py](../src/vsm_postprocessing/motor_capability_excel.py) | New live power columns, exact input-reference checks, native four-series chart, free placement, and missing/invalid-input status. |
| [excel_report_engine.py](../src/vsm_postprocessing/excel_report_engine.py) | Calls the new writer for Electric/Hybrid; records workbook provenance and a visible omission diagnostic. |
| [profile_powerpoint_report_engine.py](../src/vsm_postprocessing/profile_powerpoint_report_engine.py) | Reads this run’s generated XLSX; renders the new slide, blank editable motor fields, and numbered insertion. |
| [motor_capability.py](../src/vsm_postprocessing/motor_capability.py) | Keeps strict four-series/formula validation; adds the requested slide title and inserts it before operating maps. |
| [motor_performance_ui.py](../src/vsm_postprocessing/motor_performance_ui.py) | Explains the two-curve requirement, strictly increasing RPM, and shaft-power calculation. |
| [report_profile.py](../src/vsm_postprocessing/report_profile.py) | Removes the earlier experimental external-workbook setting; unknown settings are rejected. |

Regression files: `tests/test_motor_capability.py`,
`tests/test_motor_capability_excel.py`, `tests/test_continuous_reporting.py`.
Configuration and request documentation were updated. Validation helpers and
all review outputs are under `outputs/motor_power_2026_11_04/`.
