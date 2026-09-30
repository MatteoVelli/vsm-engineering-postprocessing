# Diesel Phase D2: native Excel integration

> Historical stage record. The Excel-only boundary and deferred PowerPoint work
> below describe D2 at completion. The current application supports combined
> Diesel Excel and PowerPoint generation. See the superseding
> [Diesel final validation](DIESEL_FINAL_VALIDATION.md) for current status.

## Starting state and protected work

Branch `main`, HEAD `0ec8529705a18bd88bd7524bdb816dc9a42efef1`.
The completed D1 suite had 298 passing tests. Before D2 edits, the focused
baseline (D1, profile Excel and native-chart readability) passed **72 tests in
185.29 s**, with no skips or failures.

Initial status:

```text
 M docs/SERGIO_REPORT_CHANGES.md
 M reference_files/README.md
 M src/vsm_postprocessing/excel_report_engine.py
 M src/vsm_postprocessing/math_engine.py
 M src/vsm_postprocessing/profile_math.py
 M src/vsm_postprocessing/profile_powerpoint_report_engine.py
 M src/vsm_postprocessing/ui_app.py
 M tests/conftest.py
 M tests/test_profile_powerpoint_report.py
?? Matteo_Velli_Manas_Studio_Invoice.pdf
?? config/report_profiles/full_size_sprayer_diesel.yaml
?? docs/DIESEL_D1_VALIDATION.md
?? logs/
?? src/vsm_postprocessing/diesel_validation.py
?? tests/test_diesel_profile.py
```

Initial tracked `git diff --stat`: nine files, 344 insertions, 156 deletions.
The three edits that predate D1 (Sergio notes, PowerPoint engine and PowerPoint
tests) are untouched in D2. A hash snapshot of every initially tracked file and
the new D1 files is stored in `outputs/diesel_d2_start_hashes.json`. The D1 math
engine, validator, profile math/statistics engines, both existing Electric/Hybrid
YAML profiles, and reference CSV remain unchanged. Existing Electric/Hybrid test
assertions were not modified.

## Architecture and isolation

The production Excel adapter imports data, resolves the YAML profile, calculates
profile math/statistics once, and builds semantic channel maps. Electric/Hybrid
use their existing workbook templates to determine channel order. Profiles
without templates use their declared raw channels followed by available math
channels. Diesel uses this existing path: it exports 24 available raw channels
and three math channels, rather than all 586 imported channels.

The generator already writes native editable scatter charts from profile plot
definitions. Diesel adds ten definitions to its YAML and reuses the same chart
builder, engineering axis scaling, colors, fonts, legends and two-column layout.
PNG preview generation is an existing intermediate step; **no PNGs or raster
charts are embedded in the workbook**.

Report-only YAML changes are `for_plot` flags, executive placement groups, plot
definitions and description. A normalized before/after comparison verified that
all raw requirements/units, math expressions/dependencies, statistic operations,
NaN policies and KPI expressions remain identical to D1.

Diesel-only writer branches select fourteen executive metrics, add a visible
Statistics sheet, expose RMS alongside bottom MAX/MIN/LAST/FIRST rows, document
omitted optional results, and replace misleading Electric-template comparison
text with Diesel-specific provenance. Time FIRST/LAST values remain fully
available on Statistics; the first column of bottom rows retains operation labels.

The Excel guard is removed. Streamlit has an isolated Excel-only branch using
the existing Excel wrapper; its download state is keyed to the current source,
profile and machine name. Existing Electric/Hybrid UI paths are unchanged.
The combined Excel/PPT wrapper explicitly rejects Diesel before generation.
The PowerPoint builder's existing missing-presentation guard also rejects Diesel.
No PowerPoint engine code, slide layout or presentation configuration was changed.

## D2 files

| File | D2-specific purpose |
|---|---|
| `config/report_profiles/full_size_sprayer_diesel.yaml` | Ten chart definitions, plot flags and fourteen executive placements; numerical definitions unchanged |
| `src/vsm_postprocessing/excel_report_engine.py` | Enable Diesel Excel; compact executive area, Statistics sheet, bottom RMS, optional-result metadata and Diesel-only chart-title placement |
| `src/vsm_postprocessing/ui_app.py` | Excel-only Diesel generation/download flow; preserve legacy combined workflow |
| `src/vsm_postprocessing/ui_config.py` | Reject Diesel combined/PPT requests while leaving Excel-only wrapper available |
| `tests/test_diesel_profile.py` | Update three D1 phase-boundary expectations for D2; all numerical assertions retained |
| `tests/test_diesel_excel_report.py` | 23 new tests for workbook numbers, selections, native charts, optional channels, formatting and UI/PPT boundaries |
| `docs/DIESEL_D2_VALIDATION.md` | This validation record |

All other modified/untracked items in the starting status belong to D1 or
pre-existing user work and were not changed by D2.

## Workbook and selected channels

Delivered workbook:

`outputs/diesel_d2_report/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.xlsx`

Worksheets, in workbook order:

1. `Full Size Sprayer Diesel` — visible data, executive results, bottom statistics,
   and ten native charts.
2. `Rename From VSM to Astauto` — visible source/report names, units, channel type,
   plot flags, semantic IDs and math expression/dependency provenance.
3. `Metadata` — hidden source/profile hashes, processing counts, visible-sheet
   inventory and unavailable optional channels/statistics/KPIs.
4. `Statistics` — visible 50 statistics and five calculated KPIs, with units,
   operations, targets, placement groups and sample counts.

Raw selection (units are unchanged from D1):

- `Track_Time [s]`, `Track_Distance [m]`, `Chassis_Speed [kph]`, `Track_Gradient [%]`.
- `Engine_Speed [rpm]`, `Engine_Torque [Nm]`, `Engine_Load [Nm]`, `Engine_Throttle [%]`.
- `Engine_FuelConsumption_absolut [kg]`, `_volumeflow [l/h]`, `_specific [g/kWh]`.
- `Engine_Oil_Temperature` and `Engine_Water_Temperature [°C]`.
- `Chassis_Acceleration_Longitudinal`, `_Lateral`, `_Vertical [m/s²]`.
- `Tyre_Fz_FL`, `_FR`, `_RL`, `_RR [N]`.
- `DriveShaft_Torque_FL`, `_FR`, `_RL`, `_RR [Nm]`.

Optional `Track_Height [m]` is exported only when actually supplied; it is absent
from the reference workbook. Source-specific consumption and both thermal signals
are retained as raw data, without introducing new efficiency or thermal KPIs.

Math selection: signed `engine_power [kW]`, net `engine_mechanical_energy [kWh]`,
and optional integrated `fuel_volume [l]`. All 223 samples of every exported
channel are verified against the D1 processing results, including the full signed
power series. Formulas are not reimplemented in Excel.

The executive area contains duration, distance, maximum speed, gradient extrema,
maximum/average RPM, maximum torque, maximum/RMS power, net energy, consumed fuel
mass and average/maximum fuel flow. The complete 50 available statistics and five
KPIs remain exposed on Statistics, including integrated fuel volume and all
chassis/driveline results. The two absent height statistics are omitted, not zeroed.

## Numerical workbook comparison

The saved workbook was independently reopened with openpyxl. Its Statistics
values were also compared against the **saved D1 numerical audit**, not merely a
second invocation of the new report code. All 55 available canonical values match
within Excel serialization precision (relative tolerance 2e-15, absolute 1e-14).
There are no engineering discrepancies.

| Metric | D1 value | Workbook stored value, rounded here |
|---|---:|---:|
| Duration | 222 s / 3.7 min | 222 s / 3.7 min |
| Distance | 0.999242 km | 0.999242 km |
| Maximum speed | 22.4686 kph | 22.4686 kph |
| Minimum/maximum gradient | −15% / +15% | −15% / +15% |
| Maximum RPM | 1996.81 rpm | 1996.81 rpm |
| Average RPM | 1537.214164414414 | 1537.214164414414 |
| Maximum torque | 826.219 Nm | 826.219 Nm |
| Maximum power | 165.938626339644 kW | 165.938626339644 kW |
| Minimum power | −8.230092194144 kW | −8.230092194144 kW |
| Power RMS | 74.570694135737 kW | 74.570694135737 kW |
| Net mechanical energy | 3.764781781161 kWh | 3.764781781161 kWh |
| Fuel consumed | 0.772836 kg | 0.772836 kg |
| Integrated fuel volume | 0.942120719458 L | 0.942120719458 L |
| Average fuel flow | 15.277633288514 L/h | 15.277633288514 L/h |
| Maximum fuel flow | 41.4741 L/h | 41.4741 L/h |

The existing three-decimal summary format is retained, so for example fuel mass
is displayed as 0.773 kg while the cell stores 0.772836 kg. The rendered Excel
installation uses decimal commas according to its locale; stored values/units
are unchanged. The audit JSON records full precision and serialization differences.

## Chart inventory and optional behavior

All charts use native editable Excel scatter/line series linked to worksheet
cells, with `Track_Time [s]` as X over rows 5–227. No invented time conversion,
height series, density conversion or fuel-economy calculation was introduced.

| Chart | Series | Y axes | Missing optional behavior |
|---|---|---|---|
| Vehicle Speed & Road Gradient | `chassis_speed`, `track_gradient` | Left kph, right % | Both mandatory; gradient works without height |
| Engine Speed | `engine_speed` | rpm | Mandatory |
| Engine Torque & Signed Power | `engine_torque`, `engine_power` | Left Nm, right kW | Both mandatory; negative power retained |
| Engine Load Torque | `engine_load` | Nm | Omit chart if absent |
| Engine Throttle | `engine_throttle` | % | Omit chart if absent |
| Fuel Flow | `fuel_flow` | L/h | Omit chart, integrated volume and dependent optional results if absent |
| Cumulative Fuel Mass | `engine_fuel_consumption` | kg | Mandatory; preserves raw cumulative counter |
| Driveshaft Torque | Four `driveshaft_torque_*` | Nm | Omit individual missing corners; omit chart if none available |
| Tyre Vertical Loads | Four `tyre_fz_*` | N | Omit individual missing corners; omit chart if none available |
| Engine Oil Temperature | `engine_oil_temperature` | °C | Omit chart if absent; coolant stays exported separately |

Oil alone is plotted because oil/coolant traces are identical in the authoritative
reference. Separate load and throttle plots prevent misleading empty axes when
only one optional signal exists. No Electric/Battery/EDU chart or KPI appears.
If every optional channel is absent, generation still succeeds with nine exported
channels (seven raw, two math), four native charts and complete required results.
Missing charts compact naturally into the existing chart grid without blank slots.

Axis limits come from the established engineering scaling function and encompass
every referenced sample. The power chart's kW axis spans −50 to 200 kW, explicitly
showing the −8.230092 kW overrun region. Required X and Y tick labels are visible;
secondary Y axes are right-aligned, and only redundant secondary X axes are hidden.
Multi-series legends sit below charts; single-series charts omit redundant legends.

## Rendered and structural inspection

Excel automation failed inside the sandbox with `RPC_E_SERVERFAULT`, including
on a minimal one-cell workbook. The authorized read-only automation run outside
the sandbox succeeded. Excel stayed hidden, macros/events were disabled, and no
in-memory review changes were saved back to the delivered workbook.

Actual Excel exports were inspected for all ten charts and the executive summary.
Initial automatic Y-axis title placement overlapped tick numbers. A Diesel-only
manual layout now reserves title margins without changing chart size, font sizes,
palette, line treatment or any legacy chart. Longer axis titles receive additional
space. The visible Statistics sheet has wider ID/description columns and wrapped
text; a rendered sample was also inspected.

Structural checks cover:

- Established navy/white raw headings, distinct orange MATH headings and original
  numeric formatting; fourteen executive values with readable wrapped headings.
- Header merges `AC1:AQ1` and `AC2:AQ2`; summary values on rows 3–4.
- Raw/math data from row 5 through 227; bottom MAX/MIN/LAST/FIRST/RMS on 228–232.
- Ten charts at AC/AK columns, rows 6, 26, 46, 66 and 86. Dimensions remain
  637 × 360 pixels at 96 DPI (6.6354167 × 3.75 inches).
- Two-column spacing and row strides exceed chart extents: no chart overlap.
- No embedded images, external workbook links, cell errors or `#REF!` entries.
- Chart references resolve to the intended semantic columns and exact sample rows.
- Existing freeze panes, mapping conventions and hidden Metadata sheet retained.

Review artifacts are under `outputs/diesel_d2_report/`: ten `excel_chart_*.png`
exports, `excel_summary.png`, `excel_statistics.png`,
`diesel_excel_chart_review.pdf`, and `numerical_workbook_audit.json`.
The PNG/PDF files are review evidence, not substitutes for workbook-native charts.
The PDF print area is an inspection-only setting, not a saved workbook layout.
Rendered inspection does not claim universal visual perfection across Excel
versions, locales, display scaling or other spreadsheet applications.

## Tests and completion

| Run | Passed | Skipped | Failed | Runtime |
|---|---:|---:|---:|---:|
| Starting D2 baseline | 72 | 0 | 0 | 185.29 s |
| Existing Excel/report/UI regression subset | 75 | 0 | 0 | 198.93 s |
| Targeted Diesel D1 + Excel | 59 (36 + 23) | 0 | 0 | 30.06 s |
| Final Diesel D1 + Excel + legacy chart readability | 85 (36 + 23 + 26) | 0 | 0 | 26.56 s |
| Targeted tests with private references absent | 41 | 18 | 0 | 26.22 s |
| Final complete regression suite | **321** | **0** | **0** | **477.21 s** |

Final command: `.venv\Scripts\python.exe -m pytest -q -x -p no:cacheprovider -o addopts='' --basetemp=outputs/d2all`.
Full output: `outputs/diesel_d2_full_final.log`. All tests completed successfully.
An earlier in-progress
full run was stopped to include the last rendered-review layout adjustments.
The subsequent full run caught two legacy chart-helper compatibility failures
(319 passed, two failed): the Diesel layout branch had introduced a metadata
requirement for lightweight plot-only callers. The implementation was corrected
to accept an explicit, default-off `diesel_layout` argument from the report
writer. Existing helper callers retain their original contract and output;
neither failing test nor its assertions was changed.
Another rerun was stopped after an unrelated bootstrap fixture failed during
`venv` creation. The long `--basetemp` path pushed a nested setuptools file beyond
this Windows environment's supported path length; invoking `ensurepip` in that
fixture reproduced the exact failing file path. The same bootstrap test passed
with shorter temporary paths (18.23 s and 18.02 s). The final full run uses
`--basetemp=outputs/d2all`, with no bootstrap code or assertion changes.
No numerical test expectations were changed. Three D1 boundary tests now assert
ten configured plots, an available Excel-only UI action, and continued PPT rejection.

Acceptance assessment:

| Workflow | Result |
|---|---|
| Electric | PASS; numerical definitions and existing regression assertions unchanged |
| RoboSprayer Hybrid | PASS; numerical definitions and existing regression assertions unchanged |
| Caiman Hybrid | PASS; existing behavior and regression assertions unchanged |
| Diesel D1 | PASS; normalized numerical definitions and all numerical assertions unchanged |
| Diesel Excel | PASS; 23 new tests, independent saved-D1 comparison and actual Excel rendering |

**D2 is complete and ready for D3 planning/implementation when requested.**
Work stops here; D3 has not been implemented.

## Deferred scope

Diesel PowerPoint/D3 and final client packaging are deliberately not implemented.
No final portable/client ZIP was rebuilt. Existing `dist/` client artifacts retain
their August 7 modification dates. Existing release-packaging regression tests
may build their isolated temporary fixtures as part of the full test suite.

D1 limitations remain: optional missing height, finite strictly increasing time,
no counter resets, sampled numerical integration, and no new density, BSFC,
fuel-per-distance or fuel-per-hectare assumptions. Coolant is exported but has no
duplicate chart; optional height is exported/statistically summarized when supplied,
but D2's road chart remains speed/gradient.

For D3, use the verified workbook/profile metrics and chart semantics to define
Diesel slide content and a dedicated presentation configuration, then add guarded
PowerPoint regressions. Do not enable the combined report workflow until that
presentation path is validated. No commits, staging or pushes were performed.

## Git audit

`git diff --check` passes. The final tracked stat includes pre-existing edits and
D1 work, as well as D2; it does not include untracked YAML/tests/documents:

```text
docs/SERGIO_REPORT_CHANGES.md                      |  27 +++
reference_files/README.md                          |   6 +
src/vsm_postprocessing/excel_report_engine.py      |  68 +++++-
src/vsm_postprocessing/math_engine.py              |  26 +++
src/vsm_postprocessing/profile_math.py             |   5 +
.../profile_powerpoint_report_engine.py            | 253 +++++++++------------
src/vsm_postprocessing/ui_app.py                   |  42 ++++
src/vsm_postprocessing/ui_config.py                |   4 +-
tests/conftest.py                                  |   5 +
tests/test_profile_powerpoint_report.py            | 172 ++++++++++++--
10 files changed, 449 insertions(+), 159 deletions(-)
```

Final status:

```text
 M docs/SERGIO_REPORT_CHANGES.md
 M reference_files/README.md
 M src/vsm_postprocessing/excel_report_engine.py
 M src/vsm_postprocessing/math_engine.py
 M src/vsm_postprocessing/profile_math.py
 M src/vsm_postprocessing/profile_powerpoint_report_engine.py
 M src/vsm_postprocessing/ui_app.py
 M src/vsm_postprocessing/ui_config.py
 M tests/conftest.py
 M tests/test_profile_powerpoint_report.py
?? Matteo_Velli_Manas_Studio_Invoice.pdf
?? config/report_profiles/full_size_sprayer_diesel.yaml
?? docs/DIESEL_D1_VALIDATION.md
?? docs/DIESEL_D2_VALIDATION.md
?? logs/
?? src/vsm_postprocessing/diesel_validation.py
?? tests/test_diesel_excel_report.py
?? tests/test_diesel_profile.py
```

Compared with the D2 starting snapshot, only the five existing files named in the
D2 file inventory changed; two new D2 files were added. All other starting files
retained their hashes. The delivered workbook SHA-256 is
`ddc9b8b0bbe35436bde9b210f623fc6c202ce2f1e50753c41650fe77c6322fee`.
