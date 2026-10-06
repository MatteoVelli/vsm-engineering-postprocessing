# Final Diesel template extension

`reference_files/Diesel Template.xlsx` is the authoritative Diesel Excel content
and plot reference. Its SHA-256 is
`cbbd42e17cc4a0df6238e67b46c222f86283608b5c124531295396a5dfa08576`.
The workbook is unchanged, and no reference assets were copied or renamed.

The existing implementation is `config/report_profiles/full_size_sprayer_diesel.yaml`.
It uses `profile_math.py`, `profile_statistics.py`, `profile_plotting.py` and
`excel_report_engine.py`; `diesel_powerpoint_report.py` supplies the existing
PowerPoint content through `profile_powerpoint_report_engine.py`. Electric and
Hybrid use the same shared infrastructure and their existing YAML profiles.
Their active Excel references remain Electric_05 and Hybrid_06, and PowerPoint
continues to use the established Astauto layout infrastructure.

Existing Diesel reports were found in `outputs/diesel_d3_final`,
`outputs/diesel_final_audit` and `outputs/diesel_rms_cleanup`. Existing Electric
and Hybrid reports were found in `outputs/final_reporting_sync`,
`outputs/electric_rr_power_validation` and `outputs/hybrid_rr_power_validation`.
Historical generated files were left intact. New review outputs are generated
in `outputs/diesel_template_extension` and `outputs/common_presentation_extension`.

## Additive coverage

All 56 nonempty selection rows in the final workbook map to distinct semantic
profile channels. [Channel coverage](DIESEL_TEMPLATE_CHANNEL_COVERAGE.csv) records
the original labels, source names, units, semantic IDs and corrections.
[Plot coverage](DIESEL_TEMPLATE_PLOT_COVERAGE.csv) maps all eight workbook charts
to generated figures, including their time or distance axes.

The original 25 raw channel definitions, three math channels, 48 statistic
definitions (46 available on the original CSV), five KPIs and ten plots remain.
The extension adds 23 raw definitions,
17 math channels, five supporting statistics and eight plots. It covers:

- Reference positions, pedals, brake pressures and steering wheel angle.
- Agrochemical force and its converted mass signal.
- Four rolling resistance force and power channels, total power, per-interval
  energy in kWh and Wh, and accumulated energy.
- Four wheel speeds, total shaft torque, corner wheel powers and total wheel power.
- Low-voltage auxiliary demand, two distinct auxiliary torque sources, their
  mechanical powers and total mechanical auxiliary power.
- Time in minutes, distance in kilometres and template per-interval engine energy.

The supplied Diesel CSV exports 47 raw and 20 math channels (67 total), computes
51 statistics and the original five KPIs, and renders all 18 native Excel plots.
Road height remains absent from that CSV. Newly selected inputs are optional:
older source files remain usable, and dependent missing results are recorded
in the plot catalogue and workbook metadata. Wheel and auxiliary comparisons
require their complete series set; absent optional inputs omit that comparison.
Existing partial-corner behaviour of the baseline shaft/load plots is preserved.

The original Diesel PowerPoint content and slide selection remain; its plots
receive the shared presentation changes. The eight additional template figures
extend Excel coverage and are also generated as reusable Matplotlib assets.

## Reference corrections and preserved definitions

The template repeats the FR label for the RR rolling resistance power channel.
The extension uses the actual RR source and labels it RR. `Generator Torque_2`
incorrectly repeats `Engine_AuxiliaryTorque_1` in the template selection table;
the second auxiliary uses `Engine_AuxiliaryTorque_2`. The agrochemical report
label resolves to the already established Electric/Hybrid hitch-force source
`HitchRear_Force_Z_VehicleCoordinates`, with the Diesel workbook's gravity
constant of 9.80665. The converted trace retains its sign and is not interpreted
as consumed chemical mass.

Chart 5 contains a distance extent ending at row 117876, whereas its values end
at 17876. All generated charts use the current aligned source extent instead
of those fixed row numbers. The template's initial rolling resistance energy
accumulation includes a forward reference and then omits the first energy
sample; the new accumulated channel includes each generated interval once.

The baseline Diesel engine power retains its exact SI conversion and negative
values. New wheel and auxiliary powers use the same exact conversion instead
of the workbook's approximate divisor 9548.8. The baseline net mechanical
energy remains the signed trapezoidal integral starting at zero. Template
per-interval energy is a separate channel, using the existing
`sample_energy_kwh` function and actual sample intervals; its first sample uses
the first observed interval, as documented by that function. No Electric or
Hybrid numerical definitions were changed.

The final template contains unused electrical summary headings and stale
references to row 46810. Those cells do not define new Diesel battery or EDU
calculations. The existing Diesel KPI definitions and executive selection remain.

## Common presentation and traceability

Every numerical value in the final right-side summary is a direct Excel cell
reference. Existing bottom result cells are referenced when present; remaining
metrics reference their canonical statistic or KPI cells on `Statistics`.
Statistics remains visible in Diesel, Electric and Hybrid so users can inspect
the canonical values directly. Channel selection, calculations, KPIs and
existing summary formula references are preserved. Formula caches preserve
immediate values for readers that do not recalculate, and Excel retains
automatic recalculation.

Native Excel charts remain editable. Every data trace uses a solid line and
disabled markers. Native charts and Matplotlib figures share a deterministic
RGB palette indexed by configured series order, so the same plot composition
uses the same colours in Excel and PowerPoint. Primary and secondary axes
share the palette allocation, avoiding repeated colours between axes.

All three profiles reuse the established Diesel inner plot-area layout:
`x=0.13`, `y=0.14`, `width=0.74`, `height=0.64`. The same left and right
axis-title gutters reserve readable margins on single- and dual-axis charts.
Chart dimensions, titles, series, data references and numerical axis ranges
are unchanged by this layout adjustment.

Diesel presentation labels now read `Engine Torque & Power`, `Engine Power`
and `MIN ENGINE POWER`; the numerical sign convention is unchanged.

## Validation

Tests cover every template selection row and all eight chart mappings,
independent asymmetric corner/auxiliary arithmetic, negative power, irregular
time intervals, cumulative energy, missing optional inputs, canonical formula
targets and caches, native chart colours/styles, and actual Matplotlib line
properties. Existing Diesel golden-value tests continue to check the original
engineering results. Electric and Hybrid regression tests check their unchanged
channel selection, calculations and report content.

Validation logs are `outputs/template_extension_full.log` and
`outputs/template_extension_acceptance.log`. The full run collected 445 tests:
440 passed, and five assertions still expected the former Diesel output count
or literal summary values. Those assertions were updated for the requested
extension and formula traceability. The final targeted run passed all 40 tests,
including every previously failing case and the cross-profile presentation,
template coverage and extracted-package checks. No unresolved failures remain.

The baseline preservation audit is
`outputs/template_extension_baseline_preservation.json`; it checks every original
Diesel channel, math, statistic and KPI definition, original plots apart from the
requested labels, unchanged PowerPoint selections, and unchanged Electric
and Hybrid profiles.

Installed Excel recalculation checks passed for all 13 Diesel, 47 Electric and
53 Hybrid summary values. Editing a canonical Statistics cell propagated to its
summary in each workbook. Read-only validation closed without saving changes.
Results and three native Excel chart renders are in
`outputs/template_extension_office`. New native fuel and wheel-power figures
were visually reviewed for distinct solid traces and visible negative power.
At this review stage, no commits or pushes were made.

The final Excel layout review regenerated all three workbooks in their existing
review output directories. All 54 charts (18 Diesel, 15 Electric, 21 Hybrid)
were exported by installed Excel at their existing dimensions and visually
inspected, including dual-axis charts. Tick labels, axis titles and legends
are visible without clipping or overlap. Native Excel also verified that every
axis-title bounding box remains inside its chart boundary.

`outputs/excel_layout_review/structural_validation.json` confirms that Diesel
chart XML is byte-identical to the pre-adjustment output, and only chart layout
fields changed for Electric and Hybrid. Data-sheet, mapping-sheet, Statistics
and drawing XML remain byte-identical in all three workbooks. Native exports,
contact sheets and their inspection index are in `outputs/excel_layout_review`.
The final layout regression run passed all 96 targeted tests; its log is
`outputs/excel_layout_review/tests.log`. At this review stage, no files were
staged, committed or pushed.

## Final wording and release validation

The `range_85_battery_km` KPI display name is now `Range for 85% Battery`, replacing
`Range for 85% Battery + Range Extender`. Hybrid inherits this label from Electric;
no duplicate Hybrid override is needed. The KPI ID, expression, dependencies,
values and final-summary references are unchanged. No PowerPoint content or
Diesel definition changed during this cleanup.

All three Excel and PowerPoint reports regenerated successfully in their
existing review directories. Package-part comparisons confirm that every chart,
report data sheet, mapping sheet and embedded PowerPoint figure is byte-identical
to the manually reviewed version. Electric and Hybrid Statistics XML differs
only by the requested label. Regeneration updates ordinary timestamps and
configuration provenance metadata.

Read-only installed Excel validation recalculated all 113 direct summary
references, found no error cells, confirmed canonical-cell edit propagation,
and checked all 54 chart geometries and axis-title bounds against the prior
visual review. All three PowerPoints opened in installed PowerPoint. Live
Matplotlib checks verified 107 data series in 56 exports: solid lines, disabled
markers and distinct deterministic colours. Negative Diesel power and the
requested engine-power labels remain preserved.

The repository doctor passed with zero warnings and failures, source compilation
passed, and the complete generic pipeline passed all seven stages using the
historical Hybrid reference workbook matching its configured channel IDs.
The generic example's fixed channel IDs are not the latest raw CSV's IDs;
only the ignored validation configuration selects the matching workbook.
The deterministic release build and checksum/content checks passed. Its ZIP
contains only policy-approved runtime assets, source and configuration files,
with no generated reports, caches or private source datasets. The environment
package snapshot is retained locally. Final logs and validation records are
in `outputs/final_release_validation`.

The complete final regression suite passed: **445 tests in 671.32 seconds
(11 minutes 11 seconds)**, with no failures or skips. Command:
`.venv\Scripts\python.exe -u -m pytest -p no:cacheprovider --basetemp=outputs/final_all --junitxml=outputs/final_release_validation/pytest_final.xml`.
The shorter temporary directory avoids Windows' path-length limit during the
bootstrap test's nested virtual-environment creation; no tests were excluded
or weakened. Pytest's cache plugin is disabled because the existing cache
directory denies writes. `git diff --check` passed, and the final inventory
contains only 19 intended configuration, source, test and documentation files.
Generated reports and validation outputs remain ignored. Pre-existing invoice
and setup logs remain intact with local-only Git exclusions.
