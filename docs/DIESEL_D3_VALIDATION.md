# Diesel Phase D3: PowerPoint integration and combined report enablement

> Historical stage record. The follow-up audit, current UI behavior and extracted
> client-package validation are recorded in
> [Diesel final validation](DIESEL_FINAL_VALIDATION.md), which is the authoritative
> final status. Final client delivery packaging remains a separate requested step.

## Starting state

Branch: `main`. HEAD: `0ec8529705a18bd88bd7524bdb816dc9a42efef1`.
The completed D2 suite passed 321 tests, with no skips or failures.
The pre-change D3 PowerPoint baseline passed **31 tests in 147.55 seconds**,
with no skips or failures. It covered both profile and legacy PowerPoint tests.

Initial tracked diff: **10 files, 449 insertions, 159 deletions**.
Initial status:

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

The starting hashes and Git snapshots are in
`outputs/diesel_d3_start_hashes.json`, `outputs/diesel_d3_initial_status.txt`
and `outputs/diesel_d3_initial_stat.txt`. The complete starting Diesel profile
is captured in `outputs/diesel_d3_profile_before.json`.

## Architecture assessment

The combined production wrapper generates Excel once and passes that same
`ProfileExcelReportResult` to the PowerPoint builder. Canonical profile metrics
are adapted into presentation statistics; existing Matplotlib profile PNGs
become the slide plot images. There is no second engineering calculation path.

The production profile renderer edits the approved ten-slide
`RoboSprayer_Electric_Report_Astauto_Colours.pptx` template, then appends road and
steering slides for Electric/Hybrid. Its slide IDs and narrative include
electrical assumptions. Reusing that narrative directly for Diesel would leak
battery/SOC/range content and demand unavailable metrics.

D3 therefore adds two Diesel-only dispatch points: configuration and rendering.
The separate Diesel adapter reuses the approved template, KPI-slot replacement,
canonical-statistics adapter, plot replacement, typography preservation, logo,
footer, numbering, file naming and manifest helpers. Existing Electric/Hybrid
rendering functions and their configuration remain unchanged.

Diesel uses the template's ten existing pages, with metric and plot selections
in its YAML. Empty optional-only pages are removed and numbering is updated.
Unused cards/panels disappear; a remaining single plot is centred at its original
dimensions. Electrical header icons are replaced with existing neutral template
icons, and unused electrical image relationships are removed from the Diesel
file. The presentation has native editable text/cards and the established
image-based plot treatment. Excel retains its native editable charts.

The existing isolated Diesel UI helper now calls the combined wrapper and uses
the shared Excel/PPT download renderer. Its result remains keyed to the current
source/profile validation key and machine name. Electric/Hybrid UI branches are
untouched. The Excel-only API remains available.

## D3 files and protected work

| File | D3-specific change |
|---|---|
| `config/report_profiles/full_size_sprayer_diesel.yaml` | Add ten presentation selections and update report-availability description |
| `src/vsm_postprocessing/diesel_powerpoint_report.py` | New isolated Diesel content/template adapter; optional content handling and source/metric notes |
| `src/vsm_postprocessing/profile_powerpoint_report_engine.py` | Two Diesel dispatch points; retain all pre-existing engine edits |
| `src/vsm_postprocessing/ui_app.py` | Upgrade only the Diesel helper to combined generation and both downloads |
| `src/vsm_postprocessing/ui_config.py` | Remove the Diesel combined-report guard |
| `tests/test_diesel_profile.py` | Update three phase-boundary expectations for D3; numerical assertions unchanged |
| `tests/test_diesel_excel_report.py` | Update two phase-boundary tests to verify PPT reuse/workbook preservation and combined downloads |
| `tests/test_diesel_powerpoint_report.py` | 29 new tests covering metrics, slides, styles, images, relationships and optional inputs |
| `docs/DIESEL_D3_VALIDATION.md` | This validation report |

Only six files present in the starting snapshot changed; the adapter, its test
file and this document are new. Other files in that snapshot retain their hashes,
including D1/D2 validation documents, the Excel engine, numerical engines,
Electric/Hybrid profiles, reference reports and existing PowerPoint tests.
No existing Electric/Hybrid/Caiman assertion was modified.

An exact YAML comparison confirms that `channels`, `statistics`, `kpis` and
`plots` are identical to D2, including units, requirements, expressions,
placements and plot flags. Only description and presentation configuration
changed. The reference Diesel CSV remains unchanged, SHA-256:
`e089fd94a9d6c6657982198c2a7106ac5cd36e0431cf5fb870c9565bb764d7b1`.

## Delivered artifacts

PowerPoint (10 slides, 10 plot images, 48 KPI-card placements, 24 distinct metrics):

`outputs/diesel_d3_final/profile_powerpoint_report/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.pptx`

Companion Excel:

`outputs/diesel_d3_final/profile_excel_report/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.xlsx`

Rendered PowerPoint PDF:

`outputs/diesel_d3_final/rendered/diesel_review.pdf`

All ten rendered slides: `outputs/diesel_d3_final/rendered/slide_01.png` through
`slide_10.png`, exported by installed Microsoft PowerPoint at 1600 × 900.
Additional optional-channel review images and synthetic review decks are in
that same folder. They are verification artifacts, not client packages.

Independent audit: `outputs/diesel_d3_final/independent_numerical_audit.json`.
Existing report manifests and summaries accompany both generated reports.
`diesel_template_reuse.json` records retained template pages and optional omissions.

PowerPoint SHA-256:
`ed09a624f9f510fff17dfe49da7efcc5a8d05fd660f164bd19e8de0e2c952aff`.

Excel SHA-256:
`480f1067052654fba5a3a8383f4d994db50f41ddf1fbfc2c694e447116e22740`.

## Slide inventory

| # | Title and purpose | Primary KPIs | Plots |
|---|---|---|---|
| 1 | Full Size Sprayer Diesel — identify the run | Distance, cycle time, fuel consumed, maximum power/torque/speed | None |
| 2 | System & Simulation Overview — source, samples and channel availability | Time, distance, gradient minimum/maximum | None |
| 3 | Executive Results — compact engineering snapshot | Time, distance, maximum speed, fuel consumed, maximum power/torque/RPM, average flow | None |
| 4 | Vehicle Operation — speed and road demand | Distance, time, maximum speed, gradient extrema | Vehicle Speed & Road Gradient |
| 5 | Diesel Engine Performance — operating envelope and signed overrun | Maximum/average RPM, maximum torque, maximum/minimum/RMS power | Engine Speed; Engine Torque & Signed Power |
| 6 | Fuel Consumption & Mechanical Energy — precise fuel and net-energy results | Fuel consumed, average/maximum flow, integrated volume, net energy | Cumulative Fuel Mass; Fuel Flow |
| 7 | Driveline & Wheel Loads — torque and tyre load demand | FL/RR driveshaft torque maxima and FL/RR tyre vertical-load maxima | Four-corner Driveshaft Torque; Four-corner Tyre Vertical Loads |
| 8 | Engine Load & Throttle — source controls in their actual units | Maximum engine load torque [Nm], maximum throttle [%] | Engine Load Torque; Engine Throttle |
| 9 | Engine Thermal Behaviour — source temperature context | Oil temperature minimum/maximum | Engine Oil Temperature |
| 10 | Simulation Summary — descriptive run conclusion | Distance, time, fuel consumed, maximum power/torque, net energy | None; summary text also includes maximum RPM and signed power extrema |

All plots use the unchanged D2 time axis in seconds. All four available corners
are plotted on the driveline slide; its cards select two diagonally opposed
corners to keep both torque and vertical load visible without overcrowding.

## Numerical validation

The independent audit reopens the final PowerPoint and compares all 55 canonical
values against the saved D1 numerical JSON. It locates every displayed KPI card
by its label and position, parses the actual visible number/unit, and compares
it to the saved D1 value at the displayed precision. All **48 cards** pass.
Slide notes also retain full-precision canonical values and the exact source
filename. There is no engineering discrepancy.

| Metric | D1/D2 canonical value | PowerPoint display |
|---|---:|---:|
| Duration | 222 s / 3.7 min | 3.70 min |
| Distance | 0.999242 km | 1.00 km |
| Maximum speed | 22.4686 kph | 22.47 kph |
| Minimum gradient | −15% | −15.00% |
| Maximum gradient | 15% | 15.00% |
| Maximum RPM | 1996.81 rpm | 1,997 rpm |
| Average RPM | 1537.214164414414 rpm | 1,537 rpm |
| Maximum torque | 826.219 Nm | 826.22 Nm |
| Maximum signed power | 165.938626339644 kW | 165.94 kW |
| Minimum signed power | −8.230092194144 kW | −8.23 kW |
| Power RMS | 74.570694135737 kW | 74.57 kW |
| Net mechanical energy | 3.764781781161 kWh | 3.76 kWh |
| Fuel consumed | 0.772836 kg | 0.77 kg |
| Integrated fuel volume | 0.942120719458 L | 0.942 L |
| Average fuel flow | 15.277633288514 L/h | 15.28 L/h |
| Maximum fuel flow | 41.4741 L/h | 41.47 L/h |

The existing presentation formatting deliberately rounds values, including RPM
to integers and mass to two decimals. This is display precision, not altered
calculation. Exact values remain in Excel and presentation notes.

Plot-image bytes are verified against the existing D2 plot assets, with no crop
or data transformation. All twelve negative-power samples remain present. The
engine plot visibly extends below zero and the minimum-power card reads −8.23 kW.
Engine load remains Nm; fuel flow uses L/h and is never called fuel economy.
Synthetic shifted-time/nonzero-counter tests confirm use of cycle differences,
not final absolute time or final cumulative fuel mass.

## Excel preservation

The final companion workbook was compared directly with the delivered D2 XLSX.
The following package parts are **byte-identical**:

- Main report worksheet, including all data, summary, formatting and bottom statistics.
- Channel mapping worksheet.
- Statistics worksheet.
- Workbook styles.
- All ten native chart XML parts.

Whole-file hashes differ because profile provenance now includes presentation
configuration and generation metadata. No D2 numerical or visible Excel change
was made. Building PowerPoint from an existing Excel result also leaves the
existing workbook bytes untouched, verified by a regression test.

The Electric and Hybrid reference-fixture decks generated by the final suite
were also compared with the pre-D3 baseline outputs: all **104 Electric** and
**106 Hybrid** PowerPoint package parts are byte-identical. ZIP-container
timestamps are excluded from that package-content comparison.
The comparison was extended to all seven generated baseline decks, including
dynamic Electric/Hybrid variants, the latest Electric dataset, Caiman Hybrid
and its traction-mapping case. Every package part matches; Caiman decks contain
115 parts each. Details are in
`outputs/diesel_d3_final/existing_powerpoint_preservation.json`.

## Rendered visual review

The approved Electric and Hybrid/Caiman references were inventoried with
python-pptx and selected cover/content pages rendered for comparison. Diesel
uses the production Electric template, as the current combined engine does.
The different historical Hybrid cover is not a new template choice.

PowerPoint could not open the file from the sandbox. The read-only automation
run outside the sandbox succeeded, with no visible presentation window and no
changes saved to the source PPTX. Temporary pytest folder permissions also
required copying synthetic review decks into the ordinary workspace review
directory before rendering. No host settings were changed.

Every reference Diesel slide was inspected as an actual PowerPoint PNG render.
After refinements, all ten final images were compared pixel-by-pixel with that
review: only slides 2 and 7 changed, and both were inspected again.

| Slides | Review findings |
|---|---|
| 1, 3, 10 | Approved dark treatment, Cambria titles and KPI cards retained; readable values and narrative; consistent logo/footer/numbering; no clipping or overlap observed |
| 2 | Source filename wraps within the card; Diesel profile, 223 samples and unavailable height are explicit; neutral source/context icons; overview cards and bottom KPIs aligned |
| 4 | Speed and gradient axes/legend readable; no fabricated height; five supporting cards fit the established row |
| 5 | Both engine plots readable at original pair dimensions; negative-power region and minimum-power card visible; Nm/rpm/kW correctly distinguished |
| 6 | Cumulative mass and flow labels are precise; L/h and L are correct; net energy remains signed; cards and plots fit |
| 7 | Four-corner plots remain legible; torque and tyre-load cards share the row; missing-corner note fits above footer |
| 8 | Load torque is Nm and throttle is %; source plots and two expanded cards are readable |
| 9 | One centred oil trace avoids duplicating identical coolant data; source-identity note is derived from actual trace equality; temperature units and extrema are correct |

Extra renders cover the seven-slide all-optional-absent deck's executive and
fuel pages, plus the missing-driveshaft deck's wheel-load page. Missing cards
are removed, remaining plots are centred and no placeholder results appear.

Structural tests also check slide bounds, title geometry/font, inherited
backgrounds, every logo/footer/page number, exact plot bytes, zero image crop,
successful reopening and all internal package relationships. No visible
Battery/SOC/EDU/Hybrid content remains in Diesel. Rendering has been verified
on this installed PowerPoint version, not every presentation application.

## Tests

| Run | Passed | Skipped | Failed | Runtime |
|---|---:|---:|---:|---:|
| Pre-change PowerPoint baseline | 31 | 0 | 0 | 147.55 s |
| Diesel D1 + Excel + PowerPoint | 88 (36 + 23 + 29) | 0 | 0 | 55.88 s |
| Same targeted group without private references | 50 | 38 | 0 | 55.91 s |
| Complete regression suite | **350** | **0** | **0** | **520.54 s** |

Final full-suite groups (all passed, no skips): Diesel D1 **36**, Diesel Excel
**23**, Diesel PowerPoint **29**, existing PowerPoint **31**, existing Excel
and chart readability **42**, UI configuration **15**. The full JUnit XML records
each case and its duration; these groups are part of the full run, not extra
independent suite invocations.

The first targeted run exposed a brittle assertion in the new Diesel style
test (78 passed before failure): it assumed a title's shape index survived
icon replacement. The test now locates the title by its text and applies the
same geometry/font assertions. No existing numerical expectation or existing
Electric/Hybrid test was changed to bypass a failure.

Full-suite command:

```powershell
.venv/Scripts/python.exe -m pytest -q -x -p no:cacheprovider -o addopts='' --basetemp=outputs/d3all1 --junitxml=outputs/diesel_d3_full.xml
```

Logs: `outputs/diesel_d3_baseline.log`, `outputs/diesel_d3_targeted_final.log`,
`outputs/diesel_d3_portable.log`, `outputs/diesel_d3_full.log`.

## Limitations and deferred work

- Road height is absent in the reference. Gradient analysis remains valid.
- Missing optional metrics/plots are omitted; optional-only slides disappear
  when no usable plot remains. With all optional channels absent, the deck has
  seven slides and four plots. All-zero plots are omitted from PPT without
  changing Excel selections or canonical results.
- Oil and coolant are identical in this reference. Oil alone is plotted; a
  dynamic note reports equality only when the supplied arrays actually match.
  Coolant-only inputs retain their Excel data but do not invent an oil plot.
- D1 source/time/counter requirements and sampled-integration limitations remain.
  No density conversion, cycle BSFC, distance/area fuel-economy formula or
  acceptance criterion was introduced.
- Presentation values use established display rounding. Very small signed
  values can display as −0.00 at that precision; full precision is retained.
- The final portable/client ZIP and clean-package delivery validation remain
  deferred. Test-suite packaging fixtures are not a client release.
- No staging, commits, pushes, repository cleanup or final packaging performed.

## Final Git audit and assessment

`git diff --check` passes. Final tracked `git diff --stat`, which includes all
pre-existing and D1/D2 work and excludes untracked new files:

```text
 docs/SERGIO_REPORT_CHANGES.md                      |  27 +++
 reference_files/README.md                          |   6 +
 src/vsm_postprocessing/excel_report_engine.py      |  68 +++++-
 src/vsm_postprocessing/math_engine.py              |  26 ++
 src/vsm_postprocessing/profile_math.py             |   5 +
 .../profile_powerpoint_report_engine.py            | 264 ++++++++++-----------
 src/vsm_postprocessing/ui_app.py                   |  37 +++
 src/vsm_postprocessing/ui_config.py                |   2 +-
 tests/conftest.py                                  |   5 +
 tests/test_profile_powerpoint_report.py            | 172 +++++++++++++-
 10 files changed, 452 insertions(+), 160 deletions(-)
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
?? docs/DIESEL_D3_VALIDATION.md
?? logs/
?? src/vsm_postprocessing/diesel_powerpoint_report.py
?? src/vsm_postprocessing/diesel_validation.py
?? tests/test_diesel_excel_report.py
?? tests/test_diesel_powerpoint_report.py
?? tests/test_diesel_profile.py
```

Final acceptance assessment:

| Workflow | Result |
|---|---|
| Electric | PASS; numerical/configuration/tests unchanged; generated PPT package content identical to baseline |
| RoboSprayer Hybrid | PASS; numerical/configuration/tests unchanged; generated PPT package content identical to baseline |
| Caiman Hybrid | PASS; existing regressions unchanged; both generated Caiman deck variants identical to baseline |
| Diesel D1 | PASS; all numerical definitions and assertions retained; saved-D1 comparison passes |
| Diesel Excel | PASS; all visible worksheet/chart/style parts identical to D2; generation from the same canonical results |
| Diesel PowerPoint | PASS; 29 new tests, 48 independently checked cards, ten rendered slides reviewed |

**D3 is complete. The deterministic application is functionally complete across
Electric, Hybrid/Caiman and Diesel within the validated scope and documented
source assumptions.** This does not assert vehicle acceptance or client-release
readiness. Work stops after D3.

Remaining work after D3 is final whole-application review, repository cleanup
and logical commits when requested, then client packaging later when requested.
