# Full Size Sprayer Diesel: final validation

Audit date: 2026-09-30. Branch `main`; audited HEAD
`2403c9488ea901207a42988719d2bf6d29758478`. D1/D2/D3 retain their older
historical commit identifier; the final audit uses the current working tree.

**DIESEL READY TO PRESENT / RELEASE.** No remaining blockers within the validated
profile/input scope. This is software/reporting readiness, not vehicle acceptance
against unspecified engineering criteria. Final delivery ZIP creation remains
deferred until requested.

## Scope and architecture

D1, D2 and D3 are successive numerical, Excel and PowerPoint implementation
stages using **one** authoritative Diesel dataset. Their historical release
boundaries are superseded by this document. No separate Diesel reference XLSX
or PPTX is installed. Existing Electric/Hybrid work, source assets, personal
invoice and logs are protected; no staging, commits, pushes or final client ZIP.

The active explicitly selected profile is
`config/report_profiles/full_size_sprayer_diesel.yaml` (UI label `Diesel`, report
name `Full Size Sprayer Diesel`). It does not inherit Electric/Hybrid.
Strict import -> exact/normalized channel and unit resolution -> Diesel input
validation -> shared deterministic math/statistics -> profile plots -> native
Excel -> Diesel PowerPoint adapter. PowerPoint consumes the same Excel result,
canonical metrics and plot data, without a second engineering calculation path.
No speculative aliases, source fallbacks or powertrain inference are introduced.

## Reference input and raw validation

`reference_files/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.csv`

SHA-256: `e089fd94a9d6c6657982198c2a7106ac5cd36e0431cf5fb870c9565bb764d7b1`.

| Check | Independently observed |
|---|---|
| Layout | Comma-separated; names row 1, units row 2, numeric rows 3-225 |
| Samples / raw channels | 223 / 586 |
| Time/index | `Track_Time [s]`, 0 through 222, 1 s spacing, 222 elapsed intervals |
| Duplicate headers / malformed-width rows | 0 / 0 |
| Missing, nonnumeric or non-finite numeric cells | 0 |
| Required inputs | All seven present with matching units |
| Optional inputs | 17 of 18 present; only `Track_Height` absent |
| Selected output channels | 24 raw + 3 MATH = 27, with no duplicates |
| Electric compatibility | 38 ElectricSystem channels are zero; both Electric/Hybrid profiles reject this source with 34 missing mandatory inputs |

Strict import rejects missing, nonnumeric and non-finite cells. Diesel validation
also rejects invalid resolved arrays supplied directly to the math API, ambiguous
matches, unit mismatches, negative RPM/flow, negative/resetting fuel or distance
counters, and fewer than two strictly increasing time samples. Signed torque,
power and load are valid. Optional absence omits dependent results; present
invalid optional data is an error, never imputed or silently discarded.
Calculated MATH channels are not requested from the source CSV.

Required source channels:

| Source | Unit |
|---|---|
| `Track_Time` | s |
| `Track_Distance` | m |
| `Chassis_Speed` | kph |
| `Track_Gradient` | % |
| `Engine_Speed` | rpm |
| `Engine_Torque` | Nm |
| `Engine_FuelConsumption_absolut` | kg, cumulative mass |

Optional source channels (18):

| Source or four-corner family | Unit | Count |
|---|---|---:|
| `Track_Height` | m | 1 |
| `Engine_FuelConsumption_volumeflow` | L/h (source `l/h`) | 1 |
| `Engine_FuelConsumption_specific` | g/kWh | 1 |
| `Engine_Load` | Nm, not percent | 1 |
| `Engine_Throttle` | % | 1 |
| `Engine_Oil_Temperature`, `Engine_Water_Temperature` | degrees C | 2 |
| `Chassis_Acceleration_Longitudinal`, `_Lateral`, `_Vertical` | m/s^2 | 3 |
| `Tyre_Fz_FL`, `_FR`, `_RL`, `_RR` | N | 4 |
| `DriveShaft_Torque_FL`, `_FR`, `_RL`, `_RR` | Nm | 4 |

## Independent numerical audit

`scripts/validate_diesel_release.py` reads the original CSV with Python's CSV
reader, calculates expected values with scalar arithmetic and independent
trapezoids, generates fresh production reports, then reopens them. It does not
use the application's expression evaluator for expected values. All **6,021**
selected raw/MATH sample values and **55** available canonical statistics/KPIs
match; maximum Excel absolute serialization difference is **5.684341886080802e-14**.
All 48 PowerPoint metric-note placements agree with independently calculated
values; existing Diesel tests separately verify displayed values and units.

Let `t` be elapsed source seconds, `T` signed engine torque in Nm, `n` RPM,
`q` volume flow in L/h, and `I(x)[i] = sum(j=1..i)
((x[j-1]+x[j])/2 * (t[j]-t[j-1]))`, with `I(x)[0] = 0`.

| MATH channel | Formula / units | First | Final |
|---|---|---:|---:|
| `engine_power` | `T * n * 2*pi / 60000`, kW | -0.0447520857147 | 11.2353108575304 |
| `engine_mechanical_energy` | `I(engine_power)/3600`, net signed kWh | 0 | 3.764781781161 |
| `fuel_volume` (optional) | `I(q)/3600`, L | 0 | 0.942120719458 |

There is no extrapolated first interval. Power retains 12 negative samples;
negative power reduces net energy. Mass uses the source cumulative counter,
with no density conversion. Source specific consumption is exported but not
averaged into a cycle BSFC or used to infer efficiency. Fuel volume is a sampled
flow estimate and need not reproduce the simulator's internal mass integrator.

Five KPIs, all available on the reference:

| KPI | Independent formula | Reference value |
|---|---|---:|
| `drive_cycle_minutes` | `(t[-1]-t[0])/60` | 3.7 min |
| `distance_km` | `(distance[-1]-distance[0])/1000` | 0.999242 km |
| `fuel_consumed_kg` | `mass[-1]-mass[0]` | 0.772836 kg |
| `engine_speed_average` | `I(n)[-1]/(t[-1]-t[0])` | 1537.214164414414 rpm |
| `fuel_flow_average` (optional) | `I(q)[-1]/(t[-1]-t[0])` | 15.277633288514 L/h |

The profile configures **52 statistics**, with **50 available**; the two absent
height extrema are explicitly unavailable. Every configured available operation
is independently evaluated: MAX/MIN over source or derived arrays, FIRST/LAST
at actual endpoints, and sample RMS `sqrt(sum(x*x)/N)`. RMS is not time-weighted.
Units remain those of the target channel. All resolved data must be finite;
no NaN omission is applied to Diesel. Counter differences support nonzero initial
offsets; synthetic tests cover these and irregular time intervals.

| Key statistic | Value |
|---|---:|
| Speed maximum | 22.4686 kph |
| Gradient minimum / maximum | -15 / +15 % |
| RPM minimum / maximum | 688.925 / 1996.81 rpm |
| Torque minimum / maximum | -71.9229 / 826.219 Nm |
| Power minimum / maximum / RMS | -8.230092194144 / 165.938626339644 / 74.570694135737 kW |
| Fuel flow maximum | 41.4741 L/h |
| Engine load torque maximum | 825.455 Nm |
| Oil and coolant minimum / maximum | 49.3053 / 79.9948 degrees C |

Independent power spot checks include sample 0 (-0.0447520857147 kW), sample 1
(2.06148871147884 kW), sample 50 (65.5488921183535 kW), maximum at sample 145,
minimum at sample 151, and final sample 222. Determinism, time-origin invariance,
zero RPM, signed overrun, invalid inputs and optional omissions are also tested.

Auxiliary consumption, tyre losses, wheel power, density/efficiency assumptions,
fuel-per-distance/area and composed work/duty-cycle calculations are not part of
this Diesel profile. Tyre vertical loads and driveshaft torques are measured
source channels with extrema/RMS; no unrelated RoboSprayer signal is substituted.

## Excel and PowerPoint acceptance

Fresh artifacts are generated under `outputs/diesel_final_audit/`:

- `profile_excel_report/<source stem>.xlsx`.
- `profile_powerpoint_report/<source stem>.pptx`.
- `independent_audit.json` (full raw inventory, expressions and all 55 values).
- `rendered/excel_review.pdf`, `rendered/excel_page_1.png` through `_3.png`.
- `rendered/powerpoint_review.pdf`, `rendered/slide_01.png` through `_10.png`.

The workbook reopens in openpyxl and installed Excel. It has the main report,
mapping, hidden Metadata and visible Statistics sheets; no Plot Templates sheet,
embedded chart images, duplicated channels, error cells or worksheet formulas.
Numerical values are materialized, so there are no formula calculation/cache
dependencies. All ten editable native charts reference the intended semantic
columns and precisely rows 5-227, with source time on X. MATH classification and
units are preserved. The rendered three-page inspection PDF includes the
executive summary and all charts: titles, ticks, legends and both axes are
readable, including the negative-power region. Temporary print settings are
not saved to the workbook.

The ten plots are speed/gradient, engine RPM, torque/signed power, engine load,
throttle, fuel flow, cumulative fuel mass, four-corner driveshaft torque,
four-corner tyre vertical loads and oil temperature.

The presentation reopens in python-pptx and installed PowerPoint. Ten slides
and ten plot images were rendered and inspected. The Astauto logo, footer,
Cambria titles, page numbering, backgrounds and card style are retained. There
are no missing images, broken internal relationships, blank/duplicated slides,
electrical/hybrid wording or observed clipping. KPI cards agree with Excel at
their displayed precision. Image bytes match the canonical plot assets and
image crops are zero. Machine-name overrides change the visible title.

Diesel intentionally uses ten template pages rather than Electric/Hybrid's
twelve: no battery/generator story or appended road/steering content is inferred.
With all optional inputs absent, seven slides/four plots remain. Optional-only
pages and all-zero PowerPoint plot panels are omitted; Excel retains available
zero traces. Both oil/coolant are exported; oil alone is plotted and the equality
note is shown only when arrays match. No Diesel reference workbook formula
defect is applicable: no Diesel reference workbook exists. Exact SI power and
elapsed integration intentionally remain independent of legacy Hybrid's divisor
and sampled-energy convention. No Electric/Hybrid numerical behavior changes.

Desktop Office rendering required execution outside the filesystem sandbox;
reports were opened read-only, with Excel macros/events disabled. Individual
hidden Excel chart PNG exports were unreliable (some zero-byte files); visual
acceptance uses the complete, successful Excel PDF export instead.

## Streamlit and package acceptance

Finalization fixes found during audit:

1. Switching away from Diesel and back previously reused old validation. Source
   or profile changes now clear validation and both report/download states.
   Failed revalidation cannot retain an earlier success. Machine-name changes
   hide downloads generated for the previous name.
2. User instructions omitted Diesel. They now describe all three explicit profiles.
3. The declared Streamlit minimum was 1.37 despite existing `width="stretch"`
   calls. It is now `>=1.49.1,<1.50`, preserving the upper bound and width cleanup.
   Streamlit's versioned [image documentation](https://docs.streamlit.io/1.49.0/develop/api-reference/media/st.image)
   and [button documentation](https://docs.streamlit.io/1.49.0/develop/api-reference/widgets/st.button)
   document the supported width interface. Client-version runtime tests pass.

Nine new acceptance tests exercise actual Streamlit AppTest execution, source
changes/removal/invalid data, Electric/Hybrid profile transitions, failed
revalidation, machine overrides, both report downloads, ten previews, and an
extracted client package. For Streamlit 1.49 AppTest, only the unsupported upload
test driver is supplied; validation and generation are the production functions.
Separate browser automation exercises the actual file uploader and download links.

The real browser run on Streamlit **1.49.1** passed: app launch, raw CSV upload,
explicit Diesel selection, required/optional/MATH feedback, validation, machine
override, generation, both actual downloads, and ten fully loaded preview images.
Downloaded XLSX/PPTX files reopen and the cover reads `Acceptance Sprayer Diesel`.
Machine changes hide previous downloads; Electric/Hybrid/Diesel transitions
invalidate old validation. A CSV missing `Engine_Torque` names that required
source and disables generation. There were no application exceptions or obsolete
width warnings. Browser evidence is in `outputs/diesel_final_audit/browser/`.
Download buttons are the current access controls; no extra desktop-open control
was introduced. Browser checks wait for Streamlit reruns/upload completion to
finish before asserting final state.

The release builder recursively collects `config`, `src` and `scripts`, including
the Diesel YAML, validator, presentation adapter and UI. Existing runtime logo
and PowerPoint template assets are mandatory and included. No new runtime
numerical/report dependencies are needed; pyproject already declares NumPy,
openpyxl, PyYAML, Matplotlib and python-pptx. A validation-only ZIP is extracted
and a separate Python process loads the extracted source, discovers Diesel,
validates a synthetic required-only CSV and generates Excel/PPT successfully.
The package excludes private CSVs, invoice and logs. Existing packaging tests
cover deterministic ZIP content and required runtime assets. Final combined
Sergio delivery ZIP creation remains deferred as requested.

## Regression evidence and protected work

| Run | Result |
|---|---|
| Initial Diesel profile/Excel/PPT + UI focused run | 103 passed, 0 skipped, 0 failed |
| Final focused run including nine new acceptance tests | 112 passed, 0 skipped, 0 failed (133.08 s) |
| Nine new acceptance tests on Streamlit 1.49.1 | 9 passed, 0 skipped, 0 failed (58.49 s) |
| Full suite | **397 passed, 0 skipped, 0 failed (542.04 s)** |
| `git diff --check` | PASS |
| `git grep -n "use_container_width" -- "*.py"` | No matches in tracked Python |

The complete suite includes Electric, Hybrid/Caiman, Diesel, legacy reports,
native-chart readability, release packaging and client bootstrap regressions.
Electric and Hybrid remain regression-clean. Final commands and machine-readable
results:

```powershell
.venv/Scripts/python.exe -m pytest tests/test_diesel_profile.py tests/test_diesel_excel_report.py tests/test_diesel_powerpoint_report.py tests/test_diesel_finalization.py tests/test_ui_config.py -q -p no:cacheprovider --basetemp=.pytest_tmp/dfocus2 --junitxml=.pytest_tmp/diesel_focus_final.xml
.venv/Scripts/python.exe -m pytest -p no:cacheprovider --basetemp=outputs/dfa --junitxml=.pytest_tmp/diesel_full_final.xml
```

Use a new `--basetemp` for subsequent runs to preserve these test artifacts.
Independent audit reproduction (also use a new destination):

```powershell
.venv/Scripts/python.exe scripts/validate_diesel_release.py reference_files/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.csv outputs/diesel_final_audit
```

The expected full count increases from 388 to 397 solely from nine new tests.
One initial full invocation was interrupted after the old bootstrap dependency
assertion failed; that assertion was updated to the intentionally raised minimum.
No engineering test expectation was relaxed. The first new UI test used an
incorrect AppTest preview element name; its selector now supports both tested
Streamlit versions without changing the ten-preview requirement.

Finalization changes: `pyproject.toml`, `src/vsm_postprocessing/ui_app.py`,
`tests/test_client_bootstrap.py`, new `tests/test_diesel_finalization.py`, new
`scripts/validate_diesel_release.py`, `docs/CONFIGURATION_GUIDE.md`,
`docs/KNOWN_LIMITATIONS.md`, `docs/SERGIO_REPORT_CHANGES.md`,
`reference_files/README.md`, historical D1/D2/D3 status notices, and this document.
Existing Diesel calculations, profile, Excel/PPT engines and existing numerical
assertions remain unchanged during finalization. SHA-256 comparison against the
starting inventory protects unrelated PowerPoint/theme work, both Electric/Hybrid
profiles, reference assets, personal invoice and logs.

Remaining scope limitations: one real Diesel source dataset; sampled integration;
finite/nonresetting source requirements; optional height absent; no inferred
density, efficiency or agronomic consumption; PowerPoint display rounding;
cross-viewer visual fidelity beyond installed Office is not claimed. These are
explicit supported-scope boundaries, not fabricated channels or results.

Final Git state: `main` at `2403c94`; 14 modified tracked paths and 14 untracked
entries, including all pre-existing work and the preserved invoice/logs. The
index is empty. The full status and preservation checks are saved in
`outputs/diesel_final_audit/protected_work_audit.json`. No staged changes, commits
or pushes; no source reset, restore, stash or deletion; no manual modification
of historical generated outputs.

Proposed commit boundary (not executed): isolate pre-existing shared PowerPoint
theme work as its own prerequisite change, then include the Diesel profile,
validator/adapter, Diesel dispatch/shared-engine hunks, UI state/dependency fix,
fixtures/tests, audit script and Diesel documentation. Review mixed shared files
by hunk; do not stage the entire working tree. Exclude private CSVs, invoice,
logs, generated reports and validation-only ZIPs. Suggested message:

`feat: finalize full-size sprayer Diesel reporting`
