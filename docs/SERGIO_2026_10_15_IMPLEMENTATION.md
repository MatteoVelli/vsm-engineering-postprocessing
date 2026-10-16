# Sergio requirements — 15 October 2026

This assessment records the initial implementation and the sources available at
that time. Subsequent real Electric validation, the selection correction and
Sergio's confirmed HV unit are covered by
[the HV assessment](HV_AUXILIARY_CONFIRMATION_2026_10_16.md). Complete real-data
validation of the updated Hybrid configuration remains pending.

## Implementation and use

The existing application now offers **Sergio - 15 October 2026** as the default
reporting revision after upload. Select Diesel, Hybrid or Electric as before.
**Legacy - Electric_05 / Hybrid_06 / Diesel** explicitly preserves the previous
profiles and conventions. No filename determines powertrain compatibility.

The new profiles are in `config/report_profiles/sergio_2026_10/`. They extend the
original profiles using validated, explicit definition overrides. The original
profiles remain available at their existing paths; no automatic driveshaft or
constant-zero high-voltage fallback is used by the new profiles.

Expand **Channels, statistics, plots and optional inputs** to select exported
channels, MATH channels, RMS, summary statistics, Excel column statistics, KPIs,
Excel plots and the plots/metrics in existing PowerPoint slide slots. Calculation
prerequisites are retained when a selected statistic, KPI or plot needs them.
Time remains exported. A plot whose data columns are not exported uses its shared
PNG in Excel, avoiding broken native chart references or unwanted data columns.
Select a sample interval to apply calculations/statistics to that input window.
Changing inputs, mappings, revision, window or selections invalidates validation
and previous downloads.

Diesel water temperature can be mapped to a temperature channel actually present
in the upload. The mapping retains its source units; temperature is not inferred
from oil, ambient temperature or other measurements. Custom mappings are saved
under the ignored `outputs/ui_workspace/profiles/` directory. The supplied older
Diesel CSV really contains `Engine_Water_Temperature`; that supported legacy name
is retained. Its presence does not resolve the missing mapping in the new workbook
or prove that the identical oil/coolant traces represent independent measurements.

## Features and status

| Area | Status and evidence |
| --- | --- |
| Diesel agrochemical mass | **IMPLEMENTED; VALIDATED WITH REAL VSM DATA** for the available CSV conversion. Raw hitch force [N] remains distinct from MATH mass [kg], using `-Fz/9.80665`. Time/distance plots remain available. |
| Diesel auxiliary powers | **IMPLEMENTED; VALIDATED WITH REAL VSM DATA AND WORKBOOK SOURCE ARRAYS**. Use each actual auxiliary torque and engine RPM, then sum the two mechanical powers; no correction multiplier or low-voltage double counting. |
| Diesel wheel torque/power | **IMPLEMENTED; VALIDATED WITH WORKBOOK SOURCE ARRAYS** and synthetic signed cases. The 17,872-sample workbook has valid zero front torques and nonzero rear torques. No front torque is fabricated. |
| Engine water temperature | **IMPLEMENTED** with export, plot, RMS/MAX/MIN/LAST and source selection. Legacy CSV mapping tested. **BLOCKED BY SOURCE DATA** for confirmation of the correct measurement in the new upstream run/template. Absence disables only this optional section. |
| Electric/Hybrid gearbox torque | **IMPLEMENTED**, all four exact `KERS_MotGenUnit_GearboxOutputTorque_1..4` mappings; signed corner/total power and torque tested using nonzero synthetic data. **BLOCKED BY SOURCE DATA** for meaningful real traction validation: `_07` measurements are all zero and the current raw exports omit these sources. |
| High-voltage auxiliaries | **IMPLEMENTED; VALIDATED WITH WORKBOOK SOURCE ARRAYS** using `User_defined_Data_High_Voltage_Auxiliary_signal1_0` plus the separate low-voltage source. Current raw CSV exports omit the new source; corresponding modern CSV analyses are explicitly unavailable. |
| Bulldozing/compaction | **IMPLEMENTED; VALIDATED WITH WORKBOOK SOURCE ARRAYS**, analytical and synthetic tests. Four-corner signed sums and accumulated kWh use actual time intervals. Both dual-axis plots are selectable. |
| Road gradient/height | **IMPLEMENTED**, shared distance basis and separate [%]/[m] axes. Both profiles map `Track_Height`, including Electric's omitted selection entry and Hybrid's `Road Heigh` label. Latest Electric CSV has height; latest Hybrid and Diesel CSVs do not. |
| Excel | **IMPLEMENTED**, original four-sheet structure, raw/MATH distinction, dynamic formula/chart ranges, editable charts and initial formula caches. Native Excel independently recalculated disposable cache-cleared reports. |
| Streamlit | **IMPLEMENTED**, existing upload/validation/report/download workflow, current/legacy revision choice, optional input mapping and selections. Actual Streamlit AppTest exercises validation and report generation. |
| PowerPoint | **IMPLEMENTED**, existing fonts, logos, slide slots and layouts retained. New plots can replace configured plots in existing slots. Unavailable plot slots are cleared and labelled; no stale reference chart is presented as a computed result. Empty selections remain valid. |

## Numerical conventions and findings

- Revised rotational power uses exact SI: `T_Nm * rpm * 2*pi / 60000` kW.
  Equivalent divisor: `9549.29658551372`. The templates/legacy profiles use
  `9548.8`, approximately **0.0052005% higher power**. Both signs are retained.
- Revised incremental energy is trapezoidal:
  `dE[0]=0`; `dE[i]=(P[i-1]/2+P[i]/2)*(t[i]-t[i-1])/3600`.
  Accumulated energy is the signed cumulative sum. There is no extrapolated
  pre-start interval. Diesel cumulative integration uses the same observed-time
  convention. Missing power is rejected, not imputed; duplicate, reversed or
  invalid timestamps are rejected. Optional invalid source channels disable
  their dependent sections with explicit diagnostics.
- Legacy profiles retain their previous `sample_energy_kwh` right-sample
  rectangles, including an initial extrapolated interval. They are labelled
  legacy in the UI; their energy results should not be compared as though they
  used the revised observed-interval convention.
- RMS is sample-based `sqrt(sum(x*x)/N_valid)` in the selected interval. The
  existing separate time-weighted statistics retain their own definitions.
  `_07` formulas divide 3,853 squared samples by 17,417; their result is about
  `sqrt(3853/17417) = 0.4703` of the correct RMS for those same squared values.
  No such divisor appears in revised output formulas.
  In each `_07` workbook, battery power RMS is **28.716636646 kW** from the
  3,853 actual samples versus the cached **13.506611298 kW**. Battery heatflow
  RMS is **2.840107066 kW** versus cached **1.335818768 kW**. Reconstructing
  the stale divisor reproduces both cached results; the discrepancy is numerical,
  not a different validated duty-cycle definition.
- Statistics use live Excel formulas with numeric-count guards. LAST uses the
  final eligible numerical sample. Empty numeric ranges return `NA()`, rather
  than a manufactured zero. Excel recalculation is automatic/full on opening.
  Selected summary metrics and unsupported functions retain traceable Python
  values where their source columns were explicitly not exported.
- Live formulas cover pointwise arithmetic, interval energy, cumulative sums,
  cumulative trapezoidal integrals and sample statistics. Header comments and the
  mapping sheet record the defining expression and dependencies. Formula caches
  are initial display values, not numerical authority.
- Both `_07` workbooks reproduce the four-corner soil **power sums exactly**.
  Independent observed-interval integration gives bulldozing **5.004134668674
  kWh**, compaction **7.163835208822 kWh**, and HV+LV auxiliaries
  **21.328880372222 kWh** in each supplied workbook. These are computations
  from identified source columns, not acceptance of cached energy formulas.
  Cached accumulated energies differ because those workbooks use sample
  rectangles: their final bulldozing/compaction values are 5.004248797153 and
  7.164000334808 kWh, respectively. The audit records both observed-interval
  trapezoidal values and reconstruction of the historical rectangle convention.
  The first accumulated bulldozing formula also looks ahead:
  `DP5=SUM(DN5:DN$6)` in Electric `_07`; subsequent rows anchor at row 6.
  Its first cached energy is approximately `5.12591e-6` kWh at the initial
  timestamp. This is an inconsistent initialization, not a physical starting
  energy. Revised accumulation starts at zero; no look-ahead range is copied.
- Diesel workbook auxiliary cached powers match its approximate divisor exactly.
  Exact-SI maximum differences are approximately **0.00062465 kW** (Auxiliary 1),
  **0.00104029 kW** (Auxiliary 2), and **0.00598146 kW** (total wheel power).
  These differences are explained by the conversion constant, not corrected
  by altering the source torques.

## Repository audit and reference scope

The initial working tree was clean at `0ec467c` (latest history also included
`0a59f82`, `fdb3ed9`, `141e5ad`, `d71da4f`). Importers, semantic profiles, math,
statistics, plotting, Excel, PowerPoint, UI, launch and release paths were reviewed.
No new reporting engine was introduced.

All three requested templates and their previous versions were present:

- `reference_files/Diesel Template_03.xlsx`: 17,872 samples, eight native charts;
  compared with `Diesel Template.xlsx`.
- `reference_files/Robo_Sprayer_Electrification_Tamplate_Electric_07.xlsx`:
  3,853 samples, 15 native charts; compared with Electric `_05`.
- `reference_files/Robo_Sprayer_Electrification_Tamplate_Hybrid_07.xlsx`:
  3,853 samples, 21 native charts; compared with Hybrid `_06`.

The existing approved Astauto PowerPoint references were available and retained.
No requested reference file was missing. Private new templates remain unchanged
and ignored by Git. Release packaging includes them when present; otherwise the
profiles generate reports in configuration order without requiring a private
workbook. An explicitly requested nonexistent template still raises an error.

Exact mapping checks verify Electric rows 208–211 / GY–HB and row 227 / HR,
and Hybrid rows 216–219 / HF–HI and row 235 / HY. Worksheet positions are audit
evidence; production resolution uses configured VSM names and units.

The latest raw inputs are different runs from the new workbooks:

| Raw CSV | Samples | Channels | New gearbox/HV sources | Height |
| --- | ---: | ---: | --- | --- |
| `Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.csv` | 223 | 586 | Not applicable to Diesel | Missing |
| `RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv` | 347 | 601 | Missing | Present |
| `Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv` | 1,770 | 600 | Missing | Missing |

The older 3,853-sample Electric reference CSV remains covered by the existing
regression suite. Different runs/sample counts are not compared as identical
source simulations.

## Validation evidence and commands

Run from the repository root in PowerShell. Final verification covers every test
module in three independent pytest processes with separate temporary directories.
The manifest in `outputs/sergio_verified_suite/manifest.json` records every file,
exact argument list, exit code and log; no test module is omitted or duplicated.

**Final regression result: 525 passed, 0 failed, 1 expected skip across all 34
test modules.** Group results: 160 passed (386.45 s), 214 passed/1 skipped
(304.85 s), and 151 passed (210.20 s). All three pytest exit codes were zero.
The corrected initial baseline was 474 passed. The additional acceptance suite
contains 51 passing cases and the one Diesel-inapplicable soil-slide skip.

The preceding full run found one legacy diagnostic-list compatibility failure
(524 passed/1 failed/1 skipped). That behavior was restored without changing the
test's assertion, and its entire 26-test statistics module passed before the
complete final verification above. Earlier RMS mock-profile and empty-selection
issues were fixed, with focused tests passing. No failing test remains.
After including this handover document in the optional client package assets,
the complete packaging module was rerun: **7 passed in 17.15 seconds**.
`git diff --check` reports no whitespace errors.

```powershell
# Initial baseline: default temp path was inaccessible (Windows sandbox).
.\.venv\Scripts\python.exe -m pytest
# Corrected pre-change baseline: 474 passed in 791.15 seconds.
.\.venv\Scripts\python.exe -m pytest -p no:cacheprovider --basetemp outputs/sergio_baseline_tmp

# Full run that exposed one legacy-diagnostics compatibility issue (since fixed).
.\.venv\Scripts\python.exe -m pytest -p no:cacheprovider -o addopts='' -q --basetemp outputs/sergio_final_full_tmp
# After the fix: every module, partitioned into three isolated pytest processes.
.\.venv\Scripts\python.exe outputs/run_sergio_verified_suite.py
.\.venv\Scripts\python.exe -m pytest -p no:cacheprovider -o addopts='' -q --basetemp outputs/sergio_package_final_tmp tests/test_release_packaging.py

# Reproduce reference inspection, hashes, cached comparisons and raw availability.
.\.venv\Scripts\python.exe scripts/audit_sergio_2026_10.py

# Generate explicitly labelled synthetic Excel and PowerPoint reports.
.\.venv\Scripts\python.exe scripts/validate_sergio_2026_10_reports.py --output-dir outputs/sergio_2026_10_validation/final_synthetic_reports
.\.venv\Scripts\python.exe scripts/validate_october_reporting.py --prepare-recalculation outputs/sergio_2026_10_validation/final_synthetic_reports/office_audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/validate_excel_recalculation.ps1 -AuditPath outputs/sergio_2026_10_validation/final_synthetic_reports/office_audit.json
.\.venv\Scripts\python.exe scripts/validate_october_reporting.py --audit-recalculation outputs/sergio_2026_10_validation/final_synthetic_reports/office_audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_sergio_2026_10_reports.ps1 -AuditPath outputs/sergio_2026_10_validation/final_synthetic_reports/office_audit.json
```

The reference audit and synthetic report commands are reproducible; independent
recalculation preparation preserves existing evidence and intentionally refuses
to overwrite its prior cache-cleared/recalculated files. Use a fresh output copy
for another Office audit.

Native Excel was launched outside the restricted session after COM launch failed
inside it. The existing recalculation script cleared all input formula caches,
called `CalculateFullRebuild` and calculated each sheet. The independent Python
audit checked **780 Diesel + 2,837 Electric + 2,955 Hybrid numeric cells**; the
maximum absolute difference was **1.7763568394e-14** in every report.
Sources and authoritative templates were not edited.

Desktop PowerPoint also exported the three synthetic decks to PDF/PNG. The
existing traction layout and configured Electric/Hybrid soil-plot variants were
visually inspected with their logos, KPI cards and chart slots preserved.

Analytical/synthetic tests include signed four-corner torque and power, zero front
Diesel torque, auxiliary sums, both soil sums, 10 kW for 3,600 seconds = 10 kWh,
zero/negative power, irregular timestamps, missing values, duplicate/reversed
time, sample-count independence, selected RMS intervals, configurable coolant,
optional height, legacy exports, genuine formulas/caches, dual-axis chart
dependencies, partial Excel/PowerPoint reports and actual Streamlit interaction.

One parametrized test is intentionally skipped: a Diesel invocation of the
Electric/Hybrid soil PowerPoint slot test. The source-dependent traction work is
reported as blocked, not represented by skipped synthetic tests or invented data.

## Important files changed

- `report_profile.py`: validated definition overrides, optional invalid-source
  diagnostics, source mapping traceability and moving/all-zero torque warning.
- `config/report_profiles/sergio_2026_10/*.yaml`: three current profiles with
  exact source mappings, optional dependency chains, energy expressions, water
  statistics and new plots. Original profile definitions remain unchanged.
- `math_engine.py`: signed observed-interval trapezoidal energy helper, reusing
  existing integration validation and expression evaluation.
- `importer.py`: explicitly permissive imports retain invalid cells as NaN for
  per-channel validation; strict imports retain their previous rejection policy.
- `profile_statistics.py`: include availability and source-quality diagnostics.
- `excel_formulas.py`, `excel_report_engine.py`: live temporal/statistic formulas,
  caches, dynamic ranges, partial sections, template revision ordering and
  configurable column statistics/exports/charts.
- `report_selection.py`: validated selections and shared dependency closure.
- `ui_app.py`, `ui_config.py`: current/legacy selection, optional coolant mapping,
  analysis interval, report selections and availability feedback.
- `profile_powerpoint_report_engine.py`, `diesel_powerpoint_report.py`: reuse
  existing slide slots, remove unavailable reference charts, allow empty
  selections and label the selected thermal measurement accurately.
- `release_builder.py`: package current private templates when available.
- `tests/test_sergio_october_15.py`: controlled and real-reference acceptance
  checks. `test_diesel_finalization.py` now expects the two explicitly requested
  new default plots and checks their IDs; its existing assertions remain intact.
- `scripts/audit_sergio_2026_10.py`,
  `scripts/validate_sergio_2026_10_reports.py`,
  `scripts/render_sergio_2026_10_reports.ps1`: reproducible source audit,
  clearly labelled synthetic artifacts and read-only Office rendering.

## Outstanding source dependencies

1. New Electric/Hybrid raw VSM exports containing all four exact gearbox-output
   torque channels, with meaningful nonzero operating torque and corresponding
   wheel rotational speeds. All-zero workbook measurements do not validate
   real traction performance.
2. New raw exports containing the exact high-voltage auxiliary source, so that
   the complete modern auxiliary/KPI chain can be validated from raw CSV.
3. Confirmation of the intended coolant source in the new Diesel VSM run. The
   older measured `Engine_Water_Temperature` channel is supported; no mapping
   is invented for a run that omits it.
4. `Track_Height` in updated Hybrid exports (and Diesel if height is requested).
   The required upstream VSM road-track reload cannot be performed by Python.
5. A raw export corresponding to the 17,872-sample Diesel workbook is needed for
   a full raw-to-workbook comparison of that exact run. The supplied 223-sample
   CSV independently exercises the supported calculations.

No commit, push, merge or publication is performed by this implementation task.

## Report artifacts

- `outputs/sergio_2026_10_validation/reference_audit.json`: source hashes,
  template mappings, actual sample counts, RMS and energy investigations.
- `outputs/sergio_2026_10_validation/final_synthetic_reports/`: explicitly
  labelled controlled-fixture Excel/PowerPoint reports, independent Excel
  recalculation evidence in `office_audit.json`, and desktop-rendered PDFs/PNGs.
- `outputs/sergio_2026_10_validation/final_real_reports/`: three real-CSV Excel
  and PowerPoint examples, with optional missing analyses excluded and recorded
  in `real_report_audit.json` and individual manifests.
- `outputs/sergio_verified_suite/manifest.json`: all final regression commands
  and results. The local runner is `outputs/run_sergio_verified_suite.py`.

Source timestamps retain their original values when selecting an analysis
window; configured LAST statistics keep their existing timestamp definition.
Observed-interval energy still starts at zero at the selected first sample.

## Final Git status

HEAD remains `0ec467c`. There are 15 modified tracked files and nine new files;
generated validation artifacts and the supplied private templates remain ignored.

```text
 M README.md
 M docs/CONFIGURATION_GUIDE.md
 M docs/KNOWN_LIMITATIONS.md
 M src/vsm_postprocessing/diesel_powerpoint_report.py
 M src/vsm_postprocessing/excel_formulas.py
 M src/vsm_postprocessing/excel_report_engine.py
 M src/vsm_postprocessing/importer.py
 M src/vsm_postprocessing/math_engine.py
 M src/vsm_postprocessing/profile_powerpoint_report_engine.py
 M src/vsm_postprocessing/profile_statistics.py
 M src/vsm_postprocessing/release_builder.py
 M src/vsm_postprocessing/report_profile.py
 M src/vsm_postprocessing/ui_app.py
 M src/vsm_postprocessing/ui_config.py
 M tests/test_diesel_finalization.py
?? config/report_profiles/sergio_2026_10/full_size_sprayer_diesel.yaml
?? config/report_profiles/sergio_2026_10/robosprayer_electric.yaml
?? config/report_profiles/sergio_2026_10/robosprayer_hybrid.yaml
?? docs/SERGIO_2026_10_15_IMPLEMENTATION.md
?? scripts/audit_sergio_2026_10.py
?? scripts/render_sergio_2026_10_reports.ps1
?? scripts/validate_sergio_2026_10_reports.py
?? src/vsm_postprocessing/report_selection.py
?? tests/test_sergio_october_15.py
```
