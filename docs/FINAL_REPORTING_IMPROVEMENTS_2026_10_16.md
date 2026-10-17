# Final reporting improvements — 16 October 2026

## A. Initial assessment

The modern Diesel profile inherited a thermal slide selecting only oil MIN/MAX
and the oil chart. The adapter explicitly explained identical oil/water traces
by showing one chart. Both measured channels and water statistics already
existed. The current Diesel Excel report already contained 20 native charts.

The Statistics worksheet stored configured KPI values as Python numbers. Its
simple statistics and the modern MATH columns already supported live formulas
with deterministic caches. Main-sheet KPI summaries referenced the static KPI
table. Existing safe expression parsing, Python calculations and formula-cache
serialization were suitable foundations for completing the dependency chain.

## B. Selected solution

The current Diesel presentation profile now selects oil MIN/MAX, water MIN/MAX,
and two independent temperature plots. The shared renderer applies common
temperature limits and distinct orange/blue colours. The adapter uses the
existing two-chart template slots, retains zero-valued measurements, and names
missing measurements without substituting another source. When both temperature
sources are missing, the current thermal slide remains with an explanatory note.
Explicit legacy Diesel profiles retain their earlier presentation selection.

The new `profile_kpi_formulas.py` compiles the existing expression AST into Excel
references. It reuses live MATH columns only when their complete prerequisite
chain is live. Otherwise it emits formula helpers, including the legacy sample
rectangle convention where configured. It preserves signed observed-interval
trapezoids for modern profiles, first/last valid statistics, cumulative sums,
time-weighted means and the median of valid Battery Energy / SOC estimates.

The hidden Formula Support worksheet holds documented expression parameters,
SOC validity tolerance, formula helpers and any raw prerequisites omitted from
the visible export selection. Raw helper headers identify the original channel
and unit. KPI cells reference documented parameter cells, including the existing
85% usable-capacity assumption. Unsupported future expressions fail explicitly
instead of receiving a misleading formula over a static result.

Visible report layout, number formats, source samples, engineering definitions,
MATH formulas, units, chart placement and Python PowerPoint values are preserved.
Comparison with the previously accepted real reports confirmed byte-identical
serialized main-sheet data cells (14,941 Diesel; 1,304,820 Electric), identical
MATH data-cell formulas (4,460 Diesel; 142,344 Electric), and identical native
chart XML for all 20 Diesel and 19 Electric charts.
Excel is used only for acceptance validation; normal generation requires no
Office installation and no new dependency.

## C. Modified files

- `config/report_profiles/sergio_2026_10/full_size_sprayer_diesel.yaml`
- `src/vsm_postprocessing/diesel_powerpoint_report.py`
- `src/vsm_postprocessing/excel_report_engine.py`
- `src/vsm_postprocessing/excel_formulas.py`
- `src/vsm_postprocessing/profile_kpi_formulas.py` (new)
- `src/vsm_postprocessing/plotting_engine.py`
- `src/vsm_postprocessing/profile_plotting.py`
- `tests/test_final_reporting_improvements.py` (new)
- `tests/test_diesel_excel_report.py`
- `tests/test_profile_excel_report.py`
- This validation report.

## D. Diesel PowerPoint validation

Native desktop PowerPoint rendered the final thermal slide. Visual inspection
confirmed two equal, aligned charts, four readable cards, complete axes and
legends, preserved branding/footer and pagination 9 / 10. The final Simulation
Summary remains slide 10.
Its slide XML and numerical speaker-note XML are byte-identical to the previously
accepted modern Diesel presentation (`summary_slide_preservation.json`).

| Measurement | Exact source | Unit | Minimum | Maximum | Time interval |
|---|---|---|---:|---:|---|
| Oil | Engine_Oil_Temperature | °C | 49.3053 | 79.9948 | 0–222 s |
| Water | Engine_Water_Temperature | °C | 49.3053 | 79.9948 | 0–222 s |

The real traces are identical; both graphs remain present. Controlled tests use
deliberately different arrays and verify that each renderer receives its own
source, common limits and a different colour. Further cases cover identical
traces, zero traces, either missing measurement and both missing measurements.

Visual evidence: `outputs/final_reporting_2026_10_16_validated/visual_checks/`
contains `diesel_slide_9.png`, `diesel.pdf` and native KPI table captures.

## E. Excel formula and independent numerical validation

Every applicable configured KPI Value cell contains a genuine formula; no
applicable KPI remains a static final result or a reference to a static copy.

| Dataset | Configured KPIs | Applicable / formulas | Unavailable | Maximum KPI difference after native recalculation |
|---|---:|---:|---:|---:|
| Diesel, latest supplied CSV | 5 | 5 / 5 | 0 | 0 |
| Electric, updated UoL `_04` CSV | 12 | 12 / 12 | 0 | 7.11e-15 |
| Hybrid, compatible older real CSV | 12 | 10 / 10 | 2 | 0 |
| Hybrid, controlled updated fixture | 12 | 12 / 12 | 0 | 0 |

The older Hybrid reference lacks the HV auxiliary measurement and road height.
Accordingly `total_energy_consumption_tyres_aux` and `road_height_range` are
unavailable; they are not represented as zero. The updated fixture is synthetic.

All 210,883 formula caches were cleared before desktop Excel performed
CalculateFullRebuild and per-sheet Calculate on disposable copies. Comparison of
1,907,392 numeric cells found no Excel errors, missing required caches or numeric
mismatches (rtol 2e-14, atol 1e-12). Maximum absolute differences across all numeric
cells were 4.55e-13 Diesel, 4.26e-14 Electric, 1.42e-13 older Hybrid and 3.79e-19
controlled Hybrid. This exercises the complete dependency chains and detects
broken references/circular calculations through independent results.
An explicit native Excel CircularReference check also returned no circular
references in all nine current, legacy and boundary-case validation workbooks
(`formula_cycle_check.json`).

Additional raw-edit checks doubled Diesel engine speed and multiplied battery
SOC by 0.9 in each of the other three disposable workbooks. All 39 applicable
KPIs recalculated to the independently expected unchanged/changed results.
No original report was saved after these edits.

A further Diesel fixture uses times [17,18,117,1817,3617], different oil/water
traces and variable engine speed/fuel flow. Its independently recalculated
time-weighted means are 1182 rpm and 11.86138888888889 L/h, with 60 min elapsed.
All 947 numeric cells agree (maximum difference 3.89e-16). This explicitly
distinguishes time weighting from a sample arithmetic mean and checks a shifted
time origin. Regressions also cover 3/17-sample datasets and exports containing
only Time while retaining a complete live battery-range dependency chain.

Additional independent native Excel checks of all three explicit legacy
profiles passed for all 29 available KPIs (5 Diesel, 12 Electric, 12 Hybrid),
with zero KPI differences. All 8,201 numeric cells agreed; the largest difference
was 9.10e-13. These controlled checks exercise the original sample-rectangle
conventions and formula helpers over MATH channels whose visible legacy values
remain static. Evidence is in `legacy_audit.json`.

A separate current Electric boundary fixture includes zero SOC and a positive
SOC below the established validity tolerance. Native Excel excludes both from
nominal-capacity inference, agrees on 50 kWh and matches all 12 KPI results.
All 3,195 numeric cells agree (maximum difference 3.80e-19). This deliberately
synthetic boundary case is recorded in `soc_edge_audit.json`.

PowerPoint uses the unchanged deterministic Python metrics. All configured
metrics exposed to its adapter match the canonical values and units. The updated
Electric auxiliary chart remains available. Existing narrow identifier and
dependency columns retain their previous widths; values/units are readable and
long formulas remain inside cells without changing the visible layout.

Machine-readable evidence in the validation directory:
`audit.json`, `kpi_comparison.json`, `raw_edit_checks.json`,
`irregular_audit.json`, `legacy_audit.json`, `soc_edge_audit.json`,
`formula_cycle_check.json`, and `baseline_preservation.json`.

## F. Regression commands and results

The complete suite runs under Python 3.11.15 and supported Streamlit 1.49.1,
with every declared dependency constraint checked by the runner. The local
default Streamlit 1.61.1 is deliberately excluded from this validation.

```powershell
.venv\Scripts\python.exe -u outputs/run_final_reporting_tests_clean.py
```

Exact subprocess pytest commands, module lists, dependency versions and logs are
recorded in `outputs/final_reporting_regression_clean/manifest.json`.

| Complete-suite group | Passed | Skipped | Failed |
|---|---:|---:|---:|
| 1 | 166 | 0 | 0 |
| 2 | 233 | 1 | 0 |
| 3 | 180 | 0 | 0 |
| Total, all 37 modules | **579** | **1** | **0** |

The sole expected skip is the existing Diesel soil-slide case: those soil slides
apply only to Electric and Hybrid. There are 22 new regression cases, increasing
the previous 557-pass baseline to 579. Initial inventory assertions expecting
four workbook sheets were updated to explicitly include the new hidden Formula
Support sheet; existing visible-sheet assertions remain unchanged.

Following the final clarification of missing-temperature subtitle text, the six
thermal cases were rerun: **6 passed, 16 deselected, 0 failures**. Exact command:

```powershell
@'
from pathlib import Path
import sys
sys.path.insert(0,str(Path('outputs/uol_streamlit_149').resolve()))
import pytest
sys.exit(pytest.main(['tests/test_final_reporting_improvements.py','-k','thermal_charts','-p','no:cacheprovider','-o','addopts=','-q','--basetemp','outputs/final_thermal_text']))
'@ | .venv\Scripts\python.exe -u - > outputs/final_reporting_thermal_final.txt 2>&1
```

`git diff --check` passed. No tests were removed or engineering tolerances relaxed.

Independent Office validation commands:

```powershell
.venv\Scripts\python.exe -u outputs/validate_final_reporting.py
.venv\Scripts\python.exe -u outputs/verify_electric_office.py prepare outputs/final_reporting_2026_10_16_validated/audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/validate_excel_recalculation.ps1 -AuditPath outputs/final_reporting_2026_10_16_validated/audit.json
.venv\Scripts\python.exe -u outputs/verify_electric_office.py audit outputs/final_reporting_2026_10_16_validated/audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_sergio_2026_10_reports.ps1 -AuditPath outputs/final_reporting_2026_10_16_validated/audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File outputs/final_reporting_office_qc.ps1
.venv\Scripts\python.exe -u outputs/validate_irregular_diesel_kpis.py
.venv\Scripts\python.exe -u outputs/verify_electric_office.py prepare outputs/final_reporting_2026_10_16_validated/irregular_audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/validate_excel_recalculation.ps1 -AuditPath outputs/final_reporting_2026_10_16_validated/irregular_audit.json
.venv\Scripts\python.exe -u outputs/verify_electric_office.py audit outputs/final_reporting_2026_10_16_validated/irregular_audit.json
```

## G. Generated deliverables

Root: `outputs/final_reporting_2026_10_16_validated/`.
`DELIVERABLES.md` in that directory provides direct links to each workbook and
presentation.

Each subdirectory below contains `profile_excel_report/<source stem>.xlsx`
and `profile_powerpoint_report/<source stem>.pptx`:

- `diesel/`: Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient
- `electric/`: RoboSprayer_3500Kg_Electric_UoL_DriveCycle_12kph_DriverPC_20pcGrad_04
- `hybrid_reference/`: Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem
- `hybrid_controlled/`: hybrid_controlled
- `diesel_controlled/`: diesel_controlled

Cache-cleared inputs, independently recalculated workbooks and raw-edit copies
are separated in `cache_cleared/`, `recalculated/` and `mutations/`. Earlier
validation deliverables remain preserved.

## H. Remaining limitations

There is no updated real Hybrid CSV containing all newly requested measurements.
Full updated-Hybrid real-data acceptance remains pending that dataset. The older
Hybrid reference and controlled fixture provide explicitly different evidence.
The existing Electric battery-capacity consistency warning remains unchanged
(median approximately 98.2263 kWh); no battery calculation was adjusted.

## I. Git status

No commit or push was performed. The ten Python/configuration/test files listed
in C and this report are the intended changes; generated reports and validation
scripts/logs remain ignored beneath `outputs/`. No dependency files or
engineering-calculation modules were changed.
There are eight modified tracked files and three new untracked files (this
report, the KPI formula compiler and the new regression module). The branch
remains `main` at `7be9bf247b7ce507ffc0c360e26fbeef74eef17d`.
