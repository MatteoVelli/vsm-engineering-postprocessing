# Confirmed HV auxiliary unit: engineering assessment, 16 October 2026

## A. Root cause and confirmed interpretation

The importer correctly preserves the VSM channel's declared `-` unit. The current
Electric/Hybrid profiles expect `kW`, so strict profile resolution previously
classified the optional source as unavailable. The resulting dependency closure
excluded total auxiliary power, interval auxiliary energy and accumulated energy.
This was appropriate before the client's physical-unit clarification.

Sergio, Astauto, confirmed on 16 October 2026 that
`User_defined_Data_High_Voltage_Auxiliary_signal1_0` measures power in **kW** and
that VSM cannot attach its physical unit to this User Defined Channel. Values are
already kW: factor 1, offset 0. The correction explicitly records this authority.
It does not treat other `-` channels as power and does not convert measurements.

The baseline is captured in `outputs/hv_auxiliary_2026_10_16/baseline.json`:
the original warning, unavailable MATH channels, slide 7 configuration and
dependencies, and the previous presentation's actual placeholder text.

## B. Solution evaluation

| Approach | Decision |
| --- | --- |
| Globally accept `-` as kW | Reject: changes the meaning of unrelated dimensionless/unknown measurements. |
| Rewrite imported units | Reject: loses the original source metadata and changes other workflows. |
| Add scaling or W-to-kW conversion | Reject: source numbers already represent kW. |
| Hard-code the source in each engine | Reject: duplicates policy across validation, statistics, plots and reports. |
| Explicit raw-channel unit declaration in the profile resolver | Selected: one auditable policy, scoped to an exact channel and declared source unit. |

Raw definitions support an optional `unit_interpretation` containing
`source_unit`, `engineering_unit` and mandatory `authority`. The engineering unit
must match the definition's expected unit. Unknown keys, including scaling and
conversion fields, fail configuration validation. Only the current Electric and
Hybrid profiles configure it, for the exact confirmed signal with source `-`.

Resolution applies the declaration only to the exact source name and exact source
unit. Aliases and normalized source-name matches do not activate it. Explicit kW
continues through ordinary unit validation. W, N, missing/blank units and unrelated
channels are not accepted by this declaration. Missing or non-finite optional
measurements retain the existing unavailable-data behavior.

`ResolvedChannel.channel` still points to the imported metadata; its `unit` stays
`-`. `effective_unit` is kW and `engineering_channel` supplies a separate metadata
view for plots/statistics. Raw arrays are unchanged. Applied interpretations are
recorded in diagnostics, workbook metadata, visible Excel mapping notes, plot
channel provenance, and structured manifest records containing source identity,
both units, authority and `values_unchanged: true`.

## C. Files changed for this correction

- `src/vsm_postprocessing/report_profile.py`: strict declaration parsing, exact
  identity/unit application, separate effective metadata and provenance.
- `src/vsm_postprocessing/profile_plotting.py`: consume the engineering metadata
  view and retain interpretation provenance.
- `src/vsm_postprocessing/profile_statistics.py`: consume effective raw units.
- `src/vsm_postprocessing/excel_report_engine.py`: visible source-unit notes in
  the mapping sheet and structured interpretation records in the manifest.
- `config/report_profiles/sergio_2026_10/robosprayer_electric.yaml` and
  `robosprayer_hybrid.yaml`: the confirmed declaration and client authority.
- `tests/test_hv_unit_interpretation.py`: 23 focused regression cases.
- `scripts/validate_electric_uol.py`: independently audit the restored auxiliary
  chain, provenance, formulas and slide availability in the actual UI workflow.
- `docs/CONFIGURATION_GUIDE.md`: declaration syntax, scope and traceability.
- This engineering assessment.

No changes were made to the importer, numerical integration, battery/tyre/torque
definitions, report templates, slide layout engine, dependency bounds or the
previous selection correction. Unrelated local modifications were preserved.

## D. Independent numerical verification

The actual Electric CSV has 608 source channels and 3,954 samples. Its SHA-256 is
`773e632fcabad1759441de495167fd43b088c3819c3851533419637a84950681`.
The validation script parses its names, units and numbers directly with Python's
CSV reader, independently of the production importer and expression evaluator.

| Quantity | Independent result | Comparison with application |
| --- | --- | --- |
| HV auxiliary power | 0 to 10 kW; last 10 kW | All 3,954 samples identical; no scaling |
| LV auxiliary power | 2.15054 kW | Actual source contribution retained |
| Total auxiliary power, HV + LV | 2.15054 to 12.15054 kW; last 12.15054 kW | All samples identical |
| Interval auxiliary energy | First 0; maximum and last 0.00337515 kWh | All samples identical |
| Accumulated auxiliary energy | Last 12.465459322222488 kWh | Independently accumulated curve identical |
| Separately summed trapezoidal integral | 12.465459322222221 kWh | Difference 2.67e-13 kWh from summation order |

The separate integrals are 10.104046927777777 kWh HV and
2.3614123944444443 kWh LV. Their sum equals the total trapezoidal integral within
floating-point tolerance, confirming that HV is included exactly once.

For seconds and kW, interval energy is
`(P[i-1] + P[i]) * (t[i] - t[i-1]) / 7200` kWh, with zero initial increment;
accumulated energy is its cumulative sum. No integration expression was changed.

Electric and Hybrid fixtures also exercise negative power and irregular times:
`t = [0,2,5,9] s`, total power `[10,-3,-7,12] kW`, interval energy
`[0,7,-15,10]/3600 kWh`, accumulated energy `[0,7,-8,2]/3600 kWh`.
The LAST statistic must be `2/3600`, not the earlier MAX `7/3600`.
These cases detect clamping, extrapolation, unintended scaling and duplicate HV
addition. Explicit kW, unknown/incompatible units, alias isolation, missing data
and malformed declarations are also covered.

The four raw gearbox torques, signed wheel powers, wheel totals and all three
resistance power/energy chains were independently checked against the real CSV
again. Their extrema/final values are unchanged from the previous accepted audit;
the total tyre resistance energy remains 28.0835308011 kWh. Both RMS statistics
and 26 other available statistics were independently checked where their raw or
independently derived arrays were available.

## E. Excel, PowerPoint and slide 7

Dedicated outputs are under
`outputs/hv_auxiliary_2026_10_16/electric_supported_149/`; previous deliverables
were not overwritten. The real Streamlit workflow ran under declared supported
Streamlit **1.49.1**, supplying only its file-upload widget through AppTest's
upload substitute. Actual selections, validation and both report generators ran.

The workbook now contains all **330 selected channels**: 294 raw and 36 MATH,
with 35 available statistics, 12 KPIs and 19 Excel plots. The restored outputs are
the raw HV column, total auxiliary power, interval/accumulated auxiliary energy,
two auxiliary statistics, the tyre-plus-auxiliary energy KPI and the auxiliary
energy chart. All **1,304,820 numeric data cells** were compared with the
calculation arrays; **142,344 live data formulas** were verified. Mapping notes
retain the source `-`, effective kW, unchanged values and client authority.

Slide 7 is **restored**. Its second plot is
`auxiliaries_energy_consumption`, which requires `time_minutes`,
`total_auxiliary_power` and `auxiliary_energy_consumption_accumulated`.
The latter two depend on the confirmed HV source and the LV source. Its first
plot, `battery_energy_distance_based`, depends on distance, speed and battery
energy and was independent of the HV failure. Resolving the unit automatically
restored the existing chart and **12.47 kWh AUX ENERGY card**; no placeholder or
artificial chart was inserted. There are no plot-unavailable placeholders anywhere
in the generated 12-slide Electric deck.

The existing template-fidelity check passes for the 13.333-by-7.5-inch deck,
10 text slots, 67 KPI-card slots and 13 plot slots. Desktop PowerPoint rendered
the complete deck to PDF. Slide 7 was also rendered to PNG and visually checked:
both charts, dual-axis units, six KPI cards, branding, title, footer and card/chart
layout are intact. The validation also rendered slide 9.

Desktop Excel recalculation uses a disposable workbook copy. All **143,739 formula
caches** were removed and verified empty before `CalculateFullRebuild` and
per-sheet calculation. All **1,306,686 numeric cells** across the worksheets agree
within **4.55e-13**, and no Excel error cells were found. The final comparison is
recorded in `electric_supported_149/office_audit.json`. The original report stays
unchanged.

## F. Regression commands and evidence

Focused tests (23 cases):

```powershell
.venv\Scripts\python.exe -m pytest -p no:cacheprovider -o addopts= -q `
  --basetemp outputs/hv_focus_tmp tests/test_hv_unit_interpretation.py
```

Supported-runtime compatibility checks (31 cases, all passed):

```powershell
.venv\Scripts\python.exe -c "import sys; sys.path.insert(0,'outputs/uol_streamlit_149'); import streamlit; print('Streamlit',streamlit.__version__); import pytest; raise SystemExit(pytest.main(['-p','no:cacheprovider','-o','addopts=','-q','--basetemp','outputs/hv149_tmp','tests/test_hv_unit_interpretation.py','tests/test_profile_selection_roundtrip.py','tests/test_sergio_october_15.py::test_ui_modern_profile_mapping_and_selections','tests/test_diesel_finalization.py::test_actual_app_generates_both_reports_and_respects_machine_name']))"
```

This includes all new unit cases, the previous Electric/Hybrid/Diesel selection
round-trip regressions and actual Diesel report generation. The isolated runtime
and its existing declared dependencies are installed only under ignored outputs;
no dependency was added to production requirements or the active environment.

Complete suite invocation:

```powershell
.venv\Scripts\python.exe outputs/run_hv_verified_suite.py
```

This runs every one of the 36 test modules in three isolated pytest processes,
using shorter workspace temporary paths. Exact per-group pytest commands, file
coverage, exit codes, summaries and logs are retained in
`outputs/hv_suite/manifest.json`. **554 passed, one expected skip, zero failures**:
group 1 had 184 passes and one skip, group 2 had 189 passes, and group 3 had 181
passes. This includes every case in the previous 531-pass baseline plus the 23
new unit cases. No tests were weakened. The sole expected skip is the pre-existing
unsupported Diesel soil-slide case. Compilation and `git diff --check` also pass.

Real-file supported-runtime invocation (actual execution was wrapped with
`subprocess.run` to preserve UTF-8 logs and the Python exit code):

```powershell
.venv\Scripts\python.exe outputs/run_supported_electric_uol.py `
  '<path-to-private-electric-vsm.csv>' `
  --output-dir outputs/hv_auxiliary_2026_10_16/electric_supported_149
```

For a new run, use a fresh output directory; the script refuses existing evidence
directories and requires the private source path in place of the placeholder.
The wrapper only prepends the isolated 1.49.1 installation before
running `scripts/validate_electric_uol.py`. With a correctly configured client
environment, run that script directly with the same arguments.

The `outputs/` helpers and evidence paths above describe the historical local
validation session; they are intentionally excluded from Git. On a fresh checkout
with supported dependencies, run the full suite directly with
`python -m pytest -p no:cacheprovider -o addopts= -q --basetemp outputs/tests`,
and run `scripts/validate_electric_uol.py` for the private Electric source. See
[the milestone audit](OCTOBER_2026_MILESTONE_AUDIT.md) for the final Git-validation
results and release boundary.

Native Excel evidence commands:

```powershell
.venv\Scripts\python.exe outputs/verify_electric_office.py prepare outputs/hv_auxiliary_2026_10_16/electric_supported_149/office_audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/validate_excel_recalculation.ps1 -AuditPath outputs/hv_auxiliary_2026_10_16/electric_supported_149/office_audit.json
.venv\Scripts\python.exe outputs/verify_electric_office.py audit outputs/hv_auxiliary_2026_10_16/electric_supported_149/office_audit.json
```

Native PowerPoint evidence commands:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/render_sergio_2026_10_reports.ps1 -AuditPath outputs/hv_auxiliary_2026_10_16/electric_supported_149/office_audit.json
powershell -NoProfile -ExecutionPolicy Bypass -File outputs/hv_auxiliary_2026_10_16/render_slide_7.ps1 -AuditPath outputs/hv_auxiliary_2026_10_16/electric_supported_149/office_audit.json
```

## G. Remaining concerns and validation limits

There is no remaining uncertainty about the confirmed signal's physical unit.
The source metadata intentionally stays `-`; its explicitly declared engineering
interpretation is kW. A missing signal or a different source unit still needs
valid source data and will not be fabricated or silently reinterpreted.

The supplied Hybrid CSV lacks the HV signal and all four new gearbox-torque
signals. The supplied Hybrid Gen80kW XLSX is a reduced 70-column report without
the confirmed HV source. Consequently, Hybrid unit handling, signed integration,
statistics, Excel and PowerPoint compatibility were validated with controlled
fixtures, not a newly complete real Hybrid simulation. Older source behavior is
covered by the existing regression suite. A fully updated Hybrid raw export is
still required for real-data Hybrid engineering validation.

The existing battery Energy/SOC capacity-consistency warning remains unchanged:
median 98.2263 kWh, maximum deviation 1.86791 kWh versus tolerance 0.982263 kWh.
Battery calculations and other unrelated KPIs were not adjusted. The active
workspace's Streamlit 1.61.1 remains outside the declared 1.49.x support range;
both the complete workspace regression suite and isolated supported-runtime
checks are reported separately.

## H. Git status and preservation

No commit, push or staging operation was performed. All pre-existing local
modifications were preserved. Section C lists the changes made for this task;
`outputs/hv_auxiliary_2026_10_16/git_status.txt` records the complete modified/new
file list, including earlier work. Audit logs, generated reports, native Office
copies, previews and isolated testing dependencies remain under ignored outputs.
