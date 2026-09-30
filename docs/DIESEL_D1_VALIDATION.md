# Diesel Phase D1: numerical configuration and validation

> Historical stage record. D1, D2 and D3 use the same Test_03 CSV; they are
> not different datasets. Excel and PowerPoint are now implemented. The current
> production behavior, final audit and release assessment supersede the phase
> boundaries below: [Diesel final validation](DIESEL_FINAL_VALIDATION.md).

## Baseline and scope

- Branch: `main`; HEAD: `0ec8529705a18bd88bd7524bdb816dc9a42efef1`.
- Initial modified files: `docs/SERGIO_REPORT_CHANGES.md`,
  `src/vsm_postprocessing/profile_powerpoint_report_engine.py`, and
  `tests/test_profile_powerpoint_report.py`. Initial untracked items:
  `Matteo_Velli_Manas_Studio_Invoice.pdf` and `logs/`. These were preserved.
- Unmodified baseline: **262 passed, 0 skipped, 0 failed, 489.73 seconds**.
  The first invocation used an inaccessible system pytest temporary directory
  (126 passed, 4 failed, 132 setup errors); rerunning with a fresh workspace
  `--basetemp` and `-p no:cacheprovider` resolved the environment issue without
  changing any code or assertions.
- Scope: Diesel profile, validation, deterministic math, statistics and KPIs.
  No Diesel report layouts, chart configuration or PowerPoint work was added.

## Architecture and integration decisions

Profiles are strict YAML definitions of semantic raw channels, expressions,
statistics, KPIs, plots and presentation settings. Resolution checks exact and
normalized source names, then explicit aliases, and verifies units. Selection
is explicit; measured channel activity does not select a powertrain.

Hybrid inherits the Electric profile. Caiman is a machine identity handled by
metadata and the existing Hybrid presentation configuration, not a separate
numerical profile. Diesel therefore does **not** inherit either existing
profile: it declares its own common vehicle channels and combustion channels.
The existing YAML discovery automatically adds `Full Size Sprayer Diesel`
after the existing Electric/Hybrid choices without changing their names/order.

Diesel-specific source validation runs at the profile math entry point, which
is also used by source validation and profile statistics. It rejects missing
required/ambiguous/unit-mismatched inputs, non-finite resolved channels, invalid
time, negative RPM/flow, and negative or resetting cumulative mass/distance
counters. Missing optional signals produce the existing unavailable-result
records. Present invalid optional data fails explicitly rather than being
silently discarded. `resolve_profile` remains a channel/units resolver;
numerical source invariants are enforced when analysis/validation runs.

The shared expression evaluator gained two additive functions:
`cumulative_trapezoid(values, time)` and `time_average(values, time)`.
Existing functions, including Hybrid's sampled-energy convention and 9548.8
power divisor, are unchanged. Existing statistics and KPI engines are reused
without modification. Diesel uses exact SI power conversion and integration
over elapsed time, with no fabricated interval before the first sample.

The report engines contain Electric/Hybrid layout assumptions. A Diesel-only
Excel entry guard and Streamlit early return prevent premature report generation
and stale report downloads. These guards implement the D1 boundary; they do not
implement Diesel reporting. The existing configurable legacy pipeline is intact.

## Authoritative source inspection

File: `reference_files/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.csv`

SHA-256: `e089fd94a9d6c6657982198c2a7106ac5cd36e0431cf5fb870c9565bb764d7b1`.
The source was read only and remains unchanged.

- Comma-separated CSV; channel names on row 1, units on row 2, numeric rows 3–225.
- 586 imported raw channels, 223 samples, no missing/invalid/non-finite cells.
- `Track_Time`: seconds, 0 through 222 inclusive, strictly increasing by 1 second.
  There are 222 elapsed intervals, not 223 seconds of duration.
- All **38** `ElectricSystem_*` channels are exactly zero. They are excluded
  from Diesel requirements/math/KPIs; no automatic powertrain inference occurs.
- Engine RPM, signed torque/load, throttle, fuel, gradient, accelerations,
  tyre normal loads and driveshaft torques are active.
- `Engine_Load` is **Nm**, not percentage; `Engine_Throttle` is percent.
- Engine oil and water temperature are both °C. The two traces are identical,
  decreasing from 79.9948°C to 49.3053°C. No thermal model or thermal KPI is
  inferred from these suspiciously identical source traces.
- `Track_Height` is absent. It is optional in Diesel; inspection also confirmed
  it was already optional in the existing Electric/Hybrid configuration.
- This CSV fails each existing profile with 34 missing mandatory channels,
  principally electric-controller/motor thermal and suspension channels. It
  passes strict CSV import. This verifies a requirements mismatch, not malformed
  CSV. Existing Electric/Hybrid validation has not been relaxed.

### Fuel semantics

| Source | Unit | First | Last | Range | Interpretation |
|---|---|---:|---:|---:|---|
| `Engine_FuelConsumption_absolut` | kg | 0 | 0.772836 | 0–0.772836 | Cumulative mass; no decreases, eight zero increments |
| `Engine_FuelConsumption_volumeflow` | l/h | 0.0116981 | 3.22648 | 0.0116981–41.4741 | Instantaneous sampled volumetric flow |
| `Engine_FuelConsumption_specific` | g/kWh | 216.960 | 238.410 | 203.004–240.627 | Source operating-point specific consumption |

All three are finite and nonnegative. The engine never reaches zero RPM in
this dataset (minimum 688.925 rpm); synthetic tests cover zero-speed power.
Positive fuel flow also occurs during negative shaft power, so a cycle BSFC
computed by naively averaging the specific-consumption channel would be unsafe.
As a cross-check only, integrating positive shaft power times source specific
consumption yields 0.779331 kg versus the 0.772836 kg source counter (about 0.84%
difference). Sampled power/flow integration is an estimate of the underlying
simulation, not an exact reconstruction of its internal fuel integrator.

The independently integrated flow is 0.942121 L. Its ratio to source mass
corresponds to about 0.820315 kg/L, which is only an observed consistency check:
**no density constant or mass/volume conversion is introduced**.

## Configuration inventory

Seven required raw channels:

`Track_Time [s]`, `Track_Distance [m]`, `Chassis_Speed [kph]`,
`Track_Gradient [%]`, `Engine_Speed [rpm]`, `Engine_Torque [Nm]`,
`Engine_FuelConsumption_absolut [kg]`.

Eighteen optional raw channels:

- `Track_Height [m]`.
- `Engine_FuelConsumption_volumeflow [l/h]`,
  `Engine_FuelConsumption_specific [g/kWh]`.
- `Engine_Load [Nm]`, `Engine_Throttle [%]`.
- `Engine_Oil_Temperature`, `Engine_Water_Temperature [°C]`.
- `Chassis_Acceleration_Longitudinal`, `_Lateral`, `_Vertical [m/s²]`.
- `Tyre_Fz_FL`, `_FR`, `_RL`, `_RR [N]` (total tyre vertical load, not the
  separate dynamic wheel-load perturbation).
- `DriveShaft_Torque_FL`, `_FR`, `_RL`, `_RR [Nm]`.

No speculative aliases or raw-channel fallbacks were added. Existing normalized
matching supports equivalent spacing/underscore spellings such as `Engine Speed`.

Three math channels:

1. `engine_power [kW] = Engine_Torque × Engine_Speed × 2π / 60000`.
   π uses its double-precision value, independently of the legacy Hybrid constant.
2. `engine_mechanical_energy [kWh] = cumulative_trapezoid(engine_power, time) / 3600`.
   This is **net signed** energy, with zero at the initial sample.
3. Optional `fuel_volume [L] = cumulative_trapezoid(fuel_flow, time) / 3600`.

The 52 configured statistics use existing RMS/MAX/MIN/first/last operations:
first/last time, distance and fuel counter; maximum vehicle speed; gradient
extrema; engine speed/torque/power MAX/MIN/RMS; final net energy and integrated
volume; maximum flow/load/throttle; optional height/temperature/tyre-load extrema;
acceleration RMS; driveshaft torque MIN/MAX/RMS. RMS remains sample-based.
The reference calculates 50 statistics; the two height statistics are explicitly
unavailable. Specific fuel consumption is retained as a raw optional signal only.

Five calculated KPIs complement the engine/vehicle maximum statistics:

- Drive cycle minutes: `(last time − first time) / 60`.
- Distance km: `(last distance − first distance) / 1000`.
- Fuel consumed kg: `last cumulative mass − first cumulative mass`.
- Average engine RPM: trapezoidal time-weighted average.
- Optional average fuel flow: trapezoidal time-weighted average in L/h.

## Numerical validation

| Metric | Reference result |
|---|---:|
| Duration | 222 s / 3.7 min |
| Distance | 999.242 m / 0.999242 km |
| Maximum vehicle speed | 22.4686 kph |
| Gradient range | −15% to +15% |
| Maximum engine speed | 1996.81 rpm |
| Average engine speed (time-weighted) | 1537.214164 rpm |
| Maximum engine torque | 826.219 Nm |
| Minimum engine torque | −71.9229 Nm |
| Maximum engine mechanical power | 165.938626340 kW |
| Minimum engine mechanical power | −8.230092194 kW |
| Engine power RMS | 74.570694136 kW |
| Net engine mechanical energy | 3.764781781 kWh |
| Total consumed source fuel mass | 0.772836 kg |
| Integrated sampled fuel volume | 0.942120719 L |
| Average fuel flow (time-weighted) | 15.277633289 L/h |
| Maximum fuel flow | 41.4741 L/h |
| Maximum engine load torque | 825.455 Nm |
| Maximum throttle | 77.1845% |

Independent checks used Python's CSV reader and scalar `math.tau`, separately
from the importer/profile evaluator. All 223 power samples were compared.

| Time / zero-based sample | Raw RPM | Raw torque (Nm) | RPM × torque × 2π / 60000 (kW) |
|---:|---:|---:|---:|
| 0 | 1300.730 | −0.328547 | −0.044752086 |
| 1 | 733.274 | 26.8464 | 2.061488711 |
| 50 | 1920.910 | 325.859 | 65.548892118 |
| **145** | **1917.890** | **826.219** | **165.938626340** |
| 151 | 1092.720 | −71.9229 | −8.230092194 |
| 222 | 756.779 | 141.771 | 11.235310858 |

For example, the independent maximum is
`826.219 × 1917.89 × 6.283185307179586 / 60000 = 165.93862633964403 kW`.
The maximum RPM occurs at a different sample; multiplying separate maxima would
be incorrect. Twelve samples have negative mechanical power. At 151 s the
gradient is −15%, front-left driveshaft torque is −378.742 Nm, engine load is
−21.3124 Nm and throttle is 4.32818%, consistent with overrun/engine braking.
The sign is retained, with no absolute-value operation. Synthetic tests verify
zero RPM, irregular sampling, shifted time origin and nonzero cumulative offsets.
The vehicle speed maximum also occurs at 151 s: 22.4686 kph against a 20 kph
driver target on the −15% descent, with longitudinal acceleration −1.03902 m/s².
This is an observed downhill overspeed/deceleration sample; the measured maximum
is not clamped to the nominal speed in the filename.

## Tests and reproducibility

Private-reference tests use `tests/conftest.py`'s repository-relative
`REFERENCE_FILES_DIR` and `VSM_TEST_REFERENCE_FILES_DIR` override. The exact
observed source filename is registered there. Missing private data is an
explicit skip; synthetic tests remain runnable from a clean checkout.

Example commands (use a fresh, unused basetemp for each invocation):

```powershell
.venv/Scripts/python.exe -m pytest tests/test_diesel_profile.py -o addopts='' -p no:cacheprovider --basetemp=outputs/diesel_targeted_tmp
.venv/Scripts/python.exe -m pytest -o addopts='' -p no:cacheprovider --basetemp=outputs/diesel_regression_tmp
```

The numerical API remains the existing profile interface:

```python
from vsm_postprocessing.importer import load_data_file
from vsm_postprocessing.report_profile import load_reporting_profile
from vsm_postprocessing.profile_statistics import calculate_profile_statistics

profile = load_reporting_profile("config/report_profiles/full_size_sprayer_diesel.yaml")
dataset = load_data_file(source_path)
result = calculate_profile_statistics(dataset, profile)
assert result.resolution.is_valid and result.math_result.is_complete and result.is_complete
# result.math_result.values_by_semantic_name, result.statistics, result.kpis
# and result.canonical_metrics carry values, units and source/formula traceability.
```

Targeted Diesel: **36 passed, 0 skipped, 0 failed, 7.68 seconds**.
Relevant profile regression run: **113 passed, 0 skipped, 0 failed, 98.09 seconds**
(35 Diesel cases plus 78 existing cases; the final targeted run includes the
additional UI boundary test). With a deliberately absent reference directory:
**33 passed, 3 explicitly skipped, 0 failed, 6.01 seconds**.
Final complete suite: **298 passed, 0 skipped, 0 failed, 468.92 seconds**.
All 262 baseline tests and 36 new Diesel tests pass. Existing Electric,
Hybrid/Caiman numerical, validation, selection, plotting, Excel and PowerPoint
behavior remains unchanged under the complete regression suite.

**Assessment: Phase D1 is complete and numerically trustworthy within the
documented source and sampling assumptions. Work stops here; D2 is not implemented.**

Local ignored audit artifacts: `outputs/diesel_d1_baseline.log`,
`outputs/diesel_d1_targeted.log`, `outputs/diesel_d1_profiles.log`,
`outputs/diesel_d1_portable.log`, `outputs/diesel_d1_full.log`,
`outputs/diesel_d1_source_inventory.json`, and
`outputs/diesel_d1_numerical_validation.json` (all 55 available canonical metrics
with units and provenance). These are numerical/test evidence, not Excel/PPT reports.

## Changed files and limitations

| File | D1 change |
|---|---|
| `config/report_profiles/full_size_sprayer_diesel.yaml` | Independent channel/units/requirements, math, statistic and KPI definitions |
| `src/vsm_postprocessing/diesel_validation.py` | Isolated Diesel source invariants |
| `src/vsm_postprocessing/math_engine.py` | Two additive deterministic integration/average functions |
| `src/vsm_postprocessing/profile_math.py` | Diesel-only validation dispatch |
| `src/vsm_postprocessing/ui_app.py` | Keep Diesel numerical validation available and defer report controls/downloads |
| `src/vsm_postprocessing/excel_report_engine.py` | Diesel-only pre-generation guard; no layout change |
| `tests/conftest.py` | Portable private Diesel fixture registration |
| `tests/test_diesel_profile.py` | Synthetic, reference, isolation, numerical and D1-boundary coverage |
| `reference_files/README.md` | Document the authoritative private Diesel reference |
| `docs/DIESEL_D1_VALIDATION.md` | This engineering evidence and reproducibility record |

Limitations: absent height, sampled integration rather than simulation-internal
integration, no resetting counters, and no omission/imputation of invalid
resolved data. At least two finite strictly increasing time samples are required.
Density-based conversion, differentiated mass flow, cycle BSFC, fuel per hectare
and distance-normalized consumption are intentionally unsupported in D1.
No plots or Diesel Excel/PPT layouts were configured.

For D2, use these verified canonical metrics to design a dedicated Diesel Excel
executive-results block and native editable engine/fuel/chassis charts. Specify
signed/net-energy labels and optional-height behavior, then add guarded workbook
regressions before enabling Diesel report generation. PowerPoint remains a later
phase. No commits or pushes were made.

## Git audit

The D1-only tracked diff is six files, **48 insertions, zero deletions**:

```text
reference_files/README.md                     |  6 ++++++
src/vsm_postprocessing/excel_report_engine.py |  2 ++
src/vsm_postprocessing/math_engine.py         | 26 ++++++++++++++++++++++++++
src/vsm_postprocessing/profile_math.py        |  5 +++++
src/vsm_postprocessing/ui_app.py              |  4 ++++
tests/conftest.py                             |  5 +++++
```

Four new D1 files are untracked and therefore absent from `git diff --stat`:
the Diesel YAML, validator, tests, and this document. The whole checkout's
tracked stat is nine files, 344 insertions and 156 deletions, including the
three pre-existing edits. SHA-256 comparisons confirm those three files,
both existing profile YAML files, and the Diesel reference are unchanged from
the start of this task. `git diff --check` passes.

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
