# Electric_05 / Hybrid_06 update

## Scope and logic

Compared the channel selection sheets and data formulas in Electric_05 versus
Electric_03, and Hybrid_06 versus Hybrid_04. The supplied workbooks are unchanged.
The existing uncommitted Diesel and presentation work was preserved. No commit
or push was performed.

- `DriveShaft_Torque_FL/FR/RL/RR` are required raw VSM inputs in both profiles.
  They have no MATH definition, placeholder, or EDU-derived fallback.
- Wheel Total Torque sums all four raw torques. Each wheel power uses its own
  wheel rotational speed and driveshaft torque divided by 9548.8. Wheel Power
  Total sums all four powers. Hybrid inherits exactly the same expressions.
- Added four MATH channels each for bulldozing, compaction, and total tyre
  resistance: power, sample kWh, sample Wh, and cumulative kWh. Each component
  power includes all four wheels. Total tyre power/energy sums rolling,
  compaction, and bulldozing. Existing rolling calculations are unchanged.
- Energy uses the existing `sample_energy_kwh` helper: power times the actual
  preceding sample interval, with the first interval repeated for the first
  sample, divided by 3600. Cumulative energy includes every sample. No change
  was made to the math engine's integration or non-finite-value policy.
- Replaced `tyre_rr_energy_accumulated_last` with
  `tyre_total_energy_accumulated_last`, labelled **Tyre Total Energy Consumption**.
  It uses the last accumulated total tyre energy, including when the series
  decreases. Tyres + Aux adds that statistic to the existing auxiliary energy
  consumption statistic. Excel, PowerPoint, and the existing tyre chart consume
  the new total. Report geometry, colours, fonts, and chart styles are preserved.
- Electric has **37 MATH channels**; Hybrid has **40**.
- Synced selection flags against every row in the latest selection sheets.
  All eight raw rolling resistance force/power flags are false; rolling total
  power and sample kWh flags are false. Electric wheel-load flags are false;
  Hybrid wheel-load flags remain true. A strict `channels.plot_overrides` mapping
  lets Hybrid retain those flags without duplicating common channel logic.
  Explicit existing report plots remain configured, including Wheel Loads.
- Road Height remains optional. Electric_05 omits it from its selection sheet;
  the application still includes and plots it when a source supplies it.
- Validation identifies required source/VSM names, optional source channels,
  and calculated MATH channels separately. It never asks for calculated MATH
  channels in a source CSV. Missing torque prevents dependent wheel calculations.
- Runtime Excel ordering and packaging now use/include Electric_05 and Hybrid_06.
  Previous workbooks remain available for historical comparisons.

## Workbook evidence and intentional differences

| Profile | Raw wheel torques | Total torque | Wheel powers | Total power |
| --- | --- | --- | --- | --- |
| Electric_05 | GX:HA | HB | HC:HF | HG |
| Hybrid_06 | HF:HI | HJ | HK:HN | HO |

Hybrid HO is `=HK5+HL5+HM5+HN5`; HQ is Wheel Steer Angle FR.
There is no HQ total-power dependency.

The new tests evaluate the actual workbook arithmetic at samples 0, 1, 17, 100,
and the final sample, then compare all wheel total samples and final cumulative
energy. Remaining differences are deliberate:

1. Both workbooks label Tyre Total Resistance Power as `kWh`. The application
   correctly reports `kW`.
2. Workbook tyre cumulative formulas use `SUM(energy_row:energy$6)`. Row 5
   includes row 6 prematurely, and subsequent rows omit row 5. The application
   retains correct cumulative integration including the first sample. Tests
   explicitly verify this reference defect and its first-sample energy offset.
3. For irregular timestamps, energy follows actual intervals; the workbook
   hardcodes one second. For one-second samples, sample energies agree.
4. Electric Road Height remains an optional application extension. Its presence
   shifts subsequent generated columns; absent optional channels also shift
   positions. Channel identities and calculations are semantic, not fixed-letter.
5. The Electric selection sheet calls its raw agrochemical force entry
   `Agrochemical Discharge`, but its data sheet still calculates mass from the
   force column. The proven `HitchRear_Force_Z_VehicleCoordinates` source mapping
   and force-to-mass conversion are retained. Matching uses the report label.
6. New energy labels disambiguate kWh/Wh and spell `Accumulated` correctly;
   the template reader accepts the workbook's `Accomulated` spelling and older
   Hybrid rolling-energy labels. Existing duplicate FR/RR rolling-power display
   labels are preserved; semantic source identities remain distinct.

## Reference CSV results

All three available Electric/Hybrid reference CSVs resolve all required inputs.
Their first, second, sample 100, and final energy values were inspected as well
as final totals. Synthetic tests independently exercise nonzero front and rear
wheel torques, because the two Electric CSVs contain zero driveshaft torques.

| Reference CSV | Samples | Rolling kWh | Bulldozing kWh | Compaction kWh | Total tyre kWh | Tyres + Aux kWh |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Electric `...Crop_Field_05.csv` | 3853 | 13.355223967067 | 5.004248797139 | 7.164000334780 | 25.523473098986 | 37.031848960098 |
| Electric `...Grad_Discharge.csv` | 347 | 1.240422740333 | 0.464302005166 | 0.670443904471 | 2.375168649969 | 3.411609455525 |
| Hybrid `...12x1Km_4000Kg_Chem.csv` | 1770 | 8.251955696373 | 1.573030481799 | 3.330067336628 | 13.155053514799 | 18.441797681466 |

Hybrid maximum total wheel power is **106.804341601039 kW**.

## Tests and validation

Added `tests/test_latest_template_math.py` with 38 parameterized cases covering:
raw torque classification, no synthesis, independent four-wheel torque/power and
tyre contributions, identical Electric/Hybrid logic, all new energy channels,
one-second and irregular timing, final-versus-maximum KPI behavior, Tyres + Aux,
each missing torque error, NaN rejection, optional Road Height, strict selection
overrides, full selection-sheet comparison, and actual workbook formulas.

Updated existing math, profile, statistics, Excel, UI, and packaging tests for
the new counts, classifications, labels, totals, and runtime assets. Generic
fallback tests now use a generic optional sensor rather than suggesting that
wheel torque may be synthesized.

Targeted new tests: **38 passed**.

Full suite command:

```powershell
.venv\Scripts\python.exe -u -m pytest -o addopts= -p no:cacheprovider --basetemp=.pytest_tmp/template_full_final
```

Full result: **388 passed in 472.70 seconds (7 minutes 52 seconds)**, with no
failures, skips, or warnings. This includes the existing Excel/PPT generation,
chart, axis, styling, Diesel, packaging, and UI tests. The complete output is
saved in `.pytest_tmp/template_full_final.log`. Pytest's cache plugin was
disabled because the existing `.pytest_cache` directory denies writes; all
tests remained enabled.
The initial full run had 386 passes and two outdated assertions (the old chart
anchor column and the removed torque fallback note); both assertions were
updated before the final run. Added Excel assertions also verify the exported
total tyre KPI label, its numerical value, and the Tyres + Aux value.

## Files changed for this task

- Profiles: `config/report_profiles/robosprayer_electric.yaml`,
  `config/report_profiles/robosprayer_hybrid.yaml`.
- Runtime: `src/vsm_postprocessing/report_profile.py`, `excel_report_engine.py`,
  `profile_plotting.py`, `profile_statistics.py`,
  `profile_powerpoint_report_engine.py`, `ui_config.py`, `ui_app.py`,
  `doctor.py`, `release_builder.py`.
- Tests: `tests/test_latest_template_math.py`, `test_profile_math.py`,
  `test_report_profile.py`, `test_profile_statistics.py`,
  `test_profile_excel_report.py`, `test_ui_config.py`, `test_release_packaging.py`.
- Documentation/configuration: `.gitignore`, `README.md`,
  `reference_files/README.md`, `docs/CONFIGURATION_GUIDE.md`,
  `docs/FINAL_ACCEPTANCE_CHECKLIST.md`, `docs/KNOWN_LIMITATIONS.md`, this report.
- Existing supplied assets newly included by Git/package rules:
  `reference_files/Robo_Sprayer_Electrification_Tamplate_Electric_05.xlsx`,
  `reference_files/Robo_Sprayer_Electrification_Tamplate_Hybrid_06.xlsx`.

Other files shown as modified/untracked by Git were already present before
this task, including the math-engine and Diesel implementation work.
