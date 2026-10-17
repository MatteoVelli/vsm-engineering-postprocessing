# Soil PowerPoint KPI enhancement — 16 October 2026

The current Electric and Hybrid Tyre-Soil Bulldozing and Tyre-Soil Compaction
pages each contain six engineering KPI cards, using the Wheel Steering Angles
card renderer. Both original plots remain on each page, below the cards.
Titles, subtitles, header icons, branding, slide numbers and footer are retained.

## Calculation conventions and sources

No raw mapping, MATH formula, Excel layout or Excel calculation was changed.
The new metrics are presentation-only results with provenance in the profile
PowerPoint manifest. They read the same validated arrays used by Excel and plots.

| KPI | Unit | Source and calculation |
| --- | --- | --- |
| MAX POWER | kW | Shared `compute_statistic(P, "max")`; signed maximum, not maximum absolute magnitude |
| MEAN POWER | kW | Existing `_time_average(P, track_time)`: trapezoidal time integral divided by observed elapsed duration |
| RMS POWER | kW | Shared `compute_statistic(P, "rms")`: `sqrt(mean(P²))`, using all samples, matching existing sample RMS conventions |
| TOTAL ENERGY | kWh | `last(E) - first(E)`, using the validated accumulated channel directly |
| ENERGY PER KM | kWh/km | Net energy divided by existing canonical `distance_km_last`, only for a finite positive distance |
| DISTANCE | km | Existing canonical `distance_km_last`; no position/displacement substitute or conflicting distance calculation |

For bulldozing, `P` is `total_bulldozing_resistance_power` and `E` is
`tyre_bulldozing_resistance_energy_accumulated`. For compaction, the channels are
`total_compaction_resistance_power` and
`tyre_compaction_resistance_energy_accumulated`. Each power channel is the
existing sum of four measured `Tyre_BulldozingResistancePower_{FL,FR,RL,RR}` or
`Tyre_CompactionResistancePower_{FL,FR,RL,RR}` sources in kW. Time is `Track_Time [s]`.
Distance is the existing `distance_km`/`distance_km_last` definition derived from
`Track_Distance [m]`. Its established Excel convention is preserved.

Signed values are retained. An offset in accumulated energy is removed by the
endpoint difference; accumulated energy is never integrated again. Power, energy
and time inputs must have the report's sample count, valid units and finite
values. No low-power sample is filtered out. In particular, the real Electric
low-power region near 2,400 s remains in both the mean and RMS calculations.

Missing/invalid measurements show N/A, with a reason in the PowerPoint manifest.
Zero distance remains a real 0 km measurement, but its energy-per-km result is
N/A. Negative/nonfinite distance is unavailable. Metrics respect presentation
selection; the canonical distance must be available to supply distance/ratio cards.
No unrelated measurement or zero is substituted for a missing dependency.

## Modified files in this task

- `src/vsm_postprocessing/profile_soil_kpis.py`: reusable presentation-only
  metrics and six-card definitions, using shared numerical helpers.
- `src/vsm_postprocessing/profile_powerpoint_report_engine.py`: extracts the
  metrics, reuses the steering card renderer, positions the original charts below
  the row and records source/calculation/availability metadata.
- `src/vsm_postprocessing/report_profile.py`: permits six configured statistics
  on additional plot-pair pages, retaining support for pages without statistics.
- `config/report_profiles/sergio_2026_10/robosprayer_electric.yaml` and
  `robosprayer_hybrid.yaml`: six presentation metric IDs per soil page.
- `tests/test_soil_powerpoint_kpis.py`: 23 focused numerical, missing-data,
  selection, unit and native slide-content/geometry checks.
- README, CHANGELOG, client quick start, configuration guide and release notes;
  this validation report.

The existing shared renderer retains its original behavior for all existing card
units. The new longer `kWh/km` value uses a wider inset text area and native
single-line fitting, retaining the same nominal 19 pt Calibri value font. Native
PowerPoint inspection checks actual text bounds, not only XML geometry.

## Output and preservation checks

The real Electric `_04`, available older real Hybrid reference, controlled Hybrid
generator-active/inactive inputs and real Diesel `_07` were regenerated.
Electric remains 23 Excel charts/14 PowerPoint slides; the older Hybrid reference
has 25/14; complete controlled Hybrid has 29/14; Diesel remains 20/10.

Across all five cases, every Excel package part is identical to the prior release
apart from generation timestamps and independently verified profile SHA-256
provenance cells. All data, formulas, comments, statistics, KPI values, row heights,
styles, chart XML and drawings are preserved. All 164 available calculated-channel
array hashes are identical to their pre-enhancement values.

For Electric and Hybrid, the only changed PowerPoint package parts are the XML
for the two soil pages. Every other slide, layout, theme, relationship and media
part is byte-identical. All four original plot images, including axes and source
data, are byte-identical. Diesel has no changed PowerPoint parts.

Desktop PowerPoint opened every deck, rendered each enhanced page beside the
Wheel Steering reference and checked exactly six labels and six numerical values
per soil page. Native text-bound checks passed for all 48 values across four
Electric/Hybrid cases. Visual inspection confirmed readable axes, legends, title,
units, card geometry, colours and fonts, with no clipping or overlap.

## Numerical results

The tables below retain nine decimal places for engineering review. Slide values
use the established two-decimal presentation precision. Each table contains the
twelve values for one validated input. Full precision and source/calculation
metadata are in `outputs/soil_kpis_2026_10_16/kpi_values.csv` and each
`final_reports/*_soil_kpis.json` file.

### electric

| KPI | Unit | Bulldozing | Compaction |
| --- | --- | ---: | ---: |
| MAX POWER | kW | 6.767643000 | 10.225460000 |
| MEAN POWER | kW | 4.920470899 | 7.354055842 |
| RMS POWER | kW | 5.133176408 | 7.689040600 |
| TOTAL ENERGY | kWh | 5.402950407 | 8.075161874 |
| ENERGY PER KM | kWh/km | 0.503110168 | 0.751940281 |
| DISTANCE | km | 10.739100000 | 10.739100000 |

### hybrid_reference

| KPI | Unit | Bulldozing | Compaction |
| --- | --- | ---: | ---: |
| MAX POWER | kW | 7.935538000 | 16.803820000 |
| MEAN POWER | kW | 3.201192311 | 6.776846398 |
| RMS POWER | kW | 3.630715801 | 7.686128295 |
| TOTAL ENERGY | kWh | 1.573030333 | 3.330067022 |
| ENERGY PER KM | kWh/km | 0.131085861 | 0.277505585 |
| DISTANCE | km | 12.000000000 | 12.000000000 |

### hybrid_active

| KPI | Unit | Bulldozing | Compaction |
| --- | --- | ---: | ---: |
| MAX POWER | kW | 4.000000000 | 4.000000000 |
| MEAN POWER | kW | 4.000000000 | 4.000000000 |
| RMS POWER | kW | 4.000000000 | 4.000000000 |
| TOTAL ENERGY | kWh | 4.000000000 | 4.000000000 |
| ENERGY PER KM | kWh/km | 13.333333333 | 13.333333333 |
| DISTANCE | km | 0.300000000 | 0.300000000 |

### hybrid_inactive

| KPI | Unit | Bulldozing | Compaction |
| --- | --- | ---: | ---: |
| MAX POWER | kW | 4.000000000 | 4.000000000 |
| MEAN POWER | kW | 4.000000000 | 4.000000000 |
| RMS POWER | kW | 4.000000000 | 4.000000000 |
| TOTAL ENERGY | kWh | 4.000000000 | 4.000000000 |
| ENERGY PER KM | kWh/km | 13.333333333 | 13.333333333 |
| DISTANCE | km | 0.300000000 | 0.300000000 |


## Tests, delivery and remaining limitations

The preceding validated baseline was 606 passed/one expected skip. This task's
focused checks passed 50 tests: 23 new KPI tests plus 27 existing soil-plot tests.
Two existing legacy compatibility checks also pass unchanged after adding an
early return for profiles without soil cards. The complete suite passed **629
tests with zero failures and one expected Diesel soil-plot skip**, across 39 test
modules. Python 3.11.15, Streamlit 1.49.1 and all declared dependency constraints
were enforced. Full commands/results are in
`outputs/soil_kpis_2026_10_16/regression_accepted/manifest.json`.

The latest updated real Hybrid export is unavailable. Hybrid acceptance uses the
available older real dataset and controlled complete active/inactive datasets.
The older dataset's unavailable newer optional channels remain unavailable. All
twelve requested soil KPIs and all four soil plots are supported for that source.

The previous validated client folder and ZIP stay in place. The rebuilt portable
folder and ZIP use `Astauto_VSM_October_16102026` inside a separate
`dist/Soil_KPI_Enhancement_2026-10-16/` delivery directory. Runtime requirements,
setup/start scripts and app version 1.3.0 are unchanged. No commit or push is made.

Fresh extraction outside the repository passed the unchanged setup with developer
Python removed from PATH. It created an independent CPython 3.11.9 runtime and
virtual environment; only cached official runtime downloads were reused.
`pip check` passed and doctor reported zero failures, with its expected warning
about the excluded private example dataset. All five report pairs were regenerated
by that installed copy and passed the independent twelve-KPI calculations and
source workbook parity checks. Desktop PowerPoint also opened/rendered those
installed-copy decks and passed all six-card/native text-bound checks.

The working tree began this task with 21 modified and six untracked files. Those
pending changes are preserved. It now contains 21 modified and nine untracked
files, including the new metric module, tests and this report. Branch remains
`main`; HEAD remains `7be9bf247b7ce507ffc0c360e26fbeef74eef17d`. Full status is in
`outputs/soil_kpis_2026_10_16/git_status.txt`. No commit, tag or push was performed.

The shipped root launcher passed real Diesel, Electric and available Hybrid
uploads, validation, report generation and download-button checks in isolated
headless Edge, from an unrelated working directory with spaces and without the
developer environment. There were zero Streamlit exceptions. Inspection of the
UI-generated presentations confirmed six numerical cards on each soil slide and
all four original plots. Owned browser/server processes were stopped; QA ports
8578 and 9278 were closed.

Accepted delivery:

- Folder: `dist/Soil_KPI_Enhancement_2026-10-16/Astauto_VSM_October_16102026/`
- ZIP: `dist/Soil_KPI_Enhancement_2026-10-16/Astauto_VSM_October_16102026.zip`
- Checksum: `dist/Soil_KPI_Enhancement_2026-10-16/Astauto_VSM_October_16102026.sha256`
- Size: 94,435,278 bytes; 116 files including the manifest.
- SHA-256: `e53ea2cb1a233fd10819913039d1123effe3672142de18704f68157d5e558db3`

The clean folder matches the ZIP byte-for-byte. Every manifest hash and shipped
source byte matches the reviewed working tree. The package excludes raw datasets,
generated reports, caches, environments and development tests. Its two canonical
builds were reproducible. The prior delivery remains unchanged at its original
location, with SHA-256
`a312662e7a5dce5f58f414f06ef3588ccf126bf54f747cbcbad892e4078c1ad6`.

Acceptance evidence under `outputs/soil_kpis_2026_10_16/` includes:

- `regression_accepted/manifest.json`, `focused_accepted.txt` and `legacy_compatibility.txt`;
- `preservation.json`, `signals_before.json`, `signals_after.json`;
- `final_reports/audit.json`, `final_reports/*_soil_kpis.json`, native PNG/PDF renders and `kpi_values.csv`;
- `fresh_installation.json`, setup/doctor/pip logs and `package_audit.json`;
- `native_powerpoint.json` and installed-copy renders under `visual_checks/`;
- `browser_validation.json`, `ui_output_audit.json`, `previous_delivery.json` and `delivery.json`.
