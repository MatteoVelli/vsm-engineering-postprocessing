# October soil-resistance plot update — 16 October 2026

The current Electric and Hybrid profiles now automatically include all four
requested quantities as separate time-based plots in Excel and PowerPoint.
No engineering calculations, source mappings, statistics, KPI formulas, runtime
dependencies or launchers were changed for this update.

## Root causes and channel mapping

All four MATH quantities already existed. The current profiles selected them
only as two combined, dual-axis plots against distance. Those existing plots
remain. The default PowerPoint configuration did not select either soil plot;
its template adapter supported the twelve existing pages only.

| Requested plot / existing semantic channel | Underlying source or existing calculation | Unit | Excel | PowerPoint |
| --- | --- | --- | --- | --- |
| Total Bulldozing Resistance Power / `total_bulldozing_resistance_power` | Sum of `Tyre_BulldozingResistancePower_FL`, `_FR`, `_RL`, `_RR` | kW | Separate chart appended | Page 12, left |
| Tyre Bulldozing Resistance Energy Accumulated / `tyre_bulldozing_resistance_energy_accumulated` | `cumulative_sum(tyre_bulldozing_resistance_energy_kwh)` | kWh | Separate chart appended | Page 12, right |
| Total Compaction Resistance Power / `total_compaction_resistance_power` | Sum of `Tyre_CompactionResistancePower_FL`, `_FR`, `_RL`, `_RR` | kW | Separate chart appended | Page 13, left |
| Tyre Compaction Resistance Energy Accumulated / `tyre_compaction_resistance_energy_accumulated` | `cumulative_sum(tyre_compaction_resistance_energy_kwh)` | kWh | Separate chart appended | Page 13, right |

All four plots use the original `Track_Time [s]` (`track_time`) samples. There
is no resampling, smoothing, time rebasing, interpolation or new integration.
Each interval-energy source already uses the current profile's signed
trapezoidal rule: initial increment zero, then
`(P[i-1]/2 + P[i]/2) * (t[i] - t[i-1]) / 3600`, followed by cumulative sum.
The accumulated channel is plotted directly and never integrated again.
Measured wheel powers are mapped in kW, with all four corners required for a
total. No force/speed proxy or sign clipping was introduced.

Both current profiles support these definitions when the required measurements
are available. Missing, nonnumeric, NaN, infinite or wrong-unit optional sources
disable their dependent pair. Excel records the unavailable plots/dependencies;
PowerPoint shows labelled unavailable-data notices. Valid measured zeros remain
real zero curves. Invalid timebases, empty series and misaligned series are
rejected by existing validation. Diesel and explicit legacy profiles retain
their original plot inventories and mathematical conventions.

## Implementation

- `config/report_profiles/sergio_2026_10/robosprayer_electric.yaml` and
  `robosprayer_hybrid.yaml`: four single-series plots, ordered 100–103 after
  existing plots, and two configured additional presentation pages.
- `src/vsm_postprocessing/report_profile.py`: validates explicitly appended
  presentation definitions without replacing the inherited presentation.
- `src/vsm_postprocessing/profile_powerpoint_report_engine.py`: renders the
  configured plot-pair additions before the final summary, using the approved
  Astauto background, fonts, caption, logo, footer and chart dimensions. Template
  chart geometry is captured before template shape replacement.
- `tests/test_october_soil_plots.py`: 27 focused engineering, inclusion,
  missing/invalid-data, empty/alignment, style, placement, selection and malformed
  configuration checks.
- `tests/test_sergio_october_15.py`: two existing current-profile slide-count
  expectations updated from 12 to 14; their original assertions remain.
- README, changelog, client quick start, release notes and configuration guide
  describe the current outputs and dated delivery layout.

The existing Excel generator needed no changes: its configured native-chart
pipeline automatically appends these definitions and retains the old grid.
An append may fill the unused right-hand slot of the last existing chart row.

## Validation evidence

Before implementation, the complete suite passed **579 tests with one expected
Diesel soil-plot skip**. The new focused suite passed **27 tests**. The complete
post-update suite passed **606 tests with the same one expected skip** across
38 test modules. Python 3.11.15, Streamlit 1.49.1 and all other declared runtime
and development dependency constraints were enforced for both complete runs.

All 164 available calculated-channel arrays across five datasets have identical
SHA-256 hashes before and after this reporting update. Independent four-wheel
sums and signed integration checks agree sample by sample.

| Validation input | Samples | Excel charts before → after | PowerPoint slides before → after | Applicable KPI formulas |
| --- | ---: | ---: | ---: | ---: |
| Real Diesel `_07` | 232 | 20 → 20 | 10 → 10 | 5 |
| Real Electric UoL `_04` | 3,954 | 19 → 23 | 12 → 14 | 12 |
| Available older real Hybrid reference | 1,770 | 21 → 25 | 12 → 14 | 10 |
| Controlled complete Hybrid, generator active | 5 | 25 → 29 | 12 → 14 | 12 |
| Controlled complete Hybrid, generator inactive | 5 | 25 → 29 | 12 → 14 | 12 |

The real Electric final energies are 5.4029504071138925 kWh for bulldozing and
8.075161873669435 kWh for compaction. The older Hybrid reference gives
1.573030332918836 and 3.330067021537997 kWh respectively. These agree with
independent integration over the actual recorded timestamps.

For every dataset, all original native chart XML, main-sheet XML, Statistics
XML, Formula Support XML, styles, equation comments and chart anchors are
unchanged. This preserves existing formulas, cell values, merged cells, row
heights, fonts, colours, alignments, statistics and KPI results. The original
twelve Electric/Hybrid slides and ten Diesel slides retain their image bytes,
positions and text, apart from updated page totals and the displayed plot count.
The final summary remains last. Diesel retains its 24 pt first row and all
20 charts.

Desktop PowerPoint and Excel opened all five generated report pairs. Native
PowerPoint renders of both added pages, plus all four native Excel chart renders,
were visually inspected for the real Electric and available Hybrid reference.
Titles, time axes, units and curves are readable; no clipping or overlap was
found. The first hidden Excel export returned zero-byte PNGs despite reporting
success; activating each chart in hidden Excel produced valid native renders,
which were the files actually inspected.

Evidence is under `outputs/soil_plots_2026_10_16/`:

- `baseline_tests/manifest.json`, `regression_tests/manifest.json` and test logs;
- `focused_final.txt`;
- `signals_before.json` and `signals_after.json`;
- `validated_reports/audit.json` and the generated Excel/PowerPoint files;
- `preservation.json`;
- `validated_reports/visual_checks/` with native PNG/PDF renders;
- `validated_reports/native_office.json` and `render_excel.ps1`, which verifies
  that every activated native Excel chart export is nonempty.

## Validation boundary

The latest updated real Hybrid CSV is not available. Hybrid validation uses the
available older real source plus controlled active/inactive cases; it does not
claim individual acceptance of an unavailable export. That older source lacks
some newer optional auxiliary/torque/height measurements, which remain clearly
unavailable. All four requested soil plots are supported and present for it.

These additions apply to the current October Electric/Hybrid profiles selected
by the client workflow. Explicit legacy reporting remains unchanged.

## Delivery and Git

The delivery folder and ZIP root are `Astauto_VSM_October_16102026`, containing
the two existing BAT launchers, `README_SERGIO.txt`, the manifest and `Tool`.
The canonical release builder supplies application/configuration/scripts,
required branding/templates and client documentation. Raw datasets, reports,
caches, virtual environments and development tests are excluded. The previous
client release is preserved.

Application version remains 1.3.0 and Python compatibility remains
`>=3.11,<3.12`. Packaging records the actual HEAD and reviewed uncommitted files;
it does not claim a clean committed snapshot. No commit or push is performed.

Fresh extraction outside the repository, in a path containing spaces, passed
the unchanged setup with developer Python removed from PATH. It created its own
CPython 3.11.9 runtime and `.venv`; only cached official runtime downloads were
reused. Dependencies were installed normally, `pip check` passed, and doctor
reported zero failures (the expected private example-dataset warning remains).
All five report pairs generated from that installed copy match the source
workbooks' native chart definitions, data/formula sheets and styles. See
`fresh_installation.json`, `fresh_setup.txt`, `fresh_doctor.txt`,
`fresh_pip_check.txt`, `fresh_reports.txt` and `package_audit.json` in the QA folder.

Git remains on `main`, HEAD `7be9bf247b7ce507ffc0c360e26fbeef74eef17d`.
The working tree already contained 15 modified and four untracked files before
this task. Those pending changes were preserved. The final status contains
21 modified and six untracked files:

```text
 M CHANGELOG.md
 M README.md
 M config/report_profiles/sergio_2026_10/full_size_sprayer_diesel.yaml
 M config/report_profiles/sergio_2026_10/robosprayer_electric.yaml
 M config/report_profiles/sergio_2026_10/robosprayer_hybrid.yaml
 M docs/CLIENT_QUICK_START.md
 M docs/CONFIGURATION_GUIDE.md
 M docs/FINAL_RELEASE_NOTES.md
 M docs/GIT_RELEASE.md
 M docs/KNOWN_LIMITATIONS.md
 M docs/TROUBLESHOOTING.md
 M src/vsm_postprocessing/diesel_powerpoint_report.py
 M src/vsm_postprocessing/excel_formulas.py
 M src/vsm_postprocessing/excel_report_engine.py
 M src/vsm_postprocessing/plotting_engine.py
 M src/vsm_postprocessing/profile_plotting.py
 M src/vsm_postprocessing/profile_powerpoint_report_engine.py
 M src/vsm_postprocessing/report_profile.py
 M tests/test_diesel_excel_report.py
 M tests/test_profile_excel_report.py
 M tests/test_sergio_october_15.py
?? docs/FINAL_REPORTING_IMPROVEMENTS_2026_10_16.md
?? docs/OCTOBER_2026_CLIENT_RELEASE_2026_10_16.md
?? docs/OCTOBER_SOIL_PLOTS_2026_10_16.md
?? src/vsm_postprocessing/profile_kpi_formulas.py
?? tests/test_final_reporting_improvements.py
?? tests/test_october_soil_plots.py
```

The shipped root `START_VSM_TOOL.bat` also passed a real Microsoft Edge headless
workflow from an unrelated working directory with spaces, with developer
environment variables removed. Real Diesel, Electric and the available older
Hybrid inputs were uploaded, validated and generated; both download buttons
were present and there were zero Streamlit exceptions. Inspection of those
UI-generated workbooks/decks confirmed all four requested charts and the two
configured pages for Electric/Hybrid. Evidence: `browser_validation.json` and
`ui_output_audit.json`. Both owned QA server/browser ports were closed afterward.

Final accepted delivery:

- Folder: `dist/Astauto_VSM_October_16102026/`
- ZIP: `dist/Astauto_VSM_October_16102026.zip`
- Checksum: `dist/Astauto_VSM_October_16102026.sha256`
- ZIP size: 94,431,681 bytes; 115 files, including its manifest.
- SHA-256: `a312662e7a5dce5f58f414f06ef3588ccf126bf54f747cbcbad892e4078c1ad6`

The clean delivery folder matches the archive byte-for-byte. Every manifest file
hash and shipped source byte matches the reviewed working tree. The two canonical
builds were reproducible. The installation and UI tests used a separate temporary
extraction, leaving this delivery free of environments, generated reports and
caches. Final acceptance details are in `outputs/soil_plots_2026_10_16/delivery.json`.
