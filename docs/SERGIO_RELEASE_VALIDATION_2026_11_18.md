# Sergio release validation — 18.11.26

**READY FOR SERGIO — PASS**

Application version remains **1.4.1**, consistent with the established dated
Sergio delivery policy. Release date/filename label is explicitly **18 November
2026 / 18.11.26**. Execution crossed midnight; validation completed at
`2026-11-18T23:14:30.305469+00:00`. No version increment was required.

| Identity | Verified value |
| --- | --- |
| ZIP | `VSM_Engineering_PostProcessing_v1.4.1_Sergio_Ready_18.11.26.zip` |
| Path | `C:\Users\user\Desktop\Agro Project\dist\Sergio_Request\VSM_Engineering_PostProcessing_v1.4.1_Sergio_Ready_18.11.26.zip` |
| Size | 94,500,737 bytes |
| Files, including manifest | 127 |
| SHA-256 | `fb8942269f5c5c489d108b074315ba77b18d413a1966694270662611fa27ba9a` |

The package uses the existing `VSM_Sergio_v1.4.1/Tool` layout and existing
collector, deterministic ZIP writer and Sergio builder. Only a validated optional
DD.MM.YY filename suffix and inclusion of the motor guide were added to packaging.
The archive's source state is explicitly an uncommitted working tree, with
per-file hashes in `RELEASE_MANIFEST.json` and the base commit recorded.

## Previous delivery and rollback

Previous active ZIP: `VSM_Engineering_PostProcessing_v1.4.1_Sergio_Ready.zip`.

Verified archive: `C:\Users\user\Desktop\Agro Project\dist\Archived_Client_Packages\2026-11-18_before_motor_excel_import_delivery\VSM_Engineering_PostProcessing_v1.4.1_Sergio_Ready.zip`.

Previous SHA-256: `420f413aa0475ae6d88fd2f0aa05b354a1bc93a721a9b2f73133e05ec35e7cec`.

Before promotion, the ZIP, checksum, release record and current-delivery pointers
were copied to the new dated archive. SHA-256 and ZIP CRC/readability checks
passed. Existing historical archives were never overwritten. Promotion uses a
verified pending copy and the established Windows File.Replace plus same-folder
rename. The original verified historical copy remains intact. Exactly one active
official ZIP is required; the dated checksum, README and release record accompany it.
The first atomic call rejected PowerShell's null backup-string binding before
replacement; the original active checksum was rechecked unchanged. The successful
retry used a new `FileReplace_previous_Sergio_Ready.zip` rollback file rather than
overwriting any existing historical backup.

## Final validation

- Full supported-environment suite: **1,232 passed,
  1 expected skip, 0 failures, 0 errors** across
  1,233 cases. Pytest reported 1,453.92 s.
- All **48 motor-import tests** and **5 dated-packaging tests** passed in that
  full run. The skip is the inapplicable Diesel soil-plot case. The two warnings
  are existing sparse peak/continuous motor-power fixture warnings.
- Python **3.11.9 x64**, Streamlit **1.49.1**, NumPy **2.4.6**, openpyxl **3.1.5**,
  PyYAML **6.0.3**, matplotlib **3.11.2**, python-pptx **1.0.2**; pytest **8.4.2**.
- Independent extraction and fresh installation at `C:\VSM QA 181126`:
  staged ZIP CRC, all manifest hashes, source equality, setup, pip check, doctor,
  actual root START BAT, HTTP health/page and real headless Edge workflows passed.
  The shipped setup bootstrapped its own CPython using its existing official
  NuGet fallback after the normal installer did not create the local interpreter.
  No developer environment, old caches or old runtime was copied.
- The new package contains 127 expected files, with a maximum extracted member
  path length of 110 characters in the clean QA folder. No Git metadata, developer
  environment, secrets, test workbooks, generated reports or obsolete ZIPs are bundled.
- Bootstrap scripts, dependency declarations and core calculation/report modules
  are byte-identical to the previous official delivery. Peak/continuous curves,
  interpolation/no extrapolation, torque limits, power, statistics and KPIs are
  exercised by the full regression and exact report comparisons.

| Actual reference workflow | Samples | Native Excel charts | PowerPoint slides | Result |
| --- | ---: | ---: | ---: | --- |
| Diesel | 232 | 21 | 11 | PASS |
| Electric | 3,954 | 30 | 17 | PASS |
| Hybrid | 1,770 | 30 | 16 | PASS |

Electric and Hybrid browser tests uploaded the reordered, unordered-RPM `.xlsx`
fixture directly, displayed four correctly paired points and torque curves,
entered QA-only ratio/efficiency, rejected duplicate RPM, blocked generation,
accepted the valid replacement and generated both reports. Zero Streamlit
exceptions occurred. Manual fallback, replacement, source isolation and reruns
also passed in the automated tests.

Installed browser outputs match independent manual-entry motor baselines using
identical uploaded VSM bytes/path, report selections and metadata. All engineering
worksheet XML (formulas and cached values), native chart XML, PowerPoint slide
XML and plotted images are exact matches. Run-specific Metadata is excluded.
Motor Duty Summary source provenance is compared without normalization. All slide
objects remain within the slide bounds. Four native motor torque/power series and
the operating maps are present; shaft power values match the existing formula.
The generated illustrative motor chart was also visually inspected.

## Reproduction and evidence

```powershell
& 'C:\VSM QA 151126\VSM_Sergio_v1.4.1\Tool\.venv\Scripts\python.exe' outputs/release_2026_11_18/run_regression.py
python scripts/build_sergio_release.py --output-dir outputs/release_2026_11_18/staging --release-date 2026-11-18 --date-label 18.11.26 --allow-uncommitted-source
& 'C:\VSM QA 181126\VSM_Sergio_v1.4.1\Tool\.venv\Scripts\python.exe' outputs/release_2026_11_18/validate_browser.py
& 'C:\VSM QA 181126\VSM_Sergio_v1.4.1\Tool\.venv\Scripts\python.exe' outputs/release_2026_11_18/validate_installed_reports.py
python outputs/release_2026_11_18/prepare_release_record.py
powershell -NoProfile -ExecutionPolicy Bypass -File outputs/release_2026_11_18/promote.ps1
```

`run_regression.py` records the exact pytest arguments and source/runtime module
paths in `regression_environment.json`. Its source-first PYTHONPATH includes only
isolated pytest tools for child processes; production dependencies resolve from
the supported installation. Evidence remains under `outputs/release_2026_11_18/`:
`regression.log`, `regression.xml`, `fresh_installation.json`, `fresh_environment.json`,
`browser_validation.json`, `installed_reports_validation.json`, `package_audit.json`
and the final release record. The official record is also beside the delivery ZIP.

The first full attempt exposed three maintenance subprocess environment failures:
pytest was initially visible only to the parent runner. Isolated test-only tools
fixed child execution; all 15 maintenance tests and the entire suite then passed.
Browser waits and numeric commits were corrected in QA automation only. Initial
comparison mismatches were source-path provenance and an unmatched default SUM
selection; identical source paths/options restored exact comparison without
altering engineering formulas or weakening numerical checks.

## Limitations and client instructions

These are real VSM reference datasets paired with **illustrative software-test
motor limits**, not manufacturer acceptance. QA ratio 17.5 and efficiency 0.97
are not universal defaults. No illustrative workbook or report is delivered to
Sergio as approved engineering data. Newer Hybrid exports and real manufacturer
curves were not supplied; acceptance of them is not claimed. Available Caiman
negative-mass diagnostics remain visible.

Validation covers the available Windows host, not every client PC/network.
First setup still needs internet. Doctor's only warning is the deliberately
unbundled example CSV; actual uploads pass. Excel, OneDrive and clipboard APIs
are not required. No unresolved release-blocking issue remains once promoted.

Sergio should extract to a fresh short local folder, run SETUP then START, open
Motor Performance Curves, enable the shared curve, download the blank template,
populate Motor RPM / Peak Torque / Cont Torque in rpm/Nm with actual manufacturer
data, upload `.xlsx`, review the curves, enter applicable ratio/efficiency, validate
and generate reports. Clipboard operations are unnecessary. The packaged client
README, quick start and motor guide provide these instructions without test values.

## Git and exact source scope

No commit, push, tag, staging or repository cleanup was performed. The initial
unrelated Diesel/reporting/maintenance changes were preserved and their hashes
checked. Generated archives, QA reports and build outputs remain outside source
control. Release-phase modified/new files:

- `CHANGELOG.md`
- `README.md`
- `docs/CLIENT_QUICK_START.md`
- `docs/FINAL_RELEASE_NOTES.md`
- `docs/GIT_RELEASE.md`
- `docs/SERGIO_CLIENT_README.md`
- `docs/MOTOR_TORQUE_EXCEL_IMPORT.md`
- `scripts/build_sergio_release.py`
- `src/vsm_postprocessing/release_builder.py`
- `tests/test_release_packaging.py`
- `docs/SERGIO_RELEASE_VALIDATION_2026_11_18.md`

The importer implementation files reviewed in this phase were unchanged from
the completed implementation. A proposed logical commit combines them with the
necessary packaging/tests/documentation files below, leaving unrelated initial
changes separate:

- `src/vsm_postprocessing/motor_torque_import.py`
- `src/vsm_postprocessing/motor_performance_ui.py`
- `tests/test_motor_torque_import.py`
- `CHANGELOG.md`
- `README.md`
- `docs/CLIENT_QUICK_START.md`
- `docs/FINAL_RELEASE_NOTES.md`
- `docs/GIT_RELEASE.md`
- `docs/SERGIO_CLIENT_README.md`
- `docs/MOTOR_TORQUE_EXCEL_IMPORT.md`
- `scripts/build_sergio_release.py`
- `src/vsm_postprocessing/release_builder.py`
- `tests/test_release_packaging.py`
- `docs/SERGIO_RELEASE_VALIDATION_2026_11_18.md`

Proposed message: `Release Sergio v1.4.1 with Excel motor map import (18.11.26)`.

The complete Git status and per-file package hashes are recorded in the official
release record. `git diff --check` passed.
