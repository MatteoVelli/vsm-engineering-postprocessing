# October 2026 client release — 16 October 2026

**Verdict: READY WITH DOCUMENTED LIMITATIONS.** The tested package replaces the
previous local Sergio deliverable. Updated real Hybrid acceptance is pending;
available Hybrid tests pass and this is not a release gate or a known defect.

## A. New package

- Path: `dist\October_2026_0ec467c\VSM_Engineering_PostProcessing_v1.3.0_Sergio_October_2026_2026-10-16.zip`
- Filename: `VSM_Engineering_PostProcessing_v1.3.0_Sergio_October_2026_2026-10-16.zip`
- Size: **94,427,439 bytes** (90.05 MiB).
- SHA-256: `e65a67b1691cc0218e0a4d465873e972c4cd9c404854bb2039fb368fee460248`
- Checksum file: `dist\October_2026_0ec467c\VSM_Engineering_PostProcessing_v1.3.0_Sergio_October_2026_2026-10-16.sha256`
- ZIP entries: **115**; CRC integrity, per-file hashes and canonical
  repeat-build determinism passed. The tested staged bytes and promoted bytes match.

```text
Astauto_VSM_October_2026/
  SETUP_VSM_TOOL.bat
  START_VSM_TOOL.bat
  README_SERGIO.txt
  RELEASE_MANIFEST.json
  Tool/
    pyproject.toml, .python-version, README.md, CHANGELOG.md
    .streamlit/, config/, scripts/, src/, docs/
    reference_files/   approved branding and layout templates only
    outputs/.gitkeep
```

The established canonical builder and Sergio Tool wrapper were reused. The
reviewed working tree, including the previously uncommitted KPI compiler and
thermal improvements, was packaged directly. The manifest records HEAD,
uncommitted source state and all shipped file hashes; it does not claim to be a
clean committed snapshot. No application version or Git tag was invented.

All runtime Python modules, modern/legacy profiles, corrected RAW mappings,
channel-specific HV kW declaration, formula generation, plotting resources,
Astauto assets and optional revised layout templates are included. Developer
environments, bootstrap caches/runtimes, test tree, raw CSV datasets, generated
reports, Git metadata, credentials and nested ZIPs are excluded. Developer tests
were executed from source; client setup does not require pytest or a Git checkout.

## B. Previous-package replacement

- Previous ZIP: `dist\October_2026_0ec467c\VSM_Engineering_PostProcessing_v1.3.0_Sergio_October_2026_0ec467c.zip`
- Previous size: **56,790,588 bytes**.
- Previous SHA-256: `c470df0db46f95d49bbe9790d86e8aec199297d35fdc32e4493989d9805571c2`
- Verified archival backup: `dist\Archived_Client_Packages\2026-10-16_before_final_update\VSM_Engineering_PostProcessing_v1.3.0_Sergio_October_2026_0ec467c.zip`
- Retired original: `dist\Archived_Client_Packages\2026-10-16_before_final_update\retired_original\VSM_Engineering_PostProcessing_v1.3.0_Sergio_October_2026_0ec467c.zip`

The old artifact was identified by its October build record and technical delivery
report, then its exact recorded checksum was confirmed. A backup and checksum
were preserved before validation. Only after all release gates passed was the new
dated ZIP installed in the same established client-deliverable directory and the
superseded original moved into the archive. The filename changed to reflect the
new build date. The historical parent directory name is retained for continuity;
the manifest records the actual source state. Both old copies remain intact.
Unrelated historical packages and validated outputs were preserved.

Replacement status: **PASS**. Current pointer:
`dist/October_2026_0ec467c/OCTOBER_2026_CURRENT_CLIENT_RELEASE.json`.

## C. Testing

### Complete supported regression suite

**579 passed, 1 expected skip, 0 failures**, all 37 modules. Groups: 166 passed;
233 passed/1 skipped; 180 passed. The skip is the inapplicable Diesel soil-slide
case. Packaging/bootstrap/portability, legacy compatibility, source-unit handling,
optional inputs, UI selections and both reporting engines are included. No tests
were removed and no numerical tolerances were relaxed.

```powershell
.venv\Scripts\python.exe -u outputs/run_october_client_release_tests.py
```

Exact pytest subprocess commands, dependencies, module lists and logs:
`outputs/october_client_release_2026_10_16/regression/manifest.json`.

### Shipped application reporting

| Dataset | Samples / source channels | KPI formulas / applicable | Excel charts | PPT slides | Native Excel maximum numeric difference |
|---|---:|---:|---:|---:|---:|
| diesel | 232 / 608 | 5 / 5 | 20 | 10 | 9.09e-13 |
| electric | 3954 / 608 | 12 / 12 | 19 | 12 | 4.26e-14 |
| hybrid_reference | 1770 / 600 | 10 / 10 | 21 | 12 | 1.42e-13 |
| hybrid_active | 5 / 300 | 12 / 12 | 25 | 12 | 2.49e-14 |
| hybrid_inactive | 5 / 300 | 12 / 12 | 25 | 12 | 3.79e-19 |

Diesel source:
`Full_Size_Sprayer_8500Kg_Diesel_1000m_20kph_4WS_DriverPC_up_to_25pcGrad_SL_LSD_07.csv`.
Import, strict profile validation, math, cached statistics, five KPI formulas,
20 native charts and ten slides passed. Independent source checks verify elapsed
time/distance/fuel deltas, actual-time weighted means and exact SI engine power.
Oil and water sources are separate, both spanning **53.2263–81.0555 °C** in this
input. Native PowerPoint inspection confirms two aligned, independently coloured
charts with common temperature limits and four readable MIN/MAX cards on slide 9.
Native Excel confirms row heights 24/24/60 pt for rows 1/2/3.

The `_07` input differs from the earlier Test_03 dataset; different engineering
values are expected. Original Test_03 golden regressions still pass. The new
input's equations were checked directly, and all installed-package workbook data,
formulas, styles and native chart XML match its source-run baseline exactly.

Electric source:
`RoboSprayer_3500Kg_Electric_UoL_DriveCycle_12kph_DriverPC_20pcGrad_04.csv`.
All 608 channels/3,954 samples were imported. Independent signed wheel-power,
HV/LV sum, auxiliary integration and soil-energy checks passed. Accumulated
auxiliary energy is **12.465459322222488 kWh**, matching the prior same-input
regression reference without altering calculations. All 12 KPI formulas,
19 charts and 12 slides passed. The auxiliary-energy chart remains on slide 7,
with no unexpected missing-plot placeholders. The source-justified nominal
battery-capacity consistency warning remains unchanged.

Hybrid uses the compatible 1,770-sample older real reference and explicitly
synthetic five-sample generator-active/inactive fixtures. Generator power matches
measured-fixture torque × RPM × 2π/60000, and both activity paths generate reports.
The older reference lacks HV auxiliary and road-height inputs: its two dependent
KPIs are omitted, not zero-filled. All 12 KPIs work in both complete fixtures.
Legacy Hybrid compatibility remains covered by the full suite. No updated real
Hybrid measurements or acceptance are claimed.

### Independent Office checks

All **212,754 formula caches** were removed before native Excel full rebuild and
per-sheet calculation. All **1,911,922 numeric cells** agree with Python at
rtol 2e-14 / atol 1e-12; maximum absolute difference is **9.10e-13**. Every applicable
KPI contains a formula and recalculates correctly. No Excel error cells, missing
required numeric caches, circular references or chart-count mismatches occurred.
All five decks opened in native PowerPoint at the expected slide counts.
Diesel thermal and Electric auxiliary slides were visually inspected; images
and PDFs are in the validation directory's `visual_checks/`.

### Clean Windows installation and actual browser workflow

Fresh extraction: `<temporary directory>\VSM October 2026 QA\Astauto_VSM_October_2026`.
This is outside the repository and its path contains spaces. Every manifest hash
was checked, with no preexisting runtime or virtual environment. Missing Python
was forced while clearing developer Python variables and restricting PATH.
The unchanged shipped setup deployed **official CPython 3.11.9**, exercised its
app-local NuGet fallback, created a new `.venv`, installed dependencies and passed
doctor and pip check. Only verified official download caches were supplied; no
installed runtime or developer virtual environment was copied. A fresh transfer
of the Python installer/NuGet download from the internet was not repeated.

The interpreter, base runtime and application module all resolve inside the
extracted Tool; pytest is absent. Twenty shipped PowerShell scripts parse.
The unchanged root START BAT ran from an unrelated directory with spaces, with
developer Python paths removed. Health returned HTTP 200/ok. Isolated headless
Edge completed real Diesel/Electric upload → validation → report generation,
with both download buttons present and zero Streamlit exceptions. An isolated
QA browser stub suppressed only the automatic visible browser; real headless
Edge connected to the actual Streamlit session. Test-owned processes were stopped.

Doctor has zero failures. Its sole warning is expected: the private example
pipeline dataset is not bundled; ordinary UI uploads work.

The first direct API harness used the QA parent directory and therefore activated
the intentionally supported template-missing/configuration-order fallback. All
Electric KPI values still matched. The final strict comparison uses the documented
Tool working directory, exactly as client_start.ps1 does; all five workbook
comparisons pass without weakening assertions or changing application code.

Evidence: `source_reports/audit.json`, `package_audit.json`, `browser_validation.json`,
`client_runtime.json`, `setup_result.json`, `native_powerpoint.json`,
`qa_cwd_diagnosis.json`, `staged_build.json`, `RELEASE_ACCEPTANCE.json` and logs under
`outputs/october_client_release_2026_10_16/`. The staged and installed shipped files
were rechecked against the manifest after setup and report generation.

## D. Version and dependencies

Application version: **1.3.0**, unchanged. Declared Python support:
**>=3.11,<3.12**. Full regression used 3.11.15; actual fresh client bootstrap and
end-to-end tests used 64-bit 3.11.9. No MATLAB or Microsoft Office installation is
required for normal report generation; Office was used only for release validation.

| Runtime dependency | Fresh client version | Declared constraint |
|---|---|---|
| numpy | 2.4.6 | `<3,>=1.26` |
| openpyxl | 3.1.5 | `<4,>=3.1` |
| PyYAML | 6.0.3 | `<7,>=6.0` |
| matplotlib | 3.11.2 | `<4,>=3.9` |
| streamlit | 1.49.1 | `<1.50,>=1.49.1` |
| python-pptx | 1.0.2 | `<2,>=1.0` |

The fresh client selected matplotlib 3.11.2 within the existing constraint;
source regression used 3.11.1. Installed-package numerical/layout checks and
native visual checks passed. Dependency specifications and launchers were unchanged.
First setup needs access to official runtime/package downloads unless caches are
available. This is a source distribution with automatic setup, not an offline EXE/MSI.

## E. Remaining limitations

1. The latest real Hybrid CSV has not been individually acceptance-tested.
   Available real-reference and controlled tests pass; this does not block release.
2. Actual fresh Windows setup passed with verified official Python download caches.
   Live installer/NuGet transfer was not repeated, and a company-specific proxy or
   security policy may require the existing documented manual/offline fallback.
3. Custom API/CLI users should run from Tool for optional template-based ordering;
   the shipped launcher does this automatically. Without templates the established
   configured-order fallback still calculates valid reports.

No outstanding software defect was identified by the release gates.

## F. Git status

Branch: **main**. HEAD: `7be9bf247b7ce507ffc0c360e26fbeef74eef17d`.
All reviewed uncommitted implementation changes were preserved and included.
This packaging task changed client documentation only; engineering calculations,
dependency constraints, launchers and bootstrap implementation were not changed.
No commit, push, reset, history rewrite, tag or GitHub release was performed.
No external delivery was performed. The final Git inventory is recorded alongside
this report in the validation directory.

Final modified/untracked inventory (15 modified tracked files, four new files):

```text
 M CHANGELOG.md
 M README.md
 M config/report_profiles/sergio_2026_10/full_size_sprayer_diesel.yaml
 M docs/CLIENT_QUICK_START.md
 M docs/FINAL_RELEASE_NOTES.md
 M docs/GIT_RELEASE.md
 M docs/KNOWN_LIMITATIONS.md
 M docs/TROUBLESHOOTING.md
 M src/vsm_postprocessing/diesel_powerpoint_report.py
 M src/vsm_postprocessing/excel_formulas.py
 M src/vsm_postprocessing/excel_report_engine.py
 M src/vsm_postprocessing/plotting_engine.py
 M src/vsm_postprocessing/profile_plotting.py
 M tests/test_diesel_excel_report.py
 M tests/test_profile_excel_report.py
?? docs/FINAL_REPORTING_IMPROVEMENTS_2026_10_16.md
?? docs/OCTOBER_2026_CLIENT_RELEASE_2026_10_16.md
?? src/vsm_postprocessing/profile_kpi_formulas.py
?? tests/test_final_reporting_improvements.py
```

## G. Client readiness

**READY WITH DOCUMENTED LIMITATIONS.** Release gates pass, including updated real
Diesel/Electric browser workflows, all available Hybrid paths, a clean actual
Windows setup, independent numeric recalculation, native visual inspection,
ZIP integrity and verified archival replacement. The latest real Hybrid input
is pending confirmation rather than a failed release condition.

Client steps: extract the entire ZIP to a short writable Windows folder; run
SETUP_VSM_TOOL.bat; after success run START_VSM_TOOL.bat. Upload a VSM result,
choose its powertrain, Validate, Generate Engineering Report and download Excel
and optional PowerPoint. Keep the Tool folder together with the root launchers.
