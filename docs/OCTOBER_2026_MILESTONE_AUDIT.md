# October 2026 development milestone: Git and validation audit

This is a development milestone, **not the final Astauto client release**.
Updated real-source Hybrid acceptance and final packaging remain pending.

## Reviewed scope

The audit started on `main`, tracking `origin/main`, at `0ec467c` in
`MatteoVelli/vsm-engineering-postprocessing`. An authenticated fetch confirmed no
local/remote divergence before publication. Existing history is preserved.

Reviewed changes cover the versioned Sergio profiles, exact gearbox torque
sources, signed wheel power, HV/LV auxiliaries, soil resistance power/energy,
height/gradient plots, Diesel coolant and wheel analyses, live Excel statistics
and integration, explicit report selections, stable Streamlit option identities,
optional-channel diagnostics, source-unit provenance and report regression tests.
The new profiles inherit previously committed Diesel agrochemical and auxiliary
definitions; existing engineering functions were reviewed together with the diff.

Sergio's 16 October confirmation is configured only for the exact HV source
`User_defined_Data_High_Voltage_Auxiliary_signal1_0`: reported `-`, effective kW,
unchanged values, factor 1, offset 0. Other unknown/incompatible units retain
strict validation. Legacy profiles retain their explicit historical conventions.

The verified cosmetic refinement is included: Diesel row 1 is 24 pt. Before/after
native Excel verification found the sole main-sheet XML difference to be row 1's
height. All 20 charts, 4,792 formulas, cell values, styles, merged cells, row 2 and
row 3 were retained. The focused header test also covers Electric and Hybrid.

## Engineering evidence and boundaries

- Diesel: real Test_03 CSV, 223 samples/586 source channels, plus controlled
  boundary cases and the supplied reference workbook arrays.
- Electric: real UoL CSV, 3,954 samples/608 channels; all four nonzero gearbox
  signals and confirmed HV auxiliaries independently checked. The report exports
  330 channels and 19 charts. Auxiliary energy is 12.4654593222 kWh; slide 7's
  chart and 12.47 kWh card are restored with the existing layout. Native Excel
  recalculated 1,306,686 numeric cells within 4.55e-13 after clearing caches.
- Hybrid: updated implementation, unit handling, signed integration and reporting
  tested with controlled fixtures. Available older sources omit the new HV and
  gearbox signals; they do not establish complete real-data acceptance.
- The existing battery Energy/SOC capacity-consistency warning remains separate:
  median 98.2263 kWh, maximum deviation 1.86791 kWh versus 0.982263 kWh tolerance.
  No battery-capacity calculation was changed to hide it.

Earlier assessments are retained as historical records, with links to the later
HV confirmation. The current detailed evidence is in
[the HV assessment](HV_AUXILIARY_CONFIRMATION_2026_10_16.md); the selection analysis
is in [its root-cause report](ELECTRIC_SELECTION_ROOT_CAUSE_2026_10_16.md).

## Final regression gate

The full suite was rerun using Python 3.11.15 and the declared supported Streamlit
1.49.1 environment. All declared runtime and test dependency constraints are
checked before test execution:

| Package | Tested version |
| --- | --- |
| NumPy | 2.4.6 |
| openpyxl | 3.1.5 |
| PyYAML | 6.0.3 |
| Matplotlib | 3.11.1 |
| Streamlit | 1.49.1 |
| python-pptx | 1.0.2 |
| pytest | 8.4.2 |

**Final pre-commit gate: 557 passed, one expected skip, zero failures across all
36 test modules.** Group results were 187 passed/one skipped (372.22 s), 189
passed (566.10 s), and 181 passed (345.91 s), all exit code zero. This comprises
the previous 554-pass baseline and three additional header-height cases.
All seven packaging tests and all 20 bootstrap tests passed within this run.

Additional checks passed: 38 package-module imports, six profile configurations,
AST parsing of candidate Python files, and Git whitespace checks. The installation
health check returned PASS with zero warnings/failures. All 33 candidate files
were reviewed as text; the personal-path and credential-pattern scan was clear.

Local invocation:

```powershell
.venv\Scripts\python.exe outputs/run_october_git_audit.py
```

The local helper runs all 36 test modules in three isolated pytest processes,
propagating the supported environment to child processes. Its full command lists,
file inventory, versions, logs and exit codes are retained under the ignored
`outputs/october_git_audit/`. Short workspace temporary paths avoid the earlier
Windows bootstrap path issue. The only deliberately inapplicable case is the
existing Diesel soil-slide parametrization. Packaging/bootstrap and actual
Excel/PowerPoint generation tests are part of the complete suite; no test is
weakened or removed.

Equivalent full-suite command in a fresh Python 3.11 environment with the declared
dependencies installed:

```powershell
python -m pip install -e '.[dev]'
python -m pytest -p no:cacheprovider -o addopts= -q -ra --basetemp outputs/tests
python -m vsm_postprocessing.doctor_cli --project-root .
git diff --check
```

Private source-backed tests run when the inputs are installed locally; they skip
explicitly when those files are absent from a clean checkout. The independent
real Electric audit is reproducible with a supplied private source:

```powershell
python scripts/validate_electric_uol.py '<path-to-private-electric-vsm.csv>' --output-dir outputs/new_electric_validation
```

## Publication hygiene and commit organization

The reviewed implementation, profiles, tests and portable validation utilities
form one coherent feature commit because they share schema/selection interfaces.
A separate documentation commit records the assessments, final validation and
remaining acceptance boundary. No existing commit is amended or rewritten.

Personal absolute paths in new documentation examples are replaced with portable
placeholders. Source data, generated reports, test folders, virtual environments,
caches and local validation helpers remain ignored and unstaged. No credentials
or local secrets are included. Existing approved tracked runtime assets remain
unchanged; optional private `_03`/`_07` workbooks stay excluded under the current
reference-file policy. Revised profiles fall back to configuration ordering when
those private ordering references are absent.

Only explicit reviewed file lists are staged. The staged diff is inspected and
checked before committing, and a second fetch/synchronization check precedes a
normal, non-forced push to `origin/main`. No final release tag or client ZIP is
created. Disposable archives made by packaging tests are validation artifacts.
