# VSM maintenance: controlled test storage

Use an existing compatible Python environment; this runner never installs
packages or changes environments. The developer `.venv` currently has Streamlit
1.61.1, outside `pyproject.toml`'s `>=1.49.1,<1.50` requirement.

The retained `outputs/release_2026_11_04/qa_env/Scripts/python.exe` was confirmed
as the interpreter used for Batch 2's 144 passing tests. Its Python 3.11.9,
Streamlit 1.49.1, matplotlib 3.11.2, numpy 2.4.6, openpyxl 3.1.5,
python-pptx 1.0.2, PyYAML 6.0.3 and pytest 8.4.2 satisfy current requirements.
Imports resolve the checkout's 1.4.1 code; installed distribution metadata is
1.4.1. The tracked `src/vsm_postprocessing.egg-info` still describes 1.3.0 and
can take precedence in metadata queries when `src` is explicitly placed first
on `sys.path`. This is stale packaging metadata, not the imported code version.
The developer environment also has a stale editable 1.3.0 installation. Neither
metadata nor dependencies are repaired by this milestone. The origin of the
developer Streamlit drift has not been established.

## Running focused tests in PowerShell

From the repository root:

```powershell
$qaPython = '.\outputs\release_2026_11_04\qa_env\Scripts\python.exe'
& $qaPython -B .\maintenance\run_tests.py --dry-run tests/test_statistics.py
& $qaPython -B .\maintenance\run_tests.py tests/test_statistics.py tests/test_math_channels.py
```

Select explicit Python test files or `file::test` nodes inside `tests/`.
No selection means an argument error; the runner never starts the whole suite
implicitly. Use `--temp-root 'D:\VSM_Test_Scratch'` to select an **existing**
external parent, or set `VSM_TEST_TEMP_ROOT`. The default is system temporary
storage. A parent inside this repository is refused. `--dry-run` shows the plan
without creating a workspace or running pytest. `--keep-success` retains a
successful workspace for deliberate review.

Each invocation creates a unique `vsm-test-*` child with an ownership marker.
Pytest's `--basetemp`, JUnit XML, combined log, process temporary directories,
Matplotlib configuration and common cache directory stay inside that child.
Bytecode generation and pytest caching are disabled. Inherited pytest options
and plugin injection are cleared; third-party plugin autoloading is disabled.
The current checkout's `src` is prepended to the child Python path, and the
repository remains the working directory for existing relative configuration
paths. No project, dataset or environment is copied by the runner.

Only a pytest exit code of zero permits automatic cleanup. Before deletion,
the runner verifies the exact parent, directory identity, ownership token and
absence of symlinks, junctions or other reparse points. It never deletes the
temporary parent, sibling runs or historical directories. It does not retry
using permission changes. Cleanup refusal is an error with evidence retained;
a filesystem error during deletion may leave a partially cleaned workspace.
Failed/interrupted runs remain at the printed path with available evidence.
The runner does not sweep previous runs on startup or exit.

This is storage routing, not an OS sandbox: a selected test with hard-coded
absolute outputs or installer/browser subprocesses can bypass these defaults.
Review such tests before running them. This milestone uses focused tests and
does not run installers, rebuild Sergio's release, or certify interactive UI
and native Microsoft Office automation. Interrupted processes may take time to
exit; inspect for remaining child processes before manually removing evidence.

## Preservation and future cleanup

The existing pytest configuration only sets `testpaths` and quiet output; it
does not manage manually selected temporary roots. Packaging regressions build
multiple ZIPs and copy source/template trees, and a bootstrap regression creates
a venv. Repeated runs in individually named repository workspaces retain those
copies. Manual report/browser validation adds further retained outputs. The
new runner changes storage handling only when explicitly invoked; existing
direct pytest commands, CI and manual validators retain their current behavior.

| Location | Handling |
| --- | --- |
| `src/`, `config/`, `scripts/`, `tests/`, documentation | Maintained source and regression inputs; preserve. |
| `reference_files/`, `tests/fixtures/` | Original simulation data, report templates and baselines; preserve. |
| `dist/Sergio_Request/` | Preserve the dated 1.4.1 ZIP, checksum and release records unchanged. |
| `dist/Archived_Client_Packages/` | Preserve rollback archives and their provenance. |
| `outputs/`, `.pytest_tmp/` | Mixed historical evidence, generated reports and active baselines; never blanket-delete. |
| `.venv`, retained QA/client runtimes | Preserve; environment repair is a separate approved operation. |
| New external `vsm-test-*` workspace | Disposable only after its own successful run, or after individual review of retained evidence. |
| Existing `maintenance/` audit files, invoices and owner files | Preserve; not runner cleanup candidates. |

The authoritative ZIP remains
`dist/Sergio_Request/VSM_Engineering_PostProcessing_v1.4.1_041126.zip`, SHA-256
`1F41414952BA9695779CD576CADC3D8B9F08864D350E6FE2CE47B1314862760D`.
Filename inconsistencies in release records are deliberately unchanged.

For future cleanup, first list exact candidate paths and sizes without deleting,
trace consumers against the retained cleanup audit, and review ownership,
reparse points, hashes and provenance. Preserve useful failure logs/JUnit and
engineering baselines before approving specific removals. A `tmp` name or an
ignore rule alone does not prove disposability. This runner supplies no bulk
cleanup command and no authorization to remove historical artifacts.

If the QA interpreter becomes unavailable, stop. Its base Python runtime lives
outside the repository, so retaining the venv alone is insufficient. An approved
recovery should create a separate Python 3.11 environment, use the retained
`outputs/release_2026_11_04/validated_constraints.txt` with the declared project
and dev requirements, then verify imports, dependency consistency and focused
tests. Do not reinstall into either existing environment without approval.

## Milestone validation: 5 November 2026

- Baseline: 29 passed, zero failures/errors/skips, with writes guarded to a
  fresh external workspace. Generated artifacts occupied 42,470,075 bytes.
- Runner safety: 15 passed, including a real Windows junction refusal,
  preservation of pre-existing siblings and expected synthetic failure evidence.
- Same application regressions through the runner: 29 passed, zero
  failures/errors/skips; its successful workspace was automatically removed.
- Coverage includes deterministic math/statistics, CSV/XLSX imports, profiles,
  native Excel chart structure and numerical cells, reopenable Electric/Hybrid
  PowerPoint decks, plot/manifest determinism and static packaging contracts.
- No installer or release build was executed. These focused results do not
  recertify the complete application, interactive UI or native Office rendering.
