# Repository cleanup audit - 24 October 2026

This cleanup changes no production code, configuration, test assertions, numerical
calculations or report layouts. The pre-cleanup suite contained 869 collected cases
(868 passed, one existing skip in the preceding accepted run). No test is removed,
renamed or consolidated.

## REMOVE

The 45 paths below are disposable pytest workspaces, generated fixture copies,
or failed optional-evaluator caches from superseded development runs. Test fixture
creation and release-package tests explain their copied sources/templates; none is
an original VSM dataset, approved reference file or client deliverable. No active
application, test, script, CI or packaging input depends on these paths. Historical
commands naming old --basetemp locations are records of completed runs, not required
inputs. All resolved deletion targets were checked against the workspace boundary,
reparse points and protected file list before removal.

- `outputs/agro_formula_all_tmp`
- `outputs/agro_formula_eval_vendor`
- `outputs/agro_formula_new_fixed_tmp`
- `outputs/agro_formula_new_tmp`
- `outputs/agro_formula_regression_tmp`
- `outputs/agro_formula_uv_cache`
- `outputs/agro_policy_accepted_all_tmp`
- `outputs/agro_policy_complete_tmp`
- `outputs/agro_policy_diesel_check_tmp`
- `outputs/agro_policy_excel_final_tmp`
- `outputs/agro_policy_excel_metadata_tmp`
- `outputs/agro_policy_excel_recheck_tmp`
- `outputs/agro_policy_final_all_tmp`
- `outputs/agro_policy_final_contract_tmp`
- `outputs/agro_policy_final_labels_tmp`
- `outputs/agro_policy_final_suite_tmp`
- `outputs/agro_policy_fix_check_tmp`
- `outputs/agro_policy_full_tmp`
- `outputs/agro_policy_new_tmp`
- `outputs/agro_policy_new_verified_tmp`
- `outputs/agro_policy_oct_check_tmp`
- `outputs/agro_policy_probe_tmp`
- `outputs/agro_policy_registry_tmp`
- `outputs/agro_policy_release_check_tmp`
- `outputs/agro_policy_rr_check_tmp`
- `outputs/agro_policy_strict_final_tmp`
- `outputs/agro_policy_target_tmp`
- `outputs/agro_policy_verified_tmp`
- `outputs/agro_policy_zero_fixed_tmp`
- `outputs/agro_policy_zero_tmp`
- `outputs/agro_simplify_all_tmp`
- `outputs/agro_simplify_contract_tmp`
- `outputs/agro_simplify_direct_debug_tmp`
- `outputs/agro_simplify_focused_tmp`
- `outputs/agro_simplify_initial_tmp`
- `outputs/agro_simplify_workflow_fixed_tmp`
- `outputs/agro_simplify_workflow_tmp`
- `outputs/machine_audit_complete_tmp`
- `outputs/machine_audit_delta_tmp`
- `outputs/machine_audit_final_full_tmp`
- `outputs/machine_audit_final_target_tmp`
- `outputs/machine_audit_full_tmp`
- `outputs/machine_audit_targeted_tmp`
- `outputs/machine_audit_verified_target_tmp`
- `outputs/machine_specs_checks_tmp`

## ARCHIVE

Three superseded engineering records moved intact to `docs/archive/2026-10-agro/`:
`AGRO_SOURCE_POLICY_2026_10_23.txt`, `AGRO_STATISTICS_FORMULAS_2026_10_23.txt`, and
`TEMPLATE_AGRO_WORKFLOW_2026_10_24.txt`. Their history, physical assumptions and
validation records are preserved. Current document references were updated; the
archive README distinguishes historical commands from current tools.

The following 45 files were copied to
`outputs/repository_cleanup_2026_10_24/historical_verification.zip`, verified by ZIP
CRC and individual SHA-256, then removed from their active locations. This local
ignored archive is historical evidence, not a runtime or packaging dependency.
It contains two superseded Python wrappers, five one-off staging/inventory scripts
and 38 unreferenced intermediate logs. No PowerShell validation tool was removed.

- `scripts/audit_agro_source_policy.py`
- `scripts/validate_agro_statistics_formulas.py`
- `outputs/prepare_agro_formula_excel_checks.py`
- `outputs/prepare_agro_removal_native.py`
- `outputs/prepare_template_agro_native.py`
- `outputs/prepare_template_chart_checks.py`
- `outputs/verify_agro_removal_inventory.py`
- `outputs/agro_formula_evaluator_install.log`
- `outputs/agro_formula_new.log`
- `outputs/agro_formula_reference_generation.log`
- `outputs/agro_policy_complete.log`
- `outputs/agro_policy_diesel_check.log`
- `outputs/agro_policy_excel_final.log`
- `outputs/agro_policy_excel_metadata.log`
- `outputs/agro_policy_excel_recheck.log`
- `outputs/agro_policy_final_all.log`
- `outputs/agro_policy_final_labels.log`
- `outputs/agro_policy_final_suite.log`
- `outputs/agro_policy_full.log`
- `outputs/agro_policy_generation.log`
- `outputs/agro_policy_generation_final.log`
- `outputs/agro_policy_new.log`
- `outputs/agro_policy_new_verified.log`
- `outputs/agro_policy_oct_check.log`
- `outputs/agro_policy_probe.log`
- `outputs/agro_policy_release_check.log`
- `outputs/agro_policy_rr_check.log`
- `outputs/agro_policy_target.log`
- `outputs/agro_policy_verified.log`
- `outputs/agro_policy_zero.log`
- `outputs/agro_policy_zero_fixed.log`
- `outputs/agro_simplify_direct_debug.log`
- `outputs/agro_simplify_focused.log`
- `outputs/agro_simplify_initial.log`
- `outputs/agro_simplify_reference_generation.log`
- `outputs/agro_simplify_workflow.log`
- `outputs/agro_simplify_workflow_fixed.log`
- `outputs/machine_audit_delta.log`
- `outputs/machine_audit_final_full.log`
- `outputs/machine_audit_full.log`
- `outputs/machine_audit_generation.log`
- `outputs/machine_audit_head_comparison.log`
- `outputs/machine_audit_targeted.log`
- `outputs/machine_audit_verified_target.log`
- `outputs/machine_specs_checks.log`

Total files deleted from obsolete workspaces and ZIP-archived original locations:
50371. The three historical documents were relocated separately.
Total removed bytes: 8977582860 (approximately 8.36 GiB), before the
small retained archive and newly generated validation evidence. The machine-readable
local manifest inventories every removed file, sizes, reasons and archive hashes:
`outputs/repository_cleanup_2026_10_24/cleanup_manifest.json`.

## KEEP

- Every test in `tests/`, including Electric/Hybrid/Diesel, dynamic source and
  unseen-filename cases, invalid data, physical safeguards, machine specifications,
  genuine Excel formulas, source-export independence, Diesel gravity, PowerPoint,
  numerical comparisons and packaging. Debugging origin is not a removal reason.
- `scripts/validate_template_agro_workflow.py`: current combined regeneration,
  four formula checks/cache-free copies, no source worksheet, changed-load unknown
  filename proof and exact saved-engine comparisons. It supersedes both removed
  wrappers, whose separate scenarios also remain covered by the unit/regression tests.
- `scripts/audit_machine_specifications.py`: imported by that validator; generates
  both Electric references, Hybrid and Diesel, compares original outputs and hashes.
- `scripts/validate_agro_statistics_formulas_excel.ps1`: genuine native Excel formula,
  value, raw-input mutation and chart-series checks; cannot be replaced by cache tests.
- All other reusable Python/PowerShell tools. Apparent Office-script overlap is
  different coverage: general legacy KPI recalculation, Diesel RMS and live edits,
  editable PowerPoint geometry/overflow, renders/PDFs and agro-specific formulas.
- Saved `outputs/machine_audit_2026_10_23/*_HEAD.npz` and
  `outputs/agro_source_policy_2026_10_23/audit.json`: active inputs to the retained
  workflow validator. Their complete surrounding evidence folders remain intact.
- Latest `outputs/agro_source_removal_2026_10_24/`,
  `outputs/agro_source_removal_native/`, `outputs/agro_removal_all_tmp/`,
  `outputs/agro_removal_contract_tmp/` and final cited logs. These retain the actual
  308-workbook inventory, 332 native checks and all evidence paths rather than
  leaving historical JSON pointing to deleted acceptance workbooks.
- All original `reference_files/` data/templates/assets, original UI runs (including
  20261023_202309_023085 and 20261023_202617_048134), `dist/`, `ui_workspace/`,
  client/final delivery folders and accepted older evidence. None selected for deletion.
- Current final-removal document and original machine-reporting root-cause audit;
  all approved/client documentation and reference inventories.
- `.gitignore`: already excludes outputs, dist, Python caches and private inputs,
  and preserves explicit runtime-asset exceptions. No additional ignore rule needed.

## Duplicated logic and dependencies

The removed source-policy wrapper repeated four-profile report generation, saved
HEAD array/statistic/KPI comparisons and a renamed-machine proof. The current
workflow performs those comparisons and a stronger halved-load unknown-machine
proof. The removed formula wrapper repeated workbook generation, four Statistics
formula/cache checks and cache clearing. Those checks remain in the current
workflow, focused tests and native Excel runner. No equivalent test scenarios or
assertions were deleted.

CI runs pytest and doctor; neither references the removed files. Setuptools finds
only `src/` packages. The client release builder recursively copies `scripts/`, so
removing the obsolete wrappers also removes them from future packages. No script,
entry point, manifest-required asset or client launch dependency imports/calls them.
The retained validator's import of `audit_machine_specifications` is preserved.
Historical documentation references to removed wrappers explicitly point to the
archive, with current reproduction commands provided alongside.

## REVIEW - left untouched

- `outputs/agro_formula_policy_tmp/`: access denied during inventory. No permission
  or ACL changes attempted and no deletion claimed; ownership/content needs review.
- Eight byte-identical Phase 1 document/inventory pairs in the root and
  `docs/phase_1/`. These are authoritative engineering records, and their intended
  canonical distribution location is not established. Recommend choosing one
  location explicitly in a later documentation-only change; retain both now.
- `Matteo_Velli_Manas_Studio_Invoice.pdf` and `Immagine 2026-08-05 085355.png`:
  user-supplied purpose/ownership cannot be established;
  retain without inspection or modification.
- Older cryptically named output folders, old delivery ZIPs/QA installations,
  `outputs/machine_audit_head/` and its ZIP: release/history ownership is uncertain.
  No blanket output purge; retain pending a retention decision.
- Older one-off utilities inside `outputs/` that are cited by previous engineering
  handoffs remain. They require a separate historical retention/reproduction audit.

## Verification

Collection before/after is identical: 869 test cases with unchanged per-file counts.
The 118 targeted packaging, formula, source-policy, machine-specification and
template-workflow tests passed in 305.36 seconds. All retained Python and PowerShell
validation scripts parse; 23 local Markdown links resolve, and no retained script
references the removed utilities. Application doctor passes with zero warnings
and zero failures. `git diff --check` passes (only existing CRLF informational
warnings). All 6,784 protected files match their before-cleanup SHA-256 hashes.

Fresh generation passed for Electric 3500/8500, Hybrid, Diesel and the derived
unknown-filename/halved-load fixture: five XLSX/PPTX report sets. Every slide's
machine-specification notes match its Excel metadata; no source worksheet exists.
All 110 legacy arrays, 161 statistics and 39 KPIs remain exactly equal to the saved
HEAD baselines. The retained native Excel runner independently recalculated all
four Statistics formulas in all five cache-free workbooks: 20 checks passed.

The complete suite was invoked once: 867 passed, one failed and one existing skip
in 1315.40 seconds (21m55s). Its client-bootstrap virtual-environment test failed
under the unnecessarily long validation path
`outputs/repository_cleanup_2026_10_24/full_tmp`. Reproducing `ensurepip` at that
same path confirms pip's explicit Windows long-path error while installing
setuptools. The identical test passes with `--basetemp outputs/clnp_bootstrap_tmp`
(one passed in 25.63 seconds). No bootstrap code, tests, ACLs or host settings were
changed to conceal the failure. The original full invocation is not an entirely
green run. Across that run plus the short-path rerun, all 868 non-skipped cases
passed; coverage and collection remain unchanged. The existing skip remains
`tests/test_sergio_october_15.py:338` (soil plots do not apply to Diesel).

The 162 generated profile/reference/native workbooks inventoried after validation
contain no Agro Mass Source worksheet, including hidden sheets. This additional
inventory excludes original template copies and the deliberate text-only XLSX
placeholders created by packaging tests; these are not generated client reports.

Local evidence is under `outputs/repository_cleanup_2026_10_24/`: `targeted.log`,
`full_suite.log`/`full_suite.xml`, `bootstrap_short_path.log`,
`bootstrap_long_path_reproduction.log`, `bootstrap_ensurepip_detail.log`,
`report_generation.log`, `generation_smoke.json`, `native_excel.log`,
`reports/native_excel/results.json`, `doctor.log`, `dependency_checks.json`,
`protected_before.json`/`protected_after.json`, `removed_paths.json`, and the
per-file `cleanup_manifest.json`, `workbook_inventory.json` and
`validation_summary.json`. Generated evidence remains ignored, not staged.

No commit, staging or push is performed. Existing production and regression
changes remain pending; cleanup did not add to or alter those changes.
HEAD remains `6112f34fd6857f42c3bf78a3114c3bc80ea5b6cc`: 22 previously modified
tracked files, 17 untracked files after cleanup, no tracked deletions. Cleanup is
ready for review/staging with the full-run temporary-path limitation documented;
it must not be presented as a failure-free full-suite invocation.
