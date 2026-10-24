# Historical Agro reporting verification

These records describe successive implementation stages, not the final client
interface. Their engineering findings, assumptions and historical test results
are preserved. The current implementation record is
[Agro Source Feature Removal](../../AGRO_SOURCE_FEATURE_REMOVAL_2026_10_24.txt);
the original machine specification/root-cause audit remains
[Machine Reporting Audit](../../MACHINE_REPORTING_AUDIT_2026_10_23.txt).

- [Source policy](AGRO_SOURCE_POLICY_2026_10_23.txt): general source resolution,
  safeguards and provenance; supersedes the initial capacity-availability decision.
- [Statistics formulas](AGRO_STATISTICS_FORMULAS_2026_10_23.txt): live Value formulas
  and native Excel evidence before removal of the dedicated source worksheet.
- [Partial template simplification](TEMPLATE_AGRO_WORKFLOW_2026_10_24.txt): an
  intermediate state superseded by the final removal record.

Paths and commands inside these forensic records are historical repository-root
paths. Older disposable pytest workspaces have been removed. Existing accepted
reference reports, numerical baselines, final/native evidence and cited logs
remain at their original paths. The two superseded Python wrappers mentioned in
these records (`scripts/audit_agro_source_policy.py` and
`scripts/validate_agro_statistics_formulas.py`) were archived with one-off staging
scripts and intermediate logs in the local, ignored
`outputs/repository_cleanup_2026_10_24/historical_verification.zip`.

For current reference regeneration use `scripts/validate_template_agro_workflow.py`
(which imports `scripts/audit_machine_specifications.py`); for native Excel use
`scripts/validate_agro_statistics_formulas_excel.ps1 -AuditPath <native_audit.json>`.
These private-reference tools still require the saved baseline artifacts and
original UI comparison sets documented in the
[cleanup audit](../../REPOSITORY_CLEANUP_2026_10_24.md). They are not clean-checkout
or client startup prerequisites.
