# Changelog

## 1.2.1 - Sergio Fidelity Preparation Patch

- Formalized reference precedence: Sergio Excel for full-mission numerics, raw VSM for supplied source profile, PowerPoint for scenario/presentation intent.
- Added Milestone 13A.1 conflict register, formal KPI registry, time convention, revised phase inventory and v2 duty-cycle specification.
- Added explicit 5–80% SOC final-cycle range-extender override alongside the normal 40–80% drive controller.
- Separated 100.0 kWh mission nominal battery scaling from the 100.77 kWh physical pack reference.
- Corrected the default source-report top RMS pair to Battery Power RMS + Battery Heatflow RMS.
- Added battery heatflow helper/statistics to the default source report.
- Renamed the source-only PowerPoint duty-cycle slide to `Source 74 ha Field Cycle`.
- Sanitized client-visible upload filenames/config paths while retaining internal SHA-256 traceability.
- Classified Sergio's `Fuel Energy = 80 * 12` and dependent engine-efficiency KPI as reference compatibility only.
- No duty-cycle composer or AI functionality has been added.

## Milestone 13A baseline - Sergio Reference Fidelity & Duty-Cycle Reverse Engineering

- Added a complete reverse-engineering report for Sergio's 17,418-sample hybrid duty cycle.
- Added a 12-phase duty-cycle inventory covering six field-work phases, one road phase, and five 900 s loading/opportunity-charging phases.
- Added per-row source-to-report provenance mapping for all 17,418 reference samples.
- Added cumulative/derived channel logic inventory and documented reference-workbook quirks.
- Added a proposed configuration-oriented duty-cycle specification for the future optional composer.
- No duty-cycle composer has been implemented yet; the validated v1.2.0 deterministic processing workflow remains unchanged.
- AI functionality remains out of scope.

## 1.2.0 - Final Deterministic Release

- Added deterministic client ZIP builder and external SHA-256 checksum.
- Added internal `RELEASE_MANIFEST.json` with per-file hashes.
- Added release/version consistency regression tests.
- Added `.python-version` for Python 3.11 project consistency.
- Added GitHub Actions regression workflow.
- Added configuration guide, known limitations, final release notes and final acceptance checklist.
- Added environment package snapshot command for release traceability.
- Final client package explicitly excludes private reference workbooks, generated outputs, tests and development history.

## 1.1.0 - Client Readiness

- Added `vsm-doctor` environment/configuration health check.
- Added `START_VSM_TOOL.bat` one-click Windows client launcher.
- Added pipeline execution log with stage-level timestamps and durations.
- Added software/Python/platform metadata to the pipeline manifest.
- Added atomic writing for top-level pipeline manifest/summary/log metadata.
- Pipeline now preserves diagnostics for unexpected Python/runtime exceptions.
- Safer cleaning: only known generated pipeline artifacts are removed; unknown files in the output root are preserved.
- Added client quick-start, troubleshooting and release-checklist documentation.

## 1.0.0

- Deterministic optional PowerPoint report generation.
- Complete 7-stage end-to-end workflow and Streamlit UI.
