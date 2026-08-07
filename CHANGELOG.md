# Changelog

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
