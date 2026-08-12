# Changelog

## 1.2.10 - Sergio PowerPoint Fidelity

- Upgraded the duty-cycle PowerPoint configuration from 4 compact slides to a 10-slide engineering feasibility-study deck.
- Added configurable cover, overview, KPI-grid, full-chart, plot-pair and conclusion slide layouts.
- Routed PowerPoint charts through the validated professional Matplotlib PNG outputs.
- Added professional document properties, footer versioning and `current/total` slide numbering.
- Added client-safe visible text handling with no local path display on generated slides.
- Added PowerPoint manifest metadata for slide titles, displayed KPI count, plot files, image placements, appendix count and document properties.
- Added Sergio PowerPoint reference-difference documentation and structural PPTX regression checks.

## 1.2.9 - Professional Matplotlib Engineering Plotting

- Added a config-driven engineering plotting style layer for publication-quality Matplotlib figures.
- Added high-resolution PNG output under `png/` plus optional SVG output under `svg/`.
- Added true Matplotlib secondary-axis rendering with combined legends and human-readable labels.
- Added engineering tick formatting, zero-line behavior and consistent typography/grid styling.
- Added optional duty-cycle phase-boundary overlays from row-level provenance.
- Added plotting metadata for dimensions, DPI, axes, SVG paths and phase-aware plot counts.
- Preserved the native Excel chart renderer as the Excel-only interactive chart path.
- Added focused plotting tests for style loading, output dimensions, SVG generation, labels, phase mode and figure cleanup.

## 1.2.5 - Full Duty-Cycle Pipeline Integration

- Added optional `duty_cycle:` stage to the standard end-to-end pipeline configuration.
- Added pipeline-safe 70-channel duty-cycle CSV export that preserves stable source channel IDs and units.
- Routed the complete 17,418-sample mission through the existing channel-selection, math, statistics, plotting, Excel and PowerPoint engines.
- Added row-level duty-cycle provenance, external-profile provenance and mission summary outputs to normal pipeline runs.
- Added raw-source vs processing-input traceability and duty-cycle config/workbook hashes to the pipeline manifest.
- Added dedicated full-mission Excel/PowerPoint presentation configs with correct full-duty-cycle wording.
- Added reuse of already-calculated statistics and plots during Excel generation to avoid expensive duplicate full-mission processing.
- Added full 8-stage 17,418-sample regression acceptance coverage while preserving the original 7-stage source-cycle pipeline.
- No AI functionality added; full Sergio 70-channel/report-layout fidelity remains Milestone 13C.

## 1.2.4 - Duty-Cycle External Profile Provider

- Added a modular `PhaseProfileProvider` contract separate from the duty-cycle composer.
- Added SHA-256-validated workbook phase-profile provider with explicit per-phase row mappings.
- Added strict positional-channel alignment mode for the fixed Sergio 70-column reference layout; positional alignment requires the exact configured workbook hash.
- Added row-level external-provider provenance and phase-level `PhaseProfileProvenance`.
- Added complete `compose_duty_cycle(...)` materialisation for all 12 phases / 17,418 samples when the missing four profiles are explicitly provided.
- Added complete duty-cycle CSV export with generation/provider provenance columns.
- Added CLI support for `--profile-config`, `--profile-workbook` and `--materialize-full`.
- Added Sergio reference-provider config for P05/P06/P08/P10 without embedding client workbooks in the repository.
- Preserved canonical deterministic differences instead of copying known spreadsheet boundary artefacts.
- Added per-channel full-mission reference-difference inventory and provider-phase inventory.
- No AI functionality added; the Streamlit/end-to-end report pipeline is not yet switched to the composed full mission.

## 1.2.3 - Duty-Cycle Supported Numerical Prefix

- Added deterministic numerical materialisation for P01-P04 using source-backed field phases and synthetic loading/opportunity charging.
- Added source-phase mission offsets, cumulative-channel continuity and deterministic derived-channel recomputation.
- Proved generator-active P05/P08/P10 and road P06 cannot be recreated by a simple generator overlay on the supplied field trace.
- Preserved a hard stop before unresolved P05 rather than filling missing phases with guessed values.

## 1.2.2 - Duty-Cycle Composer Foundation

- Added modular `duty_cycle/` models, configuration loader and composition-plan foundation.
- Reconstructed the exact 17,418-row structural phase/timestamp/report-row provenance plan from YAML.
- Added source-workbook compatibility validation and deterministic provenance export.

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
