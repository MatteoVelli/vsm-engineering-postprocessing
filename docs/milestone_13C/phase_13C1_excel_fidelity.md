# Milestone 13C.1 / 13C.1b / 13C.1c - Excel Fidelity

## Scope

This milestone upgrades the full Sergio duty-cycle Excel report only. The deterministic duty-cycle composer, math formulas, fuel integration, statistics definitions, plotting PNG path, and PowerPoint generation remain intact.

## Implemented

- Full Sergio A:BR report channel area is exported from the canonical 17,418-sample duty-cycle dataset.
- Sergio-style channel ordering is configured in `config/excel_report_duty_cycle.yaml`.
- Report-specific channel metadata marks imported workbook math columns as `math` so generated CSV imports do not lose raw/math presentation semantics.
- KPI strip expanded from 8 to 34 configured headline values.
- Bottom summary expanded to 49 configured MAX/MIN/last/SUM-style statistic IDs, placed under visible report columns where canonical channels exist.
- Battery Power RMS and Battery Heatflow RMS remain the top RMS pair.
- Battery discharge headline KPI is explicit: `max(-Battery Power)`, separate from maximum charging power `MAX(Battery Power)`.
- Native Excel scatter charts are generated for the 18 meaningful Sergio chart definitions.
- 13C.1b adds true native secondary-axis chart groups for the 12 Sergio charts that use right-side axes.
- 13C.1b moves chart anchors to the inspected Sergio top-right layout and uses readable 15 x 7.5 inch chart dimensions.
- 13C.1c fixes native chart OOXML validity by explicitly writing reciprocal real `crossAx` IDs for every axis pair.
- 13C.1c hides duplicate secondary X axes and keeps secondary Y axes crossing at `max`.
- Excel report manifest and summary now distinguish `native_excel_chart_count`, `configured_plot_count`, `plot_series_count`, and `embedded_plot_image_count`.
- PNG plot rendering is retained for PowerPoint and legacy report workflows.
- Workbook metadata remains client-safe; direct inspection found no `C:\Users` or development workspace paths in the generated workbook metadata sheet.
- Version bumped to `v1.2.8`.

## Direct Workbook Comparison

Latest generated workbook:

`outputs/end_to_end_sergio_duty_cycle/07_excel_report/vsm_engineering_report.xlsx`

Reference workbook:

`reference_files/Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx`

Observed direct inspection:

| Area | Sergio reference | Generated workbook |
| --- | ---: | ---: |
| Main samples | 17,418 | 17,418 |
| Rows | 17,424 | 17,424 |
| Data/math channels | 70, A:BR | 70, A:BR |
| Used range | A1:DN17424 | A1:DA17424 |
| Native chart objects | 20 total, 18 meaningful | 18 meaningful |
| Embedded PNG images | 0 | 0 |
| Top RMS pair | Battery Power, Battery Heatflow | Battery Power, Battery Heatflow |
| RMS merged ranges | L1:N1 and P1:S1 | L1:N1 and P1:S1 |
| Secondary-axis charts | 12 meaningful charts | 12 meaningful charts |
| Chart axes validated | Real reciprocal axis IDs | 60 axes, 0 unresolved `crossAx` references |
| Minimum chart size | 15 x 7.5 in | 15 x 7.5 in |

## Intentional Differences

The generated report preserves canonical deterministic values instead of reproducing known spreadsheet artefacts. The final generated fuel value remains approximately `39.84212 kg`; Sergio's reference workbook displays approximately `39.83806 kg`, due to the known omitted one-second loading fuel increment at a phase boundary.

The generated workbook excludes Sergio's two orphan/scratch chart objects and therefore contains 18 native charts rather than 20 total chart objects. Secondary-axis charts are implemented as true combined native scatter chart groups with right-side value axes; no series is numerically rescaled to simulate a secondary axis.

The generated workbook has a professional `Metadata` sheet with client-safe filenames and hashes. Sergio's single-sheet workbook does not include this metadata sheet.

## Validation

Run command:

```powershell
.venv\Scripts\python.exe -m vsm_postprocessing.pipeline_cli .\config\end_to_end_sergio_duty_cycle.yaml
```

Latest v1.2.8 result:

- Pipeline passed: 8/8 stages
- Samples: 17,418
- Report channels: 70
- Statistics available: 53
- Configured plots rendered for PNG/PowerPoint compatibility: 24
- Plot series rendered: 45
- Native Excel charts in workbook: 18
- Native charts with secondary axes: 12
- Chart axes validated: 60
- Unresolved `crossAx` references: 0
- Literal invalid `crossAx` references to `10` or `20`: 0
- Workbook ZIP integrity: PASS
- PowerPoint path still uses 6 configured PNG plots
- Freeze panes: B6
- Bottom summary coverage: 48 populated A:BR row-17423 cells plus Battery Power MIN on row 17424
- KPI strip: 34 configured entries from BT:DA
- Provenance: original VSM source workbook, source hash, scenario ID, external profile/reference workbook, reference hash, internal artifact, tool/config names

Focused regression:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_excel_report.py tests\test_plotting.py tests\test_pipeline_duty_cycle.py
```

Correction-pass focused tests:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_excel_fidelity_13c1b.py
```

Result: 6 passed.

Native chart validity focused tests:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_excel_chart_validity_13c1c.py
```

Result: 3 passed.

Full regression:

```powershell
.venv\Scripts\python.exe -m pytest
```

Result: 107 passed.

Renderer note: no Excel/LibreOffice command-line renderer was available in the execution shell. Visual readability was therefore checked structurally from the workbook objects and OOXML: chart anchors, chart dimensions, chart count, axis groups, right-side secondary axes, bottom X axes, non-data-table starts, and metadata leakage.
