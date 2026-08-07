# Milestone 06 — Excel Report Generator

## Goal

Create the first deterministic Excel engineering report from the processing layers already validated in Milestones 01–05.

## Why this milestone is necessary

Sergio's reference workbook combines selected channels, math channels, RMS information, per-channel summary statistics, KPI values and plots. Until this milestone those outputs existed as separate CSV/JSON/PNG artifacts. Milestone 06 composes them into one traceable workbook without moving numerical calculations into Excel.

## Inputs

- VSM `.xlsx` or `.csv` source file;
- math-channel YAML;
- statistics YAML;
- plotting YAML;
- Excel-report YAML.

The report generator consumes deterministic results produced by the existing engines. It does not recalculate engineering quantities using spreadsheet formulas.

## Workbook structure

### `Report`

- rows 1–2: report title/source and configured RMS strip;
- rows 3–4: selected channel display names plus explicit `RAW | unit` / `MATH | unit` identification;
- row 5 onward: aligned numeric samples;
- KPI strip: right of the exported channel area;
- bottom statistic rows: one explicit row per operation (`MAX`, `MIN`, `LAST`, `SUM`, `RMS`, `TIME-WEIGHTED RMS`);
- configured plots: embedded beneath the data/statistic area in a deterministic grid.

The separate operation rows deliberately replace the ambiguous mixed summary row found in the client reference workbook.

### `Metadata`

Contains source identity/hash information, time-base information, exported channel metadata, statistics definitions/results and plot provenance.

## Configuration separation

`config/excel_report_example.yaml` independently controls:

- report channels;
- top RMS entries;
- KPI entries;
- bottom statistic operations;
- plot inclusion and layout;
- output filename and retention of plot assets.

This preserves the project requirement that export, math, statistics and plotting selections remain independently configurable.

## Acceptance criteria

For the supplied source workbook the example configuration must produce:

- 1,866 samples;
- 19 report channels;
- 14 available configured statistics;
- 6 embedded plots;
- `Report` and `Metadata` sheets;
- battery RMS of approximately `64.9480711679 kW`;
- max speed of `30.0717 kph`;
- minimum battery power of `-149.18 kW`;
- final vehicle distance of `12000 m`;
- preserved traceability files alongside the workbook.

## Out of scope

- GUI;
- scenario/duty-cycle construction;
- PowerPoint generation;
- AI selection/recommendations;
- pixel-perfect reproduction of manual formatting artefacts in the Sergio workbook.
