# Milestone 05 — Plotting Engine

## Objective

Add deterministic, configuration-driven plot generation for imported VSM channels and newly calculated math channels. Plot files and metadata must be suitable for reuse by the future Excel and PowerPoint report generators.

## Why this milestone is necessary

The client report contains a repeatable set of engineering plots. Plot selection must remain independent from channel export, math-channel calculation and statistics. Hard-coding chart definitions in Python would make the tool difficult to maintain, so chart composition is declared in versioned YAML.

## Inputs

- validated XLSX or CSV source data;
- versioned plotting YAML configuration;
- optional math-channel YAML configuration;
- optional explicit import bounds for report workbooks containing data plus summary rows.

## Plot contract

Each configured plot declares:

- stable `plot_id`;
- title;
- one explicit X channel;
- one or more Y-series channel IDs;
- primary or secondary Y-axis assignment per series;
- optional explicit axis labels;
- deterministic PNG filename;
- optional Excel reference chart number.

The plotting engine creates one independent PNG per configured plot. It does not silently resample, smooth, interpolate, scale or otherwise transform engineering values.

## Channel handling

Plots can use:

- imported raw channels;
- imported Excel-calculated channels;
- newly calculated math channels produced by the deterministic math engine.

All lookups use stable `channel_id` values. Missing IDs stop execution with close-match suggestions.

## Numerical and data-quality rules

Every X and Y series must be a non-empty one-dimensional numeric array containing only finite values. A non-finite value stops plot generation and identifies the first affected sample index.

## Axis and unit handling

- X-axis labels default to the configured X channel name and unit.
- A single Y series defaults to that channel name and unit.
- Multiple Y series sharing one unit receive a common unit label.
- Mixed engineering quantities can be separated explicitly onto a secondary Y-axis.

## Reference Excel mapping

Phase 1 identified 20 embedded Excel charts. Eighteen are meaningful engineering charts and two are incomplete orphan/scratch charts.

`config/plotting_reference_report.yaml` maps all 18 meaningful charts:

```text
1–13, 16–20
```

The two incomplete charts, numbers 14 and 15, are intentionally excluded rather than reproduced.

Acceptance validation confirms:

```text
Reference samples: 17418
Meaningful plots mapped: 18
Series rendered: 34
Reference chart numbers covered: 1–13, 16–20
```

The mapping validates chart titles, chart numbers, series counts and channel resolvability against the reverse-engineered inventory. Pixel-for-pixel styling equivalence is not a requirement at this stage; the goal is a correct reusable engineering plotting layer.

## Source-workbook example

`config/plotting_example.yaml` renders six useful plots directly from the original VSM example and calculated math channels:

```text
Speed Vs Time
Battery SOC Vs Distance
Battery Power Charge-Discharge
Engine Power Required & Fuel Consumption Vs Distance
Auxiliaries Energy Consumption
Power at Wheels and EDU
```

Acceptance result:

```text
Samples: 1866
Available channels: 81
Plots rendered: 6
Series rendered: 11
```

## Outputs

For each run the output directory contains the configured PNG files plus:

```text
plot_catalogue.csv
plot_manifest.json
plotting_summary.txt
```

The manifest records the source hash, configuration hash, dimensions, DPI, plot definitions, series IDs and output paths for traceability.

## Automated tests

The plotting test suite covers:

- PNG generation;
- primary and secondary axes;
- missing-channel diagnostics;
- duplicate plot IDs and filenames;
- invalid axis definitions;
- invalid output paths;
- rejection of unknown YAML fields;
- complete mapping of the 18 meaningful Excel charts;
- source-workbook plotting acceptance;
- report-workbook plotting acceptance.

With both client workbooks present, the complete repository test suite contains 41 passing tests.

## Milestone status

Milestone 05 is complete once the source and reference plotting commands both pass locally and produce their PNG/metadata outputs.

The next milestone should implement Excel report generation using the already validated channel, math, statistics and plotting contracts.
