# Milestone 04 — Statistics Engine

## Objective

Add a deterministic, configurable statistics layer that can consume both imported VSM channels and newly calculated math channels. The outputs must be traceable and structured for the future Excel report generator.

## Why this milestone is necessary

The supplied Excel report uses different summary operations for different channels: RMS, MAX, MIN, last and SUM. Those operations cannot be inferred safely from units or names. They therefore require explicit per-channel configuration.

The supplied report also places results in three different areas:

- RMS values in the upper rows;
- per-channel summaries below the data;
- derived KPI values in a separate block.

The statistics engine records these intended placements without generating the formatted workbook yet.

## Inputs

- validated XLSX or CSV source data;
- optional math-channel YAML configuration;
- versioned statistics YAML configuration;
- optional explicit import bounds for workbooks that contain data and summaries in the same sheet.

## Supported operations

### RMS

```text
sqrt(mean(x^2))
```

This is the ordinary sample RMS definition.

### Time-weighted RMS

```text
sqrt(trapezoidal_integral(x^2 dt) / elapsed_duration)
```

This operation is separate and requires a finite, strictly increasing time channel.

### Other operations

```text
max
min
last
sum
```

## Non-finite values

Each statistic declares one policy:

- `error`: reject the statistic if any non-finite sample exists;
- `omit`: exclude non-finite samples and record how many were omitted;
- `propagate`: return NaN explicitly.

The default policy is `error`.

## Placement groups

Every statistic declares one future report placement:

```text
top_rms
bottom_channel
kpi_block
```

No formatting assumptions are embedded in the numerical engine.

## Excel reference validation

The report workbook is imported with explicit bounds:

```text
Header row: 3
Unit row: 4
Data rows: 5–17422
Data columns: A–BR
```

This prevents row 17423, row 17424 and columns BT–DM from being misinterpreted as simulation data.

Thirty-one configured statistics are compared against cached values already present in the workbook. All comparisons pass.

For ordinary RMS, the engine does not reproduce the workbook's ambiguous denominator. The workbook uses elapsed seconds (`17417`) while the dataset contains `17418` samples. Explicit tolerances document the resulting small differences.

## Outputs

```text
statistics_results.csv
statistics_by_channel.csv
statistics_validation_report.json
statistics_manifest.json
statistics_summary.txt
```

`statistics_results.csv` is the primary long-form contract for future Excel and PowerPoint generation. It preserves the statistic ID, channel ID, channel type, operation, unit, value, placement group, NaN policy and comparison evidence.

## Acceptance results

### Source workbook

```text
Samples: 1866
Imported channels: 70
New math channels available: 11
Total channels available to statistics: 81
Statistics calculated: 14
```

### Reference report workbook

```text
Samples: 17418
Channels: 70
Statistics calculated: 31
Reference comparisons passed: 31/31
```

### Automated tests

```text
30 passed
```

## Milestone status

Milestone 04 is complete.

The next milestone should implement deterministic plot generation from configured channel selections. Excel report generation should follow only after the plotting outputs and layout contracts are stable.
