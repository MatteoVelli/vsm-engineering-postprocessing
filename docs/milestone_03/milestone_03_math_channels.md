# Milestone 03 — Configurable Math Channels

## Objective

Introduce a deterministic calculation layer that creates clearly identified math channels from imported VSM channels without embedding client-specific formulas in Python source code.

## Why this milestone is necessary

The final tool must let an engineer configure derived channels independently from raw-channel export, statistics and plotting. Before RMS or report generation can be trusted, the project needs a calculation engine with explicit dependencies, units, provenance and numerical validation.

## Inputs

- A validated VSM `.xlsx` or `.csv` dataset supported by the import module.
- A versioned UTF-8 YAML configuration containing:
  - source channels to export;
  - constants;
  - math-channel metadata and expressions;
  - optional comparisons against existing source channels;
  - CSV output controls.

## Outputs

`outputs/math_channels/` contains:

- `dataset_with_math_channels.csv` — selected source channels followed by calculated channels;
- `math_channel_catalogue.csv` — expression, dependencies, units, provenance and comparison evidence;
- `math_validation_report.json` — deterministic comparison metrics;
- `math_manifest.json` — source hash, configuration, order and output traceability;
- `math_summary.txt` — concise human-readable run summary.

## Implemented behaviour

- Stable source `channel_id` values are used as expression identifiers.
- Configured math channels have their own unique identifiers and are explicitly marked `kind=math`.
- Definitions may depend on source channels, constants or other configured math channels.
- Dependencies are evaluated in topological order, independently of YAML display order.
- Missing dependencies and circular dependencies stop processing with actionable messages.
- Math IDs cannot collide with imported source IDs.
- Expressions are parsed through a restricted AST evaluator; Python `eval`, attribute access, imports, indexing and arbitrary function calls are not permitted.
- Every result must be a finite float64 vector with exactly one value per imported sample.
- Division by zero, invalid square roots and other non-finite results stop processing.
- Optional source comparisons calculate maximum absolute error, RMS error and mismatch count.
- Required comparison failures stop the run.

## Supported expression elements

Operators:

- `+`, `-`, `*`, `/`, `**`
- unary `+` and `-`

Functions:

- `abs(x)`
- `sqrt(x)`
- `square(x)`
- `minimum(a, b)`
- `maximum(a, b)`
- `clip(x, lower, upper)`
- `cumulative_sum(x)`
- `sample_energy_kwh(power_kw, time_s)`

`sample_energy_kwh` uses the actual positive time intervals. For the first sample, it uses the first observed interval. It then calculates `power_kw * delta_time_s / 3600`.

## Sergio workbook acceptance test

The supplied source workbook was processed with:

- 1,866 samples;
- 12 exported source channels;
- 11 calculated math channels;
- 23 output channels;
- 10 required comparisons against existing workbook calculations.

All ten comparisons passed. Nine produced exactly zero maximum absolute error. The battery-power-squared comparison produced only floating-point round-off:

```text
maximum absolute error = 9.09494701773e-13 kW²
RMS error              = 2.10544719513e-14 kW²
mismatches              = 0
```

## Engineering decisions

### The workbook's “Battery Power RMS” channel

The workbook channel is calculated as battery power squared at each sample. It is therefore represented in this milestone as `calc_battery_power_squared`, not as a true RMS statistic. Actual RMS aggregation remains a later statistics milestone.

### Auxiliary accumulated energy

The source workbook's accumulated auxiliary-energy column is not used as a validation target because the reverse-engineering phase identified suspicious accumulation behaviour. This milestone calculates a corrected deterministic accumulated channel as:

```text
cumulative_sum(sample_energy_kwh(total_auxiliary_power, time))
```

### Client-specific constants

The source workbook uses `9548.8` as the rpm·Nm-to-kW divisor. It is declared in YAML as `rpm_nm_to_kw_divisor`, rather than hard-coded inside the calculation engine.

## Automated tests

Version 0.3.0 has 19 passing tests covering:

- all previous import and channel-selection behaviour;
- dependency ordering;
- configured constants;
- energy calculation;
- circular dependency rejection;
- missing dependency diagnostics;
- unsafe expression rejection;
- non-finite result rejection;
- source-ID collision protection;
- required comparison failure;
- output and manifest traceability;
- acceptance against the supplied Sergio workbook.

## Scope intentionally excluded

This milestone does not yet implement:

- true RMS aggregation;
- MAX, MIN or last statistics;
- plots;
- Excel report generation;
- PowerPoint report generation;
- graphical user interface;
- AI-assisted result selection.

## Completion status

Milestone 03 is complete when the user obtains:

```text
19 passed
```

and `run_math_channels.ps1` reports `status: PASS` with ten passing source comparisons.
