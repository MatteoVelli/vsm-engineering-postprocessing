# Formal Input/Output Specification — Phase 1

## 1. Scope

This specification defines the smallest deterministic processing engine that can be implemented from the supplied files. It covers importing VSM-style CSV/XLSX data, channel registration, configurable math channels, configurable statistics, plotting definitions and Excel report data contracts. PowerPoint is an optional downstream consumer.

## 2. Input data contract

### Supported file types

- `.xlsx`
- `.csv`

### Required logical fields

- one time/index channel;
- one or more numeric data channels;
- a channel name for every data column;
- units where available.

### Source-workbook example

- sheet: `Sprayer_Caiman_SP_9300Kg_Hybrid`;
- header row: 3;
- unit row: 4;
- first data row: 5;
- time channel: `Track_Time`, unit `s`;
- sample interval: 1 s in the supplied file.

These row numbers are example-specific. The importer must support explicit configuration and should also offer safe auto-detection.

### Required importer validations

- file readable and sheet exists;
- non-empty header row;
- unique internal IDs even when display names duplicate;
- numeric coercion with invalid-value reporting;
- consistent row lengths;
- monotonic time;
- duplicate timestamps;
- missing values and non-finite values;
- unit availability;
- channel-count and sample-count summary.

### Channel identity model

Each channel must have:

- `channel_id`: stable unique internal identifier;
- `source_name`: original input header;
- `display_name`: configurable alias;
- `unit`;
- `source_column_index`;
- `kind`: `raw` or `math`;
- `dtype`;
- `provenance`;
- `dependencies` for math channels.

Do not use `display_name` as the primary key because the reference report contains duplicate display names.

## 3. Math-channel contract

A math channel definition must include:

- unique ID;
- display name;
- output unit;
- expression or registered deterministic function;
- ordered dependencies;
- NaN policy;
- optional description.

Expressions must reference channel IDs, not spreadsheet letters. Evaluation order must be resolved from dependencies, and circular dependencies must be rejected.

The first implementation should support arithmetic operations and selected deterministic functions only. Excel-specific row references must not leak into the processing core.

## 4. Statistics contract

Each configured statistic must include:

- target channel ID;
- operation: `rms`, `max`, `min`, `last`, `sum`, or a named deterministic derived KPI;
- output label;
- output unit;
- placement group: `top_rms`, `bottom_channel`, or `kpi_block`;
- NaN policy.

### RMS definition

Default for uniform sampling:

`RMS = sqrt(mean(x^2))`

For irregular sampling, a separate `time_weighted_rms` operation must use the time channel and a documented numerical integration rule.

## 5. Plot-definition contract

Each plot definition must include:

- plot ID and title;
- x-channel ID;
- one or more y-channel IDs;
- axis assignment for each series;
- axis labels and units;
- optional fixed limits;
- legend labels;
- output image filename;
- Excel placement order;
- optional PowerPoint inclusion flag.

The first supported plot type should be XY line/scatter, because all meaningful charts in the reference workbook use that family.

## 6. Excel output contract

The generated workbook should contain at least:

### Sheet 1 — Report

- top RMS area;
- two-row channel header: display name and unit;
- selected raw and math channels;
- data rows aligned on the common time base;
- bottom statistics directly beneath selected channels;
- KPI summary block;
- plots placed in a deterministic grid.

### Sheet 2 — Metadata

- source filename and hash;
- import timestamp;
- selected channels;
- math-channel definitions;
- statistic definitions;
- plot definitions;
- warnings and data-quality results;
- software version/configuration hash.

Math channels must be visually distinguishable from raw channels and also identifiable in metadata; colour alone is not sufficient.

## 7. Optional PowerPoint output contract

The PowerPoint generator consumes only:

- selected KPI values;
- rendered plot images;
- report title/subtitle metadata;
- optional annotations supplied in configuration.

It must not recalculate engineering values. All values must come from the deterministic processing result.

## 8. Configuration contract

The configuration must separately define:

- export channels;
- math channels;
- RMS channels;
- MAX/MIN/last/sum statistics;
- KPI block entries;
- plot definitions;
- PowerPoint selections.

A human-editable YAML or JSON file is appropriate for the first implementation. A GUI can edit the same configuration model later.

## 9. First runnable milestone acceptance criteria

The import/channel-inventory milestone is complete when it can:

1. load the supplied source workbook;
2. identify 70 channels and 1,866 samples;
3. detect `Track_Time` as the time channel;
4. preserve source names and units;
5. assign unique internal IDs;
6. identify 45 direct source channels and 25 formula-derived columns in the supplied example;
7. report zero missing cells in the sample region;
8. verify a 1-second monotonic time base;
9. export a machine-readable channel catalogue;
10. fail clearly on missing sheets, malformed headers, nonnumeric data or nonmonotonic time.
