# Milestone 08 - End-to-End Processing Pipeline

## Objective

Provide one deterministic entry point that executes the already validated VSM processing layers from source data through the Excel engineering report.

## Why this milestone is necessary

Milestones 01-07 proved each layer independently. A production-oriented client tool must not require a user to remember six commands or manually move outputs between stages. The orchestration layer makes the workflow repeatable while keeping numerical logic inside the existing modules.

## Input

`config/end_to_end_example.yaml` defines:

- the VSM source file;
- import strictness / optional import overrides;
- channel-selection configuration;
- math-channel configuration;
- general statistics configuration;
- plotting configuration;
- Excel-report-specific statistics configuration;
- Excel report configuration;
- a dedicated output root.

All relative paths are resolved relative to the pipeline YAML file.

## Stage order

1. Source inspection
2. Channel selection
3. Math channels
4. Statistics
5. Plotting
6. Excel report

Each stage keeps its own output directory under the pipeline output root.

## Traceability

The pipeline writes `pipeline_manifest.json` containing:

- source SHA-256;
- pipeline configuration SHA-256;
- SHA-256 for each subordinate configuration;
- import options;
- stage status and metrics;
- generated output paths;
- final report path.

`pipeline_summary.txt` is the concise human-readable equivalent.

## Failure behaviour

The pipeline stops on the first failed stage. It still writes a partial manifest and summary identifying the failed stage and the error. Outputs from previously successful stages remain available for diagnosis.

## Numerical behaviour

The pipeline adds no engineering calculations. It orchestrates the existing deterministic importer, channel manager, math engine, statistics engine, plotting engine and Excel generator.

## Known implementation note

In v0.8.0 the Excel generator internally invokes its statistics and plotting dependencies again. This deliberately avoids refactoring already validated modules during the orchestration milestone. If performance on full 500-700-channel datasets warrants it, a later optimization can pass precomputed results into the report layer without changing numerical definitions.
