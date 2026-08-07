# Milestone 02 — Configurable Channel Management

## Objective

Provide a deterministic and testable mechanism for selecting a subset of VSM channels without modifying Python source code.

## Why this milestone is necessary

The final tool must handle hundreds of available channels while exporting only approximately 30–50 selected channels. Source display names cannot be treated as unique identifiers because the supplied workbook already contains a duplicated name. Selection therefore requires stable, traceable channel identifiers.

## Inputs

1. A validated `.xlsx` or `.csv` VSM dataset.
2. A versioned YAML configuration.
3. Optional explicit import-layout overrides.

## Outputs

- selected numeric CSV data;
- selected-channel catalogue;
- JSON selection manifest;
- human-readable selection summary.

## Implemented behaviour

- Selection is performed only by `channel_id`.
- The time channel can be automatically inserted first.
- Configuration order is preserved exactly.
- Duplicate requested IDs are rejected.
- Missing IDs stop processing with suggestions.
- Unknown YAML keys are rejected.
- Output filenames cannot escape the output directory.
- The source SHA-256 and channel provenance are retained.
- No calculations are performed or altered in this milestone.

## Acceptance result on the supplied source workbook

- samples: 1,866;
- available channels: 70;
- selected channels in example: 12;
- time channel: `track_time__col_001`;
- time range: 0–1,865 s;
- nominal time step: 1 s;
- automated tests: 10 passed.

## Known boundary

Formula-derived columns already stored in the source workbook are selected using their cached numeric values. New configurable math-channel calculation is intentionally deferred to the next processing milestone.
