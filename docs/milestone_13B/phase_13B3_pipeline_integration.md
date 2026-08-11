# Milestone 13B.3 – Full Duty-Cycle Pipeline Integration

## Goal

Connect the validated 17,418-sample `DutyCycleComposition` to the existing deterministic post-processing/reporting pipeline without redesigning the importer, math, statistics, plotting, Excel or PowerPoint calculation engines.

## Design

The pipeline retains the original raw VSM file as the source-of-record. An optional `duty_cycle:` configuration block introduces a new stage after source inspection:

1. inspect the raw 1,866-sample source;
2. load the formal duty-cycle scenario;
3. validate the external phase Profile Provider;
4. compose all 12 phases / 17,418 samples;
5. export a pipeline-safe 70-channel numeric CSV using the source channel names/units;
6. validate that the regenerated stable channel IDs are identical to the source IDs;
7. use that CSV as the processing input for all existing downstream engines.

No provenance/string columns are mixed into the numeric processing dataset. Row-level and phase-level provenance remain separate diagnostic files.

## New reference pipeline configuration

`config/end_to_end_sergio_duty_cycle.yaml`

This uses:

- raw source workbook for source-backed phases;
- `config/duty_cycle_sergio_reference.yaml` for mission structure and deterministic loading logic;
- `config/duty_cycle_profiles_sergio_reference.yaml` plus the exact SHA-256-validated Sergio reference workbook for P05/P06/P08/P10;
- existing deterministic math/statistics/plotting engines;
- dedicated full-duty-cycle Excel/PowerPoint wording configs.

## Pipeline stage contract

| Stage | Input | Expected samples | Notes |
|---|---|---:|---|
| inspection | raw VSM workbook | 1,866 | Source validation remains unchanged |
| duty_cycle | source + scenario + provider | 17,418 | Produces processing dataset + provenance |
| channel_selection | composed CSV | 17,418 | Existing engine |
| math_channels | composed CSV | 17,418 | Existing engine |
| statistics | composed CSV | 17,418 | Existing engine |
| plotting | composed CSV | 17,418 | Existing engine |
| excel_report | composed CSV + precomputed results | 17,418 | Existing layout engine |
| powerpoint_report | full-mission stats/plots | 17,418 | Existing presentation engine |

## Traceability outputs

The duty-cycle stage writes:

- `duty_cycle_dataset.csv` – numeric 70-channel processing dataset;
- `duty_cycle_provenance.csv` – one provenance row per mission sample;
- `profile_provenance.csv` – source workbook/hash/row range for P05/P06/P08/P10;
- `duty_cycle_summary.txt` – phase/sample count and core mission endpoints.

The pipeline manifest records both the original source and the effective processing input, plus hashes for the scenario, provider config and provider workbook.

## Numerical acceptance

The composed pipeline stage validates:

- samples: 17,418;
- phases: 12;
- final time: 290.283333 min;
- final distance: 114.0011 km;
- final SOC: 23.9383%;
- max speed: 62.6233 kph;
- max generator power: 80.02669 kW.

Canonical fuel remains approximately 39.84212 kg because the deterministic composer intentionally does not reproduce Sergio's one-sample loading-boundary omission (~0.00406 kg).

## Performance

The Excel stage can reuse statistics and plots already produced by preceding pipeline stages when the configurations are identical. This avoids duplicate full-mission processing. The complete 8-stage reference run completed in about 36 s in the validation environment.

## Backwards compatibility

If `duty_cycle:` is absent, the pipeline remains the original seven-stage source-cycle workflow. Existing v1.2.x configs require no changes.

## Remaining work

Milestone 13C will address Sergio report fidelity: full 70-channel layout, larger KPI block, complete meaningful chart set, native Excel chart fidelity where justified, labels/units and final presentation alignment. It must not change the deterministic mission calculations established here.
