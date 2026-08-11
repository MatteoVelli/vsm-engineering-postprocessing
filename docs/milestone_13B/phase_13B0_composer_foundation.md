# Milestone 13B.0 - Duty-Cycle Composer Foundation

## Objective

Introduce the duty-cycle composer as an optional, modular layer without altering the validated v1.2.x importer, channel, math, statistics, plotting or reporting calculations.

The first working increment deliberately builds the **row-level composition plan and provenance only**. It does not invent the unresolved road profile or generator-active Sergio field variants.

## New package

```text
src/vsm_postprocessing/duty_cycle/
    __init__.py
    models.py
    config.py
    composer.py
```

The package provides:

- strict YAML scenario loading and validation;
- typed phase, source-row, report-row and provenance models;
- deterministic construction of all output row/timestamp/phase boundaries;
- deterministic source-row alignment for source-backed phases;
- explicit unresolved-profile flags;
- source-workbook compatibility validation;
- provenance CSV export.

A CLI entry point is also available:

```text
vsm-duty-cycle-plan
```

or directly:

```powershell
python -m vsm_postprocessing.duty_cycle_cli config/duty_cycle_sergio_reference.yaml --source <source.xlsx>
```

## Sergio reference scenario

`config/duty_cycle_sergio_reference.yaml` is the runtime configuration for the 12-phase reference mission established in Milestones 13A/13A.1.

The generated composition plan contains exactly:

- 17,418 samples;
- Track_Time 0 to 17,417 s at 1 Hz;
- 12 phases;
- six field-work phases;
- five 900-s loading/opportunity-charge phases;
- one road-travel phase;
- report rows 5 through 17,422;
- source-row alignment through source row 1,870 where available.

The plan is regression-tested row-for-row against `docs/milestone_13A/source_to_report_row_mapping.csv` for structural provenance fields.

## Explicit implementation gate

The following phase profile definitions remain unavailable from the supplied field workbook and are therefore **not numerically materialised**:

- P05 - generator-active Sergio field variant;
- P06 - Sergio road profile;
- P08 - generator-active Sergio field variant;
- P10 - generator-active Sergio field variant.

The composer foundation reports these phases instead of silently synthesising or copying incorrect values.

## Next increment - 13B.1

Implement numerical materialisation for the phases whose behavior is sufficiently defined:

1. source-backed field phases;
2. cumulative time/distance/SOC offsets with explicit per-channel rules;
3. 900-s loading/opportunity-charge phases;
4. row-level numerical provenance and phase-boundary validation.

P05/P06/P08/P10 remain gated until their profile generation can be justified deterministically.
