# Deterministic Release Notes - v1.2.5

## Purpose

Version 1.2.5 integrates the complete 17,418-row duty-cycle composition into the normal deterministic reporting pipeline. The Profile Provider remains explicit and traceable for the four missing numerical phases, while the existing channel-selection, math, statistics, plotting, Excel and PowerPoint engines now operate directly on the composed mission.

The source-only production workflow remains unchanged and no AI functionality is included.

## Validated source workflow

```text
VSM CSV/XLSX
    -> inspection
    -> channel selection
    -> math channels
    -> statistics
    -> plots
    -> Excel report
    -> optional PowerPoint report
```

## Duty-cycle subsystem

```text
source field VSM
    + scenario YAML
    + optional external phase provider
        -> 17,418-row duty-cycle composition
        -> row/phase provenance
        -> pipeline-safe 70-channel CSV
        -> channel selection / math / statistics / plots
        -> Excel / optional PowerPoint
```

Native/synthetic logic materialises P01-P04, P07, P09, P11 and P12. The Sergio reference provider supplies P05, P06, P08 and P10 because their generator-active/road dynamics are absent from the supplied source field workbook.

## v1.2.5 changes

- optional duty-cycle stage in the standard end-to-end pipeline;
- pipeline-safe numeric export preserving all 70 stable source channel IDs and units;
- full 17,418-sample mission routed through existing deterministic channel/math/statistics/plot/report engines;
- raw-source and effective-processing-input hashes retained separately in the pipeline manifest;
- row-level and external-profile provenance exported with normal pipeline outputs;
- full-duty-cycle Excel/PowerPoint wording configs;
- reuse of precomputed statistics/plots during Excel generation to avoid duplicate full-mission calculation/rendering;
- original seven-stage source-only pipeline remains unchanged when `duty_cycle:` is absent;
- Sergio Profile Provider architecture from v1.2.4 remains the explicit fidelity bridge for P05/P06/P08/P10.

## Numerical acceptance

With the authoritative Sergio reference workbook supplied as the external fidelity provider, the complete composition reaches:

- 17,418 samples;
- final Track_Time 17,417 s;
- final Time 290.283333 min;
- final Distance 114.0011 km;
- final Battery SOC 23.9383%;
- maximum Speed 62.6233 kph;
- maximum Generator Power 80.0266906 kW.

The canonical final fuel value is 39.84212 kg, 0.00406 kg above Sergio's Excel value. The difference is intentional: the deterministic loading action integrates every configured one-second fuel increment instead of reproducing the spreadsheet boundary omission.

## Reference policy

Sergio's Excel workbook is the primary numerical target for the full mission. The raw VSM workbook is authoritative for the supplied 1,866-row source field cycle. The PowerPoint is used for scenario intent and presentation evidence where it does not conflict with the Excel.

The Sergio external provider is a **reference-fidelity bridge**, not a claim that the missing road/range-extender vehicle physics have been inferred. Independent VSM profiles can replace it through the same provider interface.

## Client data policy

The Sergio reference Excel/PowerPoint files are validation inputs and are never bundled in repository/client release ZIPs unless explicit authorisation is provided.

## Remaining boundary before 13C

The config-driven end-to-end pipeline can now generate compact Excel and PowerPoint reports from the full 17,418-sample mission. The Streamlit UI still exposes the generic uploaded-source workflow and does not yet present a dedicated Sergio duty-cycle mode. Full 70-channel Excel layout, complete KPI block, native/chart fidelity and reference-style presentation remain Milestone 13C.
