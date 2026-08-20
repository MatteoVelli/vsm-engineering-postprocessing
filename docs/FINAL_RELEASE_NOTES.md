# Deterministic Release Notes - v1.3.0

## Purpose

Version 1.3.0 focuses the production workflow on dynamic vehicle reporting. Electric and Hybrid profiles remain the analysis definitions, while machine identity is resolved from source metadata/filenames and can be corrected in the UI before Excel and PowerPoint generation.

No AI functionality is included in engineering calculations.

## Validated Source Workflow

```text
VSM CSV/XLSX
    -> inspection
    -> profile channel mapping
    -> math channels
    -> statistics and KPIs
    -> plots
    -> Excel report
    -> optional PowerPoint report
```

## v1.3.0 Changes

- Dynamic report metadata resolves RoboSprayer Electric and Caiman SP Hybrid report identities from source filenames.
- Excel and PowerPoint titles, metadata and output filenames use the resolved machine/powertrain identity.
- Road Profile plotting includes Road Gradient and optional Road Height when `Track_Height` is present.
- Excel plot-section headers now use the established dark blue fill with white bold text, and profile charts are larger for review.
- Hybrid PowerPoint KPI cards now render exactly the available four-card set without an empty fifth card.
- Nominal battery capacity is inferred from resolved Battery Energy and SOC samples as 100% SOC capacity; initial SOC is not treated as a fixed 95% or as an operational SOC ceiling.
- Max Battery Charging Power reports the positive charging side of Battery Power and resolves to 0 kW when no positive charging samples exist.

## v1.2.10 Changes

- Electric profile now uses `Robo_Sprayer_Electrification_Tamplate_Electric_03.xlsx`.
- Hybrid profile now uses `Robo_Sprayer_Electrification_Tamplate_Hybrid_04.xlsx`.
- Optional `Track_Height` is mapped to `Road Height` and exported immediately after `Road Gradient` when present.
- Missing optional raw/math/report channels no longer invalidate profile report generation.
- Agrochemical Discharge remains `-HitchRear_Force_Z_VehicleCoordinates / 9.81`.
- PowerPoint KPI label slots now update with the same statistics as the value slots.
- The real Astauto logo from `reference_files/astauto-light-text_web.jpg` is added to every generated profile PowerPoint slide.
- Retired scenario/provider CLI, configs, UI branch, assets, and tests have been removed.

## Client Data Policy

Reference Excel/PowerPoint files are validation inputs and are not bundled in client release ZIPs unless explicit authorization is provided.
