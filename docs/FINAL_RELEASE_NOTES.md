# Deterministic Release Notes - v1.2.10

## Purpose

Version 1.2.10 focuses the production workflow on profile-driven RoboSprayer reporting. The Electric and Hybrid profiles are aligned with Sergio's latest Electric_03 and Hybrid_04 templates, and the removed composed-mission scenario/provider path is no longer exposed in code, UI, scripts, or release packaging.

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
