# Deterministic Release Notes - v1.3.0

Final Sergio delivery, 17 October 2026: the Electric total-tyre page now says
"absorbed power and accumulated energy versus time" and its energy legend says
"Accumulated Tyre Resistance Energy". Existing VSM source mappings, ASTAUTO
display labels and workbook-template conventions are preserved. Display-label
differences remain a separate review topic. Values, integration and statistics
are unchanged.
The approved covers, 15-slide Electric order and Diesel torque/throttle
correction are preserved. Setup remains Python >=3.11,<3.12 with CPython 3.11.9
bootstrap support. Latest-client Hybrid acceptance still requires its source;
available reference and controlled cases establish compatibility.

## Purpose

Version 1.3.0 provides dynamic vehicle reporting for Diesel, Hybrid and Electric. Machine identity is resolved from source metadata/filenames and can be corrected in the UI before Excel and PowerPoint generation.

No AI functionality is included in engineering calculations.

## October 2026 Reporting Refinements

- Tyre-Soil Bulldozing and Tyre-Soil Compaction each have six KPI cards matching Wheel Steering Angles: signed maximum power, time-weighted mean power, sample RMS power, net accumulated energy, energy per kilometre and the existing distance KPI. The original plots retain their data, axes and dimensions below the cards. Excel calculations/layouts and all other slides are unchanged. Missing inputs show N/A with reasons in the presentation manifest.

- Every exported mathematical channel has an authoritative equation note on its main Excel heading; the mapping worksheet remains available.
- Current profiles use live arithmetic, signed observed-interval trapezoidal energy, cumulative and statistic formulas. Every applicable configured KPI Value cell is a real formula, with documented parameters and formula helpers on the hidden Formula Support sheet. Python supplies validated initial caches and PowerPoint values.
- Hybrid and Electric upper RMS cells use native formulas matching sample RMS and the configured missing-data rules. Diesel RMS is unchanged.
- Diesel Auxiliary Torque headings and Auxiliary Power chart labels reflect the existing source channels.
- Current Diesel reports provide 20 native charts when their inputs are available. The first row is 24 pt; charts start below frozen headers. The existing thermal slide contains separate oil/water plots and four MIN/MAX cards, preserving the ten-slide complete-data presentation.
- Electric and Hybrid use measured gearbox-output torque and exact SI wheel-power conversion, channel-specific confirmed HV kW interpretation, signed auxiliary energy, soil-resistance power/energy and optional road height. Complete Electric reports now provide 25 charts, 12 KPI formulas and 15 slides, including the existing auxiliary-energy chart and total tyre resistance power/energy before the soil contribution pages.
- Current Electric and Hybrid profiles add four separate plots versus Track Time [s]: Total Bulldozing Resistance Power [kW], Tyre Bulldozing Resistance Energy Accumulated [kWh], Total Compaction Resistance Power [kW] and Tyre Compaction Resistance Energy Accumulated [kWh]. Existing four-wheel sums and signed cumulative trapezoidal energy are reused without any calculation changes. Existing combined soil charts remain; two Astauto plot-pair slides are added before the summary. Missing/invalid inputs disable the affected plots explicitly; measured zeros are retained.
- The dated portable delivery is `Astauto_VSM_October_16102026` (folder and ZIP root), with the existing setup/start launchers beside `Tool`. Legacy and Diesel plot/slide inventories are unchanged.
- Streamlit channel labels carry stable unique IDs. Duplicate selections remain errors; optional missing measurements disable dependent content without becoming fabricated zeros.
- Diesel `_07` and Electric UoL `_04` real VSM inputs are used for release validation. Available Hybrid reference/controlled integrations pass; the latest updated real Hybrid CSV has not been individually acceptance-tested and is not a release gate. This is a validation boundary, not a known software defect.
- Legacy profiles retain their original conventions. Visible legacy MATH values may remain Python-calculated; KPI formula helpers keep the dependency chain live. Regenerate after source changes to refresh retained values, charts and PowerPoint.
- Application version and Python 3.11 runtime requirements are unchanged.
- Tested Streamlit is 1.49.1; supported dependency constraints are recorded in `pyproject.toml`. Automatic Windows bootstrap targets CPython 3.11.9 and creates a package-local environment. No MATLAB or Office runtime is required to generate reports.

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

Only the required branding and report-layout templates are bundled as runtime resources. Private simulation datasets and generated validation reports are excluded from client release ZIPs.
