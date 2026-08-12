# Milestone 13C.2b - Sergio PowerPoint Fidelity

## Purpose

Milestone 13C.2b upgrades the PowerPoint generator from a compact 4-slide summary into a configurable engineering presentation for the full Sergio duty-cycle workflow.

The PowerPoint remains downstream of the validated deterministic pipeline:

```text
canonical data -> statistics + Matplotlib plotting -> PowerPoint report
```

No importer, duty-cycle composer, Profile Provider, math, statistics, fuel, RMS, Excel or Matplotlib architecture changes are required for this milestone.

## Reference Deck Inventory

The Sergio reference presentation is 13.333 x 7.5 in widescreen and contains 21 slides. Direct inspection showed:

- Slide 1: cover / feasibility-study context.
- Slides 2, 6, 16, 19, 20: static system, battery, motor or engine reference content.
- Slides 3-5 and 13-15: diesel or hybrid/diesel comparison/reference material.
- Slides 7-12: repeated hybrid duty-cycle and opportunity-charging result narrative.
- Slides 17-18: motor, wheel and duty-cycle demand reference.
- Slide 21: auxiliary energy reference.

The automated deck intentionally does not recreate all 21 slides. It consolidates repeated duty-cycle result views and excludes static screenshot-like reference slides unless their content is available as structured configuration or canonical pipeline data.

## Automated Deck Structure

`config/powerpoint_report_duty_cycle.yaml` defines a 10-slide automated deck:

1. Hybrid SP Caiman Sprayer - Feasibility Study
2. System and Mission Overview
3. Duty-Cycle Executive Results
4. Duty-Cycle Vehicle Operation
5. Battery and Electrical Energy System
6. Battery Power and Energy Recovery
7. Range Extender and Generator
8. Opportunity Charging During Loading
9. Traction, EDU and Auxiliary Energy Demand
10. Simulation Summary

Appendix slide count is currently zero. Static reference appendix generation is deferred until structured source data is configured.

## Configuration-Driven Layout

The PowerPoint YAML now supports these slide types:

- `cover`
- `overview`
- `kpi_grid`
- `plot_full`
- `plot_pair`
- `conclusion`
- legacy `summary`

Each slide declares its title, subtitle, statistic IDs, plot IDs and optional body text. The Python renderer implements generic slide layouts and does not hard-code Sergio plot filenames.

## Matplotlib Figure Usage

The deck uses the professional Matplotlib PNG outputs from Milestone 13C.2a. It does not use Excel screenshots and does not reimplement plotting inside the PowerPoint module.

The duty-cycle deck uses 10 unique plot IDs with 11 image placements:

- Vehicle Speed vs Distance
- Battery SOC
- Battery Energy vs Time
- Battery Power Charge and Discharge
- Released and Recuperated Energy
- Generator Power and Fuel Consumption
- Engine Power and Fuel Consumption vs Time
- Agrochemical Discharge and Battery SOC vs Time
- Wheel and EDU Power
- Tyre Rolling Resistance Energy

## KPI Presentation

PowerPoint consumes `StatisticResult` objects from the validated statistics engine. It does not recalculate KPI values. Engineering values are formatted for presentation with sensible unit display and reduced precision.

Major canonical values retained:

- Distance: 114.00 km
- Time: 290.28 min
- Final SOC: 23.94 %
- Fuel consumption: 39.84 kg
- Max total generator power: 80.03 kW

Known Sergio numerical differences remain intentionally preserved in favor of canonical pipeline results.

## Client-Safe Metadata

Visible slide text is restricted to presentation titles, configured body text, KPI values, footer/version text and slide numbers. Local development paths and SHA-256 hashes are not shown on slides.

Document properties are populated with:

- Title
- Subject
- Author
- Keywords
- Comments

Detailed source/config hashes remain in external pipeline/report manifests.

## Footer

The footer format is:

```text
VSM Engineering Post-Processing Tool | v<version>
```

Slide numbering is shown as:

```text
current/total
```

## Validation

Automated checks cover:

- PPTX generation and ZIP integrity
- 10-slide duty-cycle deck structure
- slide title order and no duplicate unintended titles
- document properties
- major formatted KPI values
- footer version and slide numbering
- 11 image placements / 10 unique Matplotlib media files
- existing image relationship targets
- no visible absolute path leakage
- source-cycle 4-slide PowerPoint compatibility
- native Excel chart architecture remains separate

## Remaining Visual Difference From Sergio

The generated deck is not pixel-perfect against Sergio's presentation. It intentionally uses a cleaner generated-report style, professional Matplotlib figures and canonical deterministic KPI values. It does not copy Sergio's static screenshots, diesel comparison slides or repeated near-duplicate duty-cycle slides.

No local slide renderer was available during implementation, so structural PPTX checks were completed in code. Final Microsoft PowerPoint visual inspection remains required before calling v1.3.0 final.
