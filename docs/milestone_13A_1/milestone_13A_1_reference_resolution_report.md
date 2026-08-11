# Milestone 13A.1 – Sergio Reference Conflict Resolution & Formal KPI Specification

## 1. Purpose

Milestone 13A established the 12-phase / 17,418-row mission structure. A second pass against Sergio's Excel, Sergio's PowerPoint, the supplied 1,866-row VSM source, and the current generated Excel/PowerPoint exposed several definition conflicts that must be resolved before a composer is implemented.

This milestone therefore freezes the **reference hierarchy, KPI semantics, time conventions and control overrides**. It also applies a small set of safe source-only reporting fixes. It does **not** implement the duty-cycle composer and does **not** redesign the validated deterministic calculations.

## 2. Reference hierarchy

The formal precedence is now:

1. **Sergio Excel (`Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx`)** – primary target for full-mission numerical values, row counts, formulas and KPI equivalence.
2. **Raw/source VSM workbook** – primary target for the 1,866-row source field-cycle profile actually supplied.
3. **Sergio PowerPoint** – scenario intent, presentation structure, nominal/rounded values and explanatory control logic.
4. **Current tool output** – regression evidence for the source-only workflow, not a numerical authority for the future full mission.

This resolves the earlier risk of trying to force contradictory Excel and PowerPoint values to match simultaneously.

## 3. Key resolved conflicts

### 3.1 Top RMS pair

Sergio's Excel uses:

- Battery Power RMS (`N2`);
- Battery Heatflow RMS (`P2`).

The source-only report previously used Battery Power RMS + Auxiliary Power RMS. Version 1.2.1 restores the Sergio-aligned pair and adds the Battery Heatflow helper channel/statistic to the source report.

### 3.2 Battery power semantics

Battery power is signed:

- negative = battery discharge;
- positive = battery charge.

Sergio's `Max Battery Power (EDU+I+A)` is **not** `MAX(M)`. It is:

`-MIN(Battery Power) = 149.18 kW`.

The future KPI layer must keep two different quantities:

- `max_charge_power_kw = MAX(M)`;
- `max_discharge_power_magnitude_kw = -MIN(M)`.

### 3.3 Range-extender logic

The default driving control remains:

- start below 40% SOC;
- stop at/above 80% SOC.

The **final field cycle has an explicit 5–80% override**. This is supported by both the PowerPoint duty-cycle table and the supplied source filename (`57-4pcSOC_5-80`). It explains why the final source cycle runs from 57.4% to 23.9383% while generator power remains zero.

P05 -> road P06 also requires state continuity: the road starts with the drive range extender already on after P05. Loading/opportunity-charge mode is a separate generator state and does not imply that the next field drive controller starts on.

### 3.4 Battery capacity metadata

Two different values serve different purposes:

- **100.0 kWh** – Sergio mission nominal capacity and SOC/energy scaling;
- **100.77 kWh** – physical pack value shown on the dedicated battery slide.

The composer must not collapse them into one variable.

### 3.5 Time semantics

The complete mission has:

- 17,418 samples;
- timestamps from 0 to 17,417 s;
- timestamp span 17,417 s;
- displayed Excel time 290.283333 min;
- PowerPoint phase-duration total 17,416 s.

Therefore `sample_count`, `timestamp_span_s`, `configured_event_duration_s` and energy-integration semantics are separate concepts.

A 900 s loading event produces 900 rows at 1 Hz. Those 900 rows span 899 s internally, while the next phase begins 900 s after the loading phase starts.

### 3.6 Fuel-energy KPI

Sergio's Excel defines:

- Generator continuous power = 80 kW;
- Fuel Energy Consumed = `80 * 12 = 960 kWh`;
- Engine Efficiency = `219.301 / 960 = 22.84%`.

The factor 12 is not explained by the supplied sources. This formula is therefore classified as **reference compatibility only**, not accepted as a deterministic physical model.

## 4. Safe v1.2.1 source-report upgrades

The validated numerical core is unchanged. The following reporting/configuration changes are included:

1. source Excel top RMS pair is now **Battery Power RMS + Battery Heatflow RMS**;
2. battery heatflow raw/helper/statistic channels are included in the default report path;
3. source PowerPoint slide previously called `Duty Cycle` is renamed **Source 74 ha Field Cycle**;
4. the electrical-energy slide uses battery heatflow RMS/max rather than the ambiguous battery charging/min-power pair;
5. UI-hashed upload filenames are stripped from client-visible PowerPoint source text;
6. visible Excel Metadata shows the original client filename and configuration basenames rather than absolute local workspace paths;
7. software/package patch version is **1.2.1**.

These changes improve client fidelity without changing source data, physics, statistics definitions or mission composition.

## 5. Still unresolved

The remaining blockers for exact 70-channel mission reconstruction are:

- provenance/source of the 2,536-row road profile;
- provenance/source of generator-on field values after activation;
- battery chemistry metadata conflict (NMC vs CALB LFP);
- unexplained `80 * 12` fuel-energy reference formula;
- rationale for manual SOC reset values and obvious Excel boundary artifacts.

They are documented in `unresolved_ambiguities_revised.md`.

## 6. Phase 13B implementation gate

The **architecture** of Milestone 13B may now start because the phase sequencing, timing semantics, controller thresholds and KPI definitions are explicit.

However, exact cell-for-cell 70-channel equivalence is still impossible from the supplied source workbook alone. The composer must therefore support explicit external/reference profiles and preserve row-level provenance. It must never invent the missing road or generator-on VSM profile.

## 7. Deliverables

- `reference_conflict_register.csv`
- `kpi_definition_registry.csv`
- `phase_time_convention.md`
- `phase_time_inventory.csv`
- `duty_cycle_phase_inventory_revised.csv`
- `cumulative_channel_logic_revised.csv`
- `proposed_duty_cycle_spec_v2.yaml`
- `unresolved_ambiguities_revised.md`

The original row-level `source_to_report_row_mapping.csv` remains valid and is not duplicated.
