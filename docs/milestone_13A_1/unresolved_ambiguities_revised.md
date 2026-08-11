# Phase 13A.1 – Remaining unresolved ambiguities

The mission structure and KPI semantics are now sufficiently defined to begin the **architecture** of Milestone 13B. The following items remain genuinely unresolved and must not be replaced by guesses.

## 1. Road source/profile provenance

Sergio's Excel contains a 2,536-row road phase spanning **42.0087 km** and reaching **62.6233 kph**. The supplied VSM source contains only the ~30 kph field cycle. We still do not know whether the road profile came from:

- a separate 60 kph VSM run;
- another workbook/revision;
- a manually constructed profile.

**Policy:** treat the road as an explicit reference/external profile until the source is known.

## 2. Generator-on field profile provenance

Fields P05, P08 and P10 initially align with the supplied source but diverge when the drive range extender activates near 40% SOC. Their post-activation speed/pedal/electrical/engine/wheel values cannot be derived cell-for-cell from the supplied no-generator field source.

**Policy:** do not synthesize those 70-channel values without a documented model or source profile.

## 3. Battery chemistry metadata

The PowerPoint is internally inconsistent:

- duty-cycle table: **NMC 1C+2G**;
- cover/dedicated battery slide: **CALB LFP 163 Ah**.

This does not currently affect deterministic calculations, so chemistry remains configurable/unresolved.

## 4. Fuel Energy Consumed = 960 kWh

Sergio's Excel computes:

`80 kW × 12 = 960 kWh`

and then engine efficiency as engine delivered energy / 960. The meaning and unit basis of the factor **12** is not established by the supplied files.

The fuel-mass logic, by contrast, is well supported by the 1700 rpm generator BSFC/fuel-flow information.

**Policy:** 960 kWh and the derived 22.8% efficiency may be reproduced only as clearly labelled reference-compatibility KPIs.

## 5. Exact rationale for manual SOC resets

The Excel deliberately starts phases at rounded/scenario SOC values such as 77%, 59%, 42%, 41%, 58% and 57.4%. These are not always perfectly continuous with the preceding row.

**Policy:** phase initial SOC is a scenario input with provenance, not an inferred physical reset rule.

## 6. Reference boundary artifacts

A small number of Excel cells remain obvious copy/edit artifacts:

- distance discontinuities at phase boundaries;
- first loading rows with mismatched BO/BR generator state;
- inherited fuel/auxiliary states;
- shifted AL formulas in field 5;
- inconsistent AX cumulative start references.

**Policy:** preserve these only in a difference log, not as deterministic physics.

## Resolved in 13A.1

The following are no longer ambiguities:

- numerical authority hierarchy: **Excel > PPT for full-mission numerical targets**;
- top RMS selection: **Battery Power RMS + Battery Heatflow RMS**;
- Sergio `Max Battery Power`: **discharge magnitude = -MIN(Battery Power)**;
- default drive range-extender thresholds: **40% start / 80% stop**;
- final-cycle override: **5% start / 80% stop**;
- mission nominal battery scaling: **100.0 kWh**, distinct from **100.77 kWh physical pack reference**;
- time concepts: sample count, timestamp span, event duration and integration semantics are separate;
- PowerPoint phase-duration/intermediate-SOC cells are descriptive, not row-generation inputs;
- source-only report must not be labelled as the complete duty cycle.
