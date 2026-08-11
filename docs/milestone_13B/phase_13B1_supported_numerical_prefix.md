# Milestone 13B.1 - Deterministic Numerical Prefix

## Objective

Materialise the longest Sergio duty-cycle prefix that can be justified from the supplied VSM source workbook and the formally resolved scenario actions, while preserving a hard gate before any phase whose physical profile is not available.

Milestone 13B.1 does **not** attempt to force a 17,418-row numerical result by copying or guessing unresolved profiles. The existing validated source-processing pipeline remains unchanged; the composer is optional.

## Implemented numerical scope

The composer now materialises the first four reference phases:

| Phase | Type | Samples | Materialisation |
|---|---|---:|---|
| P01 | field work 1 | 1,762 | source replay with scenario SOC/cumulative offsets |
| P02 | loading + opportunity charge 1 | 900 | deterministic synthetic scenario action |
| P03 | field work 2 | 1,797 | source replay with scenario SOC/cumulative offsets |
| P04 | loading + opportunity charge 2 | 900 | deterministic synthetic scenario action |

Total supported numerical prefix: **5,359 samples**.

The result stops before P05 and reports `stopped_before_phase_id = P05`.

## Source-backed field logic

For source-backed field phases the implementation:

- takes the configured source-row slice;
- replaces local source time with continuous mission `Track_Time`;
- keeps mission distance continuous using a deterministic offset;
- applies the configured phase-start battery SOC/energy initialization;
- offsets cumulative recuperated energy, released energy and fuel from the previous materialised phase;
- applies configured source-restart cleanup rules;
- recalculates deterministic math channels from their dependencies rather than copying spreadsheet formulas.

The first EDU-speed samples are explicitly reset according to the scenario cleanup rule because the source first row contains the already documented 22,750 rpm restart artefact.

## Loading/opportunity-charge logic

Each 900-s synthetic loading phase uses the formally resolved scenario values:

- battery charging power: +60 kW;
- battery heatflow: 1.08 kW;
- battery capacity used for mission SOC: 100.0 kWh;
- generator/engine speed: 1,700 rpm;
- generator-1 torque: 449.379 Nm;
- engine torque: 459.378 Nm;
- specific fuel consumption: 182.7 g/kWh;
- fuel volumetric flow: 17.4 L/h;
- fuel density: 0.84 kg/L;
- tank-loading rate: 4.4429266055 kg/s;
- low-voltage auxiliary demand: 10 kW;
- high-voltage auxiliary demand: 0 kW;
- generator 2: zero.

At 1 Hz, +60 kW for 900 samples adds 15 kWh and therefore 15 SOC percentage points on the 100 kWh mission battery scale.

## Canonical differences from Sergio spreadsheet artefacts

The composer intentionally follows the deterministic specification instead of copying known cell-boundary artefacts. In the supported prefix this causes small, documented differences from the reference workbook:

1. **Distance continuity** - the composer holds distance exactly during loading and does not reproduce the +1 m spreadsheet boundary jump before P03.
2. **P04 first generator sample** - the configured opportunity-charge action is active from the first P04 sample; Sergio's first P04 reference cell temporarily reports zero generator power.
3. **P04 fuel integration** - the composer integrates all 900 one-second loading samples. Sergio's P04 first row omits one fuel increment, so the canonical result is +0.00406 kg higher at the end of P04.

These differences are regression-tested and are intentional.

## Why P05/P08/P10 remain unresolved

A dedicated comparison was made between the source-aligned field rows and Sergio's reference rows before and after range-extender activation.

The pre-trigger source-aligned portions match the supplied source profile (apart from previously documented restart/cumulative cleanup). At the generator trigger, the reference does **not** merely add engine/generator values on top of the same vehicle trace. The post-trigger vehicle dynamics themselves change.

| Phase | Exact source-aligned portion | Generator-active portion |
|---|---|---:|
| P05 | report rows 5364:6318 (955 samples) | 763 samples |
| P08 | report rows 10518:10561 (44 samples) | 1,519 samples |
| P10 | report rows 12981:13889 (909 samples) | 767 samples |

After the trigger, differences appear throughout independent raw/profile channels including longitudinal acceleration, vehicle speed, accelerator demand, EDU electrical power, battery heatflow, EDU speed/torque, driveshaft torque and tyre rolling-resistance force. For example, maximum source-vs-reference speed differences after the trigger are approximately 32.0 kph (P05), 32.2 kph (P08) and 31.8 kph (P10).

Therefore these phases are genuine range-extender VSM/reference variants, not a source replay plus a simple generator overlay. Synthesising them from the supplied final-cycle field workbook would be unjustified.

## P06 road profile

P06 remains a separate >60 kph road profile with no corresponding source dataset in the supplied 30 kph field workbook. It remains explicitly gated.

## Acceptance criteria achieved

- 17,418-row structural composition plan still reproduces the Phase 13A row/phase/source-alignment mapping.
- Supplied VSM workbook compatibility validation passes.
- P01-P04 materialise deterministically to 5,359 rows.
- Source and synthetic phase math channels are recalculated deterministically.
- Cumulative channels continue across the supported prefix under canonical rules.
- Composer stops before P05 rather than returning guessed/NaN profile data.
- Reference comparisons assert both resolved numerical agreement and documented canonical differences.

## Next increment

Milestone 13B.2 should introduce an explicit **profile-provider contract** for the unresolved physical profiles. Exact Sergio reconstruction can then proceed only when P05/P06/P08/P10 are supplied by a traceable source (for example separate VSM runs) or deliberately supplied as a clearly labelled legacy/reference compatibility profile.

The composer must not silently derive those profiles from unrelated source rows.
