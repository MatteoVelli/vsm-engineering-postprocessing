# Milestone 13B.2 - External Phase Profile Provider

## Objective

Milestone 13B.2 removes the architectural blocker identified in 13B.1 without inventing missing vehicle dynamics.
The duty-cycle composer can now accept an explicit external numerical profile provider for phases whose dynamics are not contained in the supplied 30 kph field workbook.

For the Sergio fidelity scenario, the unresolved phases are:

- **P05** - third field cycle after range-extender activation;
- **P06** - independent road-travel profile;
- **P08** - fourth field cycle with range extender active;
- **P10** - fifth field cycle with range extender active.

The current fidelity provider reads only those four phase ranges from Sergio's authoritative Excel workbook. This is deliberately classified as a **reference-fidelity bridge**, not a predictive model and not an independently reconstructed VSM simulation.

## New architecture

```text
DutyCycleScenario
       |
       +--> native source phase ----------+
       +--> deterministic loading phase --+--> DutyCycle Composer --> complete mission
       +--> external PhaseProfileProvider -+
```

New public components:

- `WorkbookProfileProviderConfig`
- `WorkbookRowProfileProvider`
- `ProfileProviderValidation`
- `PhaseProfileProvenance`
- `compose_duty_cycle(...)`
- `export_duty_cycle_composition(...)`
- `load_profile_provider_config(...)`

The provider contract is separate from `composer.py`. A future road/range-extender VSM workbook can therefore replace the Sergio reference provider without changing composition logic.

## Provider configuration

`config/duty_cycle_profiles_sergio_reference.yaml` freezes:

- provider identity;
- provider type;
- exact supported phase IDs;
- exact source row range for each phase;
- import/header bounds;
- channel-alignment policy;
- value policy;
- expected workbook filename;
- expected SHA-256.

The Sergio report uses renamed column headers compared with the source workbook. Therefore the reference-fidelity provider uses explicit **column-position alignment** for the fixed 70-column working area. Positional alignment is accepted only when the workbook SHA-256 is configured and matches exactly. This prevents a different workbook with a superficially similar layout from being silently accepted.

## Provenance policy

Every externally supplied row is marked in the composition plan with:

- `generation_mode = external_profile:<provider_id>`;
- `profile_provider_id`;
- `profile_provider_source`.

Every externally materialised phase also carries a `PhaseProfileProvenance` record containing:

- provider ID/type;
- value policy;
- source filename;
- source SHA-256;
- source start/end row;
- sample count;
- channel count.

The four Sergio provider phases are copied as **absolute reference values**. The provider verifies that their time channel exactly matches the global duty-cycle plan before returning the data.

## Full numerical composition

With the source field workbook and the Sergio reference provider available, the composer now materialises all 12 phases and all 17,418 samples:

| Quantity | Full composition |
|---|---:|
| Samples | 17,418 |
| Final Track_Time | 17,417 s |
| Final Time | 290.283333 min |
| Final Distance | 114.0011 km |
| Final Battery SOC | 23.9383 % |
| Max Speed | 62.6233 kph |
| Max Generator Power | 80.0266906 kW |
| Canonical final Fuel Consumption | 39.84212 kg |

The final fuel value is 0.00406 kg above Sergio's Excel value because the canonical loading implementation integrates every configured one-second fuel increment. The reference workbook omits one increment at a documented loading boundary. This intentional difference remains in the deterministic composer.

## Reference-equivalence result

The external phases **P05/P06/P08/P10 are exact row/column replays** of their configured provider ranges.
Across the complete mission, the core electrical trajectory is numerically equivalent to the Sergio workbook for:

- Track_Time;
- Time [min];
- Battery Power;
- Battery Heatflow;
- Battery Energy;
- Battery SOC.

Remaining differences are restricted to already documented canonical-vs-reference choices and legacy boundary/formula behaviour. The complete per-channel inventory is in:

`profile_provider_reference_difference_inventory.csv`

The phase/provider mapping is in:

`profile_provider_phase_inventory.csv`

## Engineering boundary

13B.2 does **not** claim that the missing generator-active/road physics have been inferred from the supplied source workbook. The exact Sergio reference workbook is currently required for the four missing profiles in reference-fidelity mode.

For a reusable client scenario that does not depend on Sergio's output workbook, the next profile-source upgrade is to provide independent VSM road and generator-active field result files through the same provider interface.

## Completion criteria

13B.2 is complete when:

1. provider configuration is strictly validated;
2. source/provider channel layout and 1 Hz timing are validated;
3. provider workbook SHA-256 is verified for positional alignment;
4. all four missing phases are resolved only when an explicit provider is supplied;
5. full composition contains exactly 17,418 samples and 12 phases;
6. externally provided phases are byte-for-numeric-value equivalent to their configured source rows;
7. row-level and phase-level provider provenance is retained;
8. the existing prefix mode still stops before P05 when no provider is supplied;
9. the full regression suite remains green.
