# Phase 13A.1 – Formal time and integration convention

## Decision

Sergio's workbook mixes four quantities that must no longer be called simply `duration`:

1. **sample_count** – number of output rows in a phase or mission;
2. **timestamp_span_s** – `last_timestamp - first_timestamp`;
3. **configured_event_duration_s** – scenario duration used to create a synthetic event such as a 900 s loading/charging stop;
4. **integration_semantics** – the rule used to convert sampled power into energy.

For the full Sergio Excel mission:

- `sample_count = 17,418`;
- first timestamp = `0 s`;
- last timestamp = `17,417 s`;
- `timestamp_span_s = 17,417 s`;
- Excel displays time from the final timestamp: `17,417 / 60 = 290.283333 min`;
- Sergio's PowerPoint phase-duration cells sum to `17,416 s`, so they are **not** an authoritative row-generation source.

## Composer rule

The future composer must generate the mission row count from the phase/profile definitions, not from PowerPoint duration cells.

At 1 Hz:

- a configured **900 s loading event produces 900 rows**;
- those rows have an internal timestamp span of 899 s;
- the next phase begins 900 s after the loading phase's first row.

Moving/reference phases are defined by their output sample counts/profile lengths. PPT duration values are descriptive only.

## Integration rule

Two separate modes must be explicit:

- `reference_per_sample_rectangle`: every output row contributes `power_kw * 1 s / 3600`, matching Sergio's per-row Excel energy construction. With 17,418 rows this produces 17,418 one-second contributions even though the timestamp span is 17,417 s.
- `timestamp_interval`: energy is integrated only over actual timestamp intervals. This is the physically cleaner interpretation but is not guaranteed to reproduce Sergio's cached totals.

Milestone 13B must never silently switch between these modes. Reference-equivalence tests should use `reference_per_sample_rectangle` where the Sergio workbook does; engineering calculations outside compatibility mode should declare their interval convention explicitly.

## Phase inventory

| Phase | Rows | Samples | Start s | End s | Timestamp span s | PPT duration s |
|---|---:|---:|---:|---:|---:|---:|
| P01 | 5:1766 | 1762 | 0 | 1761 | 1761 | 1761 |
| P02 | 1767:2666 | 900 | 1762 | 2661 | 899 | 900 |
| P03 | 2667:4463 | 1797 | 2662 | 4458 | 1796 | 1798 |
| P04 | 4464:5363 | 900 | 4459 | 5358 | 899 | 900 |
| P05 | 5364:7081 | 1718 | 5359 | 7076 | 1717 | 1717 |
| P06 | 7082:9617 | 2536 | 7077 | 9612 | 2535 | 2536 |
| P07 | 9618:10517 | 900 | 9613 | 10512 | 899 | 900 |
| P08 | 10518:12080 | 1563 | 10513 | 12075 | 1562 | 1563 |
| P09 | 12081:12980 | 900 | 12076 | 12975 | 899 | 900 |
| P10 | 12981:14656 | 1676 | 12976 | 14651 | 1675 | 1676 |
| P11 | 14657:15556 | 900 | 14652 | 15551 | 899 | 900 |
| P12 | 15557:17422 | 1866 | 15552 | 17417 | 1865 | 1865 |

The mismatch is therefore documented rather than normalized away. `phase_time_inventory.csv` contains the machine-readable version of this table.
