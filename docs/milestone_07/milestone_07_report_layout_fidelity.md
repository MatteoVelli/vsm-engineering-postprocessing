# Milestone 07 - Report Layout & Template Fidelity

## Objective

Bring the generated Excel report substantially closer to the workbook supplied by Sergio without changing any deterministic engineering calculation.

## Why this milestone exists

Milestone 06 proved that the full processing chain can generate a valid workbook. Its layout was intentionally generic. The client example and Sergio's written requirements show a more specific reporting convention:

- RMS labels and values occupy rows 1 and 2 above selected channels;
- selected channel names and units occupy rows 3 and 4;
- samples start on row 5;
- selected MAX/MIN/last/SUM results appear directly below their relevant channel;
- a second selection of KPI values is displayed in rows 3 and 4 to the right of the channels;
- plots are aligned underneath that secondary selection;
- math channels must be visually distinguishable from raw VSM channels.

## Implemented layout profile

A new `sergio_reference` Excel layout profile has been added. It is configured in `config/excel_report_example.yaml`.

Visible report layout:

| Area | Placement |
|---|---|
| RMS labels | row 1, above the RMS source channel |
| RMS values | row 2 |
| Channel names | row 3 |
| Units | row 4 |
| Data samples | row 5 onward |
| Bottom channel statistics | first row immediately after final sample; second row when a channel has a second statistic such as battery MIN |
| KPI/secondary selection | row 3 labels and row 4 values, after one blank separator column |
| Plots | row 6 onward beneath the KPI strip |
| Freeze panes | `B5` |

## Styling decisions derived from the reference workbook

- Calibri 11 is retained.
- Raw channel headers use a white background.
- Math channel headers use a pale-yellow background, matching the visual convention found on formula-derived columns in the client workbook.
- RMS cells use a blue fill with bold white text.
- Channel and KPI headers use centered wrapped text with borders.
- Header row 3 is intentionally tall to accommodate engineering channel names.
- The report sheet no longer uses rows 1 and 2 for a generic document title; those rows are reserved for RMS as requested by Sergio.
- Document identity and provenance remain available on the `Metadata` sheet.

## Numerical architecture

No calculations were moved into Excel. All RMS, statistics and math channels still come from the validated Python processing engines. Excel remains a reporting layer only.

## New configuration capability

`statistics.bottom_summary` defines the exact statistic IDs to display underneath channel columns. This is intentionally separate from the calculation configuration and from the KPI selection.

The previous `bottom_operations` configuration remains available for the Milestone 06 engineering layout, preserving backwards compatibility.

## Plot layout

The plotting engine still produces deterministic PNG assets. Under the `sergio_reference` profile those images are embedded beneath the KPI strip rather than below the full 1,866-row dataset. This reflects the supplied client workbook and makes plots immediately visible when the report opens.

Native editable Excel charts are not introduced in this milestone. They remain a possible later enhancement if Sergio specifically requires editable chart objects.

## Scope intentionally not added

This milestone does not add:

- GUI/channel-selection interface;
- PowerPoint generation;
- scenario/duty-cycle construction;
- AI recommendations;
- live Excel formulas for engineering calculations.

## Acceptance criteria

The milestone is complete when the generated workbook demonstrates:

1. rows 1-2 reserved for configured RMS values;
2. rows 3-4 used for channel and KPI headers/values;
3. data beginning at row 5;
4. raw/math visual distinction;
5. configured channel-specific summary values below the data;
6. battery MAX and MIN occupying consecutive bottom rows;
7. KPI strip separated from the channel data by a blank column;
8. six configured plots positioned beneath the KPI strip;
9. unchanged deterministic numerical results;
10. metadata/provenance retained.


## 0.7.1 hotfix

The Sergio-reference acceptance test exposed a configuration mismatch: `report_total_edu_power_max` targeted `calc_total_edu_elect_power`, but that channel was not exported in the report. The fix replaces the non-essential exported `edu_mech_power_rl__col_024` channel with `calc_total_edu_elect_power`. This preserves the 19-column report geometry while making every bottom statistic visible under its own exported channel. No numerical engine logic changed.
