# Phase 1 Reverse-Engineering Report

## Status

**Phase 1 file inspection is complete.** The three supplied files are sufficient to define the first formal input/output specification and begin the import/channel-management engine. They are not sufficient to validate the stated 500–700-channel scale because the supplied source workbook contains 70 channels.

## Authoritative source package

| Role determined from file contents | File | SHA-256 |
|---|---|---|
| Supplied VSM/source-data example, already partially post-processed | `Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx` | `176a68563304953fbda8433dd27798851f96cb7e68cf2121952a605f76de35e9` |
| Desired Excel post-processing/report example | `Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx` | `1ac1fd0cc4b73a584410e78edbbd65d966077e15499da89ec004070c8130cd80` |
| Desired engineering PowerPoint example | `Sprayer Caiman SP_Hybrid-Electric_02.pptx` | `74cbc889f016b7e3075d6f0dd24468eadc6b9c1e98b198c317d9f2e5b878f6a3` |

## 1. File-role conclusion

The workbook with the long scenario filename is the closest thing to the supplied VSM source file. It has one sheet, no charts, 70 columns (`A:BR`), and 1,866 one-second samples. However, it is **not a pristine VSM export**: 25 of its 70 columns contain Excel formulas, while 45 columns are direct numeric source channels.

The `Electrification_03` workbook is the output/reference report. It contains the same 70-channel working area, a KPI block to the right, RMS values above the channel headers, bottom-of-column statistics, and 20 embedded Excel scatter charts.

The PowerPoint is a manually assembled engineering presentation. It contains no native PowerPoint charts and no embedded Excel workbook; the Excel tables and plots are pasted as static pictures.

## 2. Source workbook structure

| Item | Finding |
|---|---|
| Sheet | `Sprayer_Caiman_SP_9300Kg_Hybrid` |
| Used range | `A3:BR1870` |
| Header row | 3 |
| Unit row | 4 |
| Data rows | 5–1870 |
| Samples | 1,866 |
| Time channel | `Track_Time` in column A, seconds |
| Time range / step | 0–1865 s, exactly 1 s |
| Distance at end | 12 km |
| Channels | 70 total: 45 direct source channels + 25 formula-derived channels |
| Charts / tables / filters | None |
| Freeze panes | First column and top five rows |

The data region is rectangular and numeric: no missing cells or error-valued cells were found in the 1,866 × 70 sample area.

## 3. Channel organisation and naming

The source workbook uses VSM-style names such as `Chassis_Speed`, `ElectricSystem_Battery_Power`, and `Engine_AuxiliaryTorque_1`. The report keeps the same column order but replaces many source names with reader-friendly aliases such as `Speed`, `Battery Power`, and `Generator Torque_1`.

The report contains duplicate display names (`Time`, `Distance`, `Agrochemical Discharge`, and `Tyre Rolling Resistance Energy`). Therefore, the future program must use a unique internal channel identifier rather than treating the display name as a key. The recommended identity is:

`source_name + source_column_index`, with an optional configurable display alias.

The complete mapping is in `channel_mapping_inventory.csv`.

## 4. Math channels found in the supplied source workbook

The source example contains 25 formula-derived channels. Main formula families are:

- unit conversion: seconds → minutes, metres → kilometres;
- electrical/mechanical totals;
- power from speed × torque using divisor `9548.8`;
- per-sample energy using division by `3600`;
- cumulative energy;
- force-to-mass conversion using `-Fz/9.81`;
- squared helper channels for later RMS calculation.

Important detail: columns N and P are labelled as RMS channels but actually contain **squared instantaneous values** (`M²` and `O²`) with units `kW²`. The final RMS values are calculated separately in the top rows. A professional implementation should calculate RMS directly and should not require exporting square-helper columns unless explicitly requested.

The complete formula list is in `math_channels_inventory.csv`.

## 5. Desired Excel report structure

| Item | Finding |
|---|---|
| Sheet | `Hybrid_1C2G_30-60kph (2)` |
| Used range | `A1:DN17424` |
| Main channel area | `A:BR` (70 columns) |
| Blank separator | `BS` |
| KPI block | `BT:DM`, labels in row 3 and values/formulas in row 4 |
| Data rows | 5–17422 |
| Samples | 17,418 |
| Duration | 4.8381 h (17417 s) |
| Bottom statistics | Row 17423, plus battery-power MIN in row 17424 |
| RMS placement | Labels in merged cells on row 1; values in row 2 |
| Charts | 20 embedded scatter charts: 18 meaningful and 2 apparent scratch/orphan charts |
| Freeze panes | First column and top five rows |

The KPI block contains 46 summary columns. Several are intentionally blank because the required thermal/efficiency source channels are not present in this example.

## 6. RMS logic

The workbook calculates:

- Battery power RMS: `SQRT(SUM(N5:N17422)/17417)`
- Battery heat-flow RMS: `SQRT(SUM(P5:P17422)/17417)`

There are 17,418 samples but the denominator is 17417, the elapsed duration in seconds. This produces:

| Quantity | Workbook value | Sample-mean RMS | Relative difference |
|---|---:|---:|---:|
| Battery power | 55.132631 kW | 55.131048 kW | 0.002871% |
| Battery heat flow | 4.554083 kW | 4.553952 kW | 0.002871% |

The numerical difference is small for this one-second dataset, but the method is ambiguous. The processing engine should define RMS explicitly:

- ordinary RMS: `sqrt(mean(x²))` for uniformly sampled data;
- time-weighted RMS for irregular sampling, using the time channel and a documented integration convention.

## 7. MAX, MIN and last-value logic

Row 17423 is not a single universal operation. It mixes:

- `MAX` for speed, acceleration, torque, power, heat flow, and similar demand channels;
- final/last values for elapsed time, distance, SOC, cumulative energy and fuel;
- `SUM` for incremental energy columns;
- derived combinations such as total wheel power;
- capacity used as initial battery energy minus final battery energy.

Row 17424 contains only `MIN(Battery Power)`, later converted to a positive maximum discharge demand in the KPI block.

This confirms that statistics must be configured **per channel and per statistic type**, not inferred from units alone. The exact formulas are in `bottom_statistics_inventory.csv`.

## 8. Duty-cycle construction discovered in the report

The report is not simply the supplied 31-minute source dataset with formatting applied. It constructs a 290-minute, 114 km duty cycle from repeated/extended segments.

Five explicit opportunity-charging blocks were found:

- rows 1767–2666;
- rows 4464–5363;
- rows 9618–10517;
- rows 12081–12980;
- rows 14657–15556.

Each block is 900 seconds long. Battery energy and SOC are increased by `15/900` per second, corresponding to a 15 kWh/SOC-point increase over each 15-minute stop. The PowerPoint arrows identify these as charging while loading the agrochemical tank.

This duty-cycle assembly is scenario-specific and should **not** be hard-coded into the general post-processing engine. It belongs in a later, explicit scenario-construction module only if Sergio confirms it is required.

## 9. Formula and workbook quality findings

The reference workbook is valuable as a layout target, but it contains manual artefacts that should not be copied blindly:

- the source workbook's `Auxiliary Energy Consumption Accumulated` column does not accumulate correctly; the report fixes it with `current increment + previous accumulated value`;
- the first tyre-energy accumulated formula has an off-by-one-looking range that includes the following row;
- cells `DL4 = A43920` and `DM4 = C4` are stray/broken references in the KPI area;
- cells `T2`, `U2`, `DK1:DM2` appear to be scratch calculations;
- chart 14 and chart 15 are untitled/incomplete and positioned far down the worksheet;
- several labels contain typos or mismatches, including `Haigh`, `Accomulated`, and the report's AT channel labelled `FR` although the source channel is `RR`;
- no Excel tables, named ranges, filters, data validation or conditional formatting are used.

The new tool should preserve numerical intent while eliminating these manual artefacts.

## 10. Chart structure

All meaningful charts are XY scatter charts. They use either time in minutes or distance in kilometres as the x-axis. Several use a secondary y-axis.

The meaningful chart families include:

- speed versus distance/time;
- battery energy and SOC;
- battery charge/discharge power;
- generator power and fuel consumption;
- engine power and fuel consumption;
- energy released/recuperated;
- auxiliary and tyre energy consumption;
- agrochemical load/discharge;
- wheel and EDU power comparison.

The detailed chart source ranges, series and worksheet anchors are in `excel_chart_inventory.csv`.

## 11. Excel-to-PowerPoint relationship

The PowerPoint has 21 slides and is 16:9. Slides 7–12 are the key reporting examples:

- slide 7: a static image of the Excel KPI table;
- slides 8–9: the same KPI strip plus Speed vs Distance and Battery SOC plots;
- slides 10–12: the same KPI strip plus Generator Power and Fuel Consumption plots;
- red arrows and explanatory captions are native PowerPoint annotations.

Slides 8 and 9 are effectively duplicates; slides 10, 11 and 12 are also effectively duplicates. This reinforces that the PowerPoint should be treated as a style/content-selection reference rather than a clean reusable template.

Because the Excel visuals are pasted as pictures, future PowerPoint generation can use deterministic Python-rendered KPI tables and plot images. It does not need live Excel chart links.

## 12. Formal implementation implications

The core architecture should separate:

1. source import and schema detection;
2. unique channel registry and aliases;
3. deterministic math-channel evaluation;
4. RMS/statistics evaluation;
5. plot definitions;
6. Excel report generation;
7. optional PowerPoint composition;
8. user configuration/interface.

The next implementation milestone should be **data import + channel inventory only**. It should load the supplied source workbook, detect the header/unit/data rows, distinguish direct source channels from formula-derived columns, validate the time base, and export a clean channel catalogue. No GUI or PowerPoint generation should be started yet.

## Deliverables generated in this package

- `phase_1_reverse_engineering_report.md`
- `phase_1_formal_io_specification.md`
- `channel_mapping_inventory.csv`
- `math_channels_inventory.csv`
- `kpi_summary_inventory.csv`
- `bottom_statistics_inventory.csv`
- `excel_chart_inventory.csv`
- `powerpoint_slide_inventory.csv`
