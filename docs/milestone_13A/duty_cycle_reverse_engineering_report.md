# Milestone 13A – Sergio Reference Fidelity & Duty-Cycle Reverse Engineering

> **13A.1 clarification:** KPI semantics, time conventions, the 5–80% final-cycle range-extender override, reference hierarchy, and remaining conflicts are formalized in `../milestone_13A_1/`. The row mapping in this report remains valid, but the v2 specification supersedes the original Phase 13A spec.

## 1. Scope and conclusion

Phase 13A was limited to reverse engineering. No duty-cycle composer was implemented.

The Sergio reference workbook contains one data sheet, with the mission data in **rows 5:17422 (17,418 samples)** at **1 Hz**. Row 17423 is a mixed MAX/last/SUM summary row and row 17424 contains the battery-power MIN. The supplied VSM source contains **1,866 samples in rows 5:1870** and approximately 12 km of 30 kph field operation.

The 17,418-row mission is a structured scenario, not a simple repetition of the 1,866-row source. The numerically supported high-level sequence is:

**field ×3 -> road -> field ×3**, with **five 900 s agrochemical-loading/opportunity-charging stops** inserted between work sections. This matches the PowerPoint scenario `222 hectares of work – 40 km route – 222 hectares of work` and the repeated opportunity-charging slides, while the Excel itself actually contains **42.0087 km** of road travel.

## 2. Exact phase inventory

| Phase | Type | Excel rows | Samples | Source rows aligned | Distance delta km | SOC start % | SOC end % | Generator |
|---|---|---:|---:|---:|---:|---:|---:|---|
| P01 | field_work | 5:1766 | 1762 | 5:1766 | 11.9969 | 95.0000 | 61.8803 | off |
| P02 | loading_opportunity_charge | 1767:2666 | 900 | — | 0.0000 | 61.8970 | 76.8803 | ~80 kW charge |
| P03 | field_work | 2667:4463 | 1797 | 5:1801 | 11.9995 | 77.0000 | 43.7648 | off |
| P04 | loading_opportunity_charge | 4464:5363 | 900 | — | 0.0000 | 43.7815 | 58.7648 | ~80 kW charge |
| P05 | field_work | 5364:7081 | 1718 | 5:1722 | 11.9989 | 59.0000 | 43.0362 | on/controlled |
| P06 | road_travel | 7082:9617 | 2536 | — | 42.0087 | 42.0000 | 26.2109 | on/controlled |
| P07 | loading_opportunity_charge | 9618:10517 | 900 | — | 0.0000 | 26.2276 | 41.2109 | ~80 kW charge |
| P08 | field_work | 10518:12080 | 1563 | 5:1567 | 11.9989 | 41.0000 | 42.9043 | on/controlled |
| P09 | loading_opportunity_charge | 12081:12980 | 900 | — | 0.0000 | 42.9210 | 57.9043 | ~80 kW charge |
| P10 | field_work | 12981:14656 | 1676 | 5:1680 | 11.9978 | 58.0000 | 42.3984 | on/controlled |
| P11 | loading_opportunity_charge | 14657:15556 | 900 | — | 0.0000 | 42.4151 | 57.3984 | ~80 kW charge |
| P12 | field_work | 15557:17422 | 1866 | 5:1870 | 12.0000 | 57.4000 | 23.9383 | off |

The 12 phase sample counts sum exactly to **17,418**.

### Loading / opportunity charging placement

The five 900-row loading/charging blocks are exactly:

- rows **1767:2666**
- rows **4464:5363**
- rows **9618:10517**
- rows **12081:12980**
- rows **14657:15556**

At 1 Hz each block contains 900 per-sample charging increments. The worksheet timestamps span 899 s inside each block because the next phase starts at the +900 s timestamp; the configured/event duration is therefore 900 s.

## 3. Source-to-report mapping

### Field 1

Rows 5:1766 are a direct replay of source rows 5:1766 for the core vehicle profile. The key deliberate change is the battery state: **source 57.4% -> report 95.0%**, a constant **+37.6 percentage-point offset** through this phase.

### Field 2

Rows 2667:4463 replay source rows 5:1801. Distance is source distance plus **11,997.9 m** and Q/R are source values plus **19.6 points**. Cumulative released/recuperated energy and fuel are offset by the prior mission totals.

### Field 3

Rows 5364:7081 are time-aligned to source rows 5:1722. Rows 5364:6318 match the core supplied source profile; at row **6319** SOC crosses from **40.0058% to 39.9828%** and the range extender begins. From that point the reference profile diverges from the supplied source and generator power rises toward ~80 kW.

### Road

Rows 7082:9617 are a separate reference profile. It starts at zero speed, explicitly resets SOC to **42.0%**, sets agrochemical load to zero, accelerates to ~60 kph, and ends at **78.0044 km**. The actual road distance is **42.0087 km**. No supplied source workbook contains this >60 kph profile.

### Field 4

Rows 10518:12080 are time-aligned to source rows 5:1567. Only the first 44 rows are exact core matches. The generator turns on at row **10562**, where SOC crosses **40.0072% -> 39.9834%**.

### Field 5

Rows 12981:14656 are time-aligned to source rows 5:1680. The first 909 rows match the core source profile. The generator turns on at row **13890**, where SOC crosses **39.9959% -> 39.9786%**. The AL agrochemical formulas near the end of this phase are visibly misaligned and should not be treated as authoritative physics.

### Field 6

Rows 15557:17422 are a full 1,866-row source replay aligned to source rows 5:1870, with global distance/energy/fuel offsets. It starts at the source SOC **57.4%** and therefore ends exactly at the source final SOC **23.9383%**.

A per-row provenance mapping is provided in `source_to_report_row_mapping.csv`.

## 4. Battery / charging logic

The reference assumes a **100 kWh battery**: KPI cell CQ4 is `Q5 / 0.95 = 100`. Q (`Battery Energy`, kWh) and R (`Battery SOC`, %) are therefore numerically identical.

During each loading block:

- Battery power M = **+60 kW**.
- Q and R increment by **15/900 = 0.0166667** per row.
- 60 kW × 900 s / 3600 = **15 kWh**, which is **+15 SOC points** for a 100 kWh pack.
- The next work phase does not always begin at the exact preceding charging result; it is manually reset/rounded to **77%, 59%, 41%, 58%, 57.4%**.

Initial SOC is not a different source segment: the first report phase is the same source profile with a +37.6 point state offset.

## 5. Generator / fuel logic

Field phases 1, 2 and 6 have generator power zero. In fields 3, 4 and 5, generator operation is triggered when SOC falls below approximately **40%**. Once triggered, the reference keeps the range extender running through the remainder of the phase.

During opportunity charging, the nominal loading-state values are:

- engine speed ~**1700 rpm**
- generator-1 torque ~**449.379 Nm**
- generator power ~**80.0042 kW**
- battery charging power **60 kW**
- fuel flow ~**17.4 l/h**
- fuel mass increment ~**0.00406 kg/s** using 0.84 kg/l
- auxiliary power mostly **10 kW**

Generator 2 (BP/BQ) is zero throughout this reference workbook.

## 6. Agrochemical loading logic

AL is computed from AK using `AL = -AK / 9.81`. During loading, the tank increases by approximately **4.4429266 kg/s**, reaching about **3998.63 kg** after the 900 charging rows and approximately **3998.65 kg** at the next field initialization. Road travel uses zero agrochemical load.

The reference has some formula alignment errors in AL, particularly around field 5. The composer should model tank mass directly rather than copy those formulas.

## 7. Cumulative and derived channels

The detailed inventory is in `cumulative_channel_logic.csv`. The important deterministic rules are:

- B = A/60.
- G = F/1000.
- L = J+K.
- N = M²; P = O².
- X = V×W/9548.8; AB = Z×AA/9548.8; AC = X+AB.
- AI = AG×AH/9548.8.
- AJ = AI×dt/3600 and is summed for total engine energy.
- AU = AQ+AR+AS+AT.
- AV = AU×dt/3600; AW = 1000×AV; AX should be a true cumulative sum.
- BK = BI+BJ; BL = BK×dt/3600; BM is a true cumulative sum.
- BO = AG×BN/9548.8; BQ = AG×BP/9548.8; BR = BO+BQ.

## 8. Numerical validation against Sergio KPIs

| KPI | Reference / reconstructed value | Status |
|---|---:|---|
| Data samples | 17,418 | exact |
| Final track time | 17,417 s = 290.283333 min | exact |
| Final distance | 114.0011 km | exact |
| Final SOC | 23.9383 % | exact |
| Final fuel consumption | 39.838060 kg | exact |
| Max speed | 62.6233 kph | exact |
| Max generator power | 80.0266906 kW | exact |
| Energy recuperated | 1.760681 kWh | exact |
| Energy released | 161.0646 kWh | exact |
| Used EDU energy (released - recuperated) | 159.303919 kWh | exact |
| Auxiliary energy | 109.3526082 kWh | exact |
| Tyre RR energy | 77.67730125 kWh | exact |
| Engine energy SUM(AJ) | 219.3011311 kWh | exact |

## 9. Reference quirks that should not become composer rules

1. RMS formulas divide by 17,417 although 17,418 samples exist. Reference battery-power RMS = **55.132631 kW**; using all 17,418 samples gives **55.131048 kW**.
2. Small distance discontinuities occur at some phase boundaries.
3. SOC is manually rounded/reset after loading phases.
4. Some loading boundary cells momentarily inherit incorrect auxiliary/fuel/generator states.
5. AX cumulative formulas use inconsistent absolute starting rows in places.
6. AL formulas are misaligned in parts of field 5.
7. Generator total BR is forced to zero in a few first loading rows even when BO already contains ~80 kW.

These should be documented differences if the new deterministic composer chooses physically coherent logic rather than cell-for-cell reproduction of obvious spreadsheet artifacts.

## 10. Formal duty-cycle specification for Phase 13B

A configuration-oriented specification is provided in `proposed_duty_cycle_spec.yaml`. It intentionally contains **no Python implementation**. It defines the 12 phases, source-row alignment, 1 Hz timing, SOC threshold, 900 s loading/charging action, 100 kWh battery, and the reference KPI targets.

## 11. Remaining ambiguity before exact 70-channel equivalence

The main blocker is not the mission phase structure anymore. It is the origin of the **road profile** and the **generator-on field variants**. They cannot be reproduced cell-for-cell from the supplied no-generator 30 kph VSM workbook alone. If Sergio has separate 60 kph and generator-active VSM runs, obtaining them would remove most of the remaining uncertainty.

Until then, the phase structure and all cumulative/control logic above are sufficiently constrained to design the Phase 13B composer interface, while exact values for reference-only profiles should remain explicit external profile inputs rather than guessed formulas.
