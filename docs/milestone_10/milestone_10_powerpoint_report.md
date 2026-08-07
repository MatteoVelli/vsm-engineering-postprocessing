# Milestone 10 - Deterministic PowerPoint Report Generator

## Objective

Add an optional PowerPoint reporting layer after the validated Excel report without changing any deterministic engineering calculation.

## Reference findings used

Sergio's supplied 21-slide deck is 16:9 widescreen. The duty-cycle slides (especially slides 8-12) repeatedly use a useful presentation pattern:

- large centered engineering title;
- compact KPI strip near the top;
- two plot images side-by-side;
- optional annotation/footer text.

The reference deck uses static Excel screenshots rather than native PowerPoint charts. Milestone 10 therefore embeds the already validated PNG plot assets, preserving deterministic visual output.

Manual arrows/callouts from the client deck are intentionally not recreated because they are scenario-specific annotations rather than deterministic processing results.

## Inputs

The PowerPoint layer consumes:

- validated statistics from `statistics_excel_report.yaml`;
- validated plots from `plotting_example.yaml`;
- slide composition from `powerpoint_report_example.yaml`;
- the same source hash and sample count used by the rest of the pipeline.

## Output

Default standalone output:

```text
outputs/powerpoint_report/
├── vsm_engineering_report.pptx
├── powerpoint_report_manifest.json
├── powerpoint_report_summary.txt
└── plot_assets/
```

End-to-end output:

```text
outputs/end_to_end/07_powerpoint_report/vsm_engineering_report.pptx
```

## Default slide plan

1. Engineering summary - six KPI cards.
2. Duty cycle - speed/time and SOC/distance plots plus KPI strip.
3. Electrical energy system - battery/generator and auxiliary-energy plots plus KPI strip.
4. Powertrain demand - wheel/EDU and engine/fuel plots plus KPI strip.

## Engineering safeguards

- no AI selection or interpretation;
- only configured statistic IDs and plot IDs can be placed;
- missing statistics or plot assets fail explicitly;
- no recalculation inside PowerPoint;
- 16:9 deterministic slide dimensions;
- images preserve aspect ratio;
- manifest records source/config hashes, slides, plots and statistics;
- PowerPoint generation is optional in both the pipeline and UI.

## UI integration

The UI now contains a `PowerPoint` tab with a `Generate PowerPoint report` checkbox. When enabled, the selected statistics and plots are filtered into the slide template. The UI exposes separate download/open actions for Excel and PowerPoint.

## Validation

Milestone 10 was validated against Sergio's source workbook:

- PowerPoint standalone generation: PASS;
- 4 slides generated;
- 14 unique statistics used;
- 6 plot assets used;
- slide rendering: PASS;
- slide overflow test: PASS;
- end-to-end pipeline: 7/7 stages PASS;
- UI end-to-end acceptance: PASS;
- complete project regression inventory: 63 test cases, validated in groups to avoid long single-process runtime limits.
