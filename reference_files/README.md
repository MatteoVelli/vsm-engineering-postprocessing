# Reference files

This directory contains both tracked runtime assets and optional private
regression data.

## Tracked runtime assets

The client requires the following unchanged source assets. They must be committed
at these exact, case-sensitive paths; `.gitignore` explicitly allows them:

```text
astauto-light-text_web.jpg
RoboSprayer_Electric_Report_Astauto_Colours.pptx
Caiman_SP_Hybrid_Report_Astauto_Colours.pptx
RoboSprayer_Electric_Report_Astauto_v7.pptx
Robo_Sprayer_Electrification_Tamplate_Electric_03.xlsx
Robo_Sprayer_Electrification_Tamplate_Hybrid_04.xlsx
Robo_Sprayer_Electrification_Tamplate_Electric_05.xlsx
Robo_Sprayer_Electrification_Tamplate_Hybrid_06.xlsx
```

The JPG supplies Astauto branding; the presentations and workbooks supply runtime
layout/reference information. They are required source inputs, not generated report
outputs. Doctor and the release builder retain mandatory checks for these files.
Do not replace them with generated reports or duplicate/rename the logo.

The existing `RoboSprayer_Electric_Report_FINAL.pptx` and
`RoboSprayer_Hybrid_Engineering_Report.pptx` also remain tracked for existing uses.

## Optional private local regression data

Original client source datasets are intentionally excluded from Git and release
ZIPs because they contain client data. When they are installed locally, the
full reference-backed acceptance tests run. When they are absent, those tests
skip explicitly.

Current private regression inputs include:

```text
RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv
Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx
Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx
```

Milestone 13B.2 uses `Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx` as an
external reference-fidelity phase provider for P05, P06, P08 and P10. The
provider configuration locks the expected filename and SHA-256; a
modified/different workbook is intentionally rejected.

Electric_05 and Hybrid_06 now supply the active channel ordering. The previous
Electric_03 and Hybrid_04 workbooks remain unchanged for reference comparison.
