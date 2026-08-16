# Reference files

This directory contains both tracked runtime assets and optional private
regression data.

## Tracked runtime assets

These files are required by the profile-driven PowerPoint generator and should
be present in a clean checkout:

```text
RoboSprayer_Electric_Report_FINAL.pptx
RoboSprayer_Hybrid_Engineering_Report.pptx
```

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
