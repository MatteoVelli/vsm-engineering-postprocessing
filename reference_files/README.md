# Reference files

Place the original client files in this directory. They are intentionally excluded from Git and release ZIPs because they contain client data.

Required for the complete reference-backed acceptance suite:

```text
Sprayer_Caiman_SP_9300Kg_Hybrid_Gen80kW_30kph_74Ht_4000KgAQ_57-4pcSOC_5-80_1C2G_02.xlsx
Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx
```

Milestone 13B.2 also uses `Sprayer_Caiman_SP_9300Kg_Electrification_03.xlsx` as an **external reference-fidelity phase provider** for P05, P06, P08 and P10. The provider configuration locks the expected filename and SHA-256; a modified/different workbook is intentionally rejected.

The PowerPoint reference may also be stored here for later reporting-fidelity milestones.
