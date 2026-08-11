# Known Limitations

Version 1.2.5 is a production-ready deterministic processing release for the validated VSM workflow, subject to the following boundaries.

1. The reference source workbook contains 70 channels. The architecture is intended for the 500-700 channel requirement, but full-scale acceptance should be repeated when a representative 500-700 channel VSM export is supplied.
2. The current automatic header/unit detection is validated against the supplied Sergio workbooks and normal CSV/XLSX layouts. Unusual multi-sheet or multi-level-header exports may require explicit import settings.
3. Math channels are configured expressions, not a free-form graphical formula editor.
4. Plots are configuration-driven rather than a fully interactive chart designer.
5. Excel charts are embedded deterministic PNG assets rather than native editable Excel chart objects.
6. PowerPoint report content is deterministic and configuration-driven. Manual annotations present in the supplied reference presentation are not automatically inferred.
7. `START_VSM_TOOL.bat` requires `uv` to be installed and available in PATH. The first setup may require internet access to obtain the locked Python environment/packages if they are not already cached.
8. The release is not yet a fully standalone/offline Windows executable or MSI installer.
9. AI-assisted KPI/plot recommendations are not included in the deterministic release. If introduced later, AI must remain advisory and must not replace numerical calculations.
10. The config-driven end-to-end pipeline supports the complete 17,418-row mission, but the Streamlit UI does not yet expose a dedicated full-duty-cycle scenario/provider selector; its normal uploaded-source workflow remains source-file based.
11. Full 17,418-row composition is available in Sergio reference-fidelity mode only when an explicit external provider supplies P05, P06, P08 and P10. The bundled repository does not contain that client workbook.
12. Independent VSM road and generator-active field profiles have not been supplied. Until they are, the Sergio reference provider is a traceable fidelity bridge rather than an independently predictive duty-cycle reconstruction.
13. The canonical composer intentionally differs from a few Sergio spreadsheet boundary artefacts, including one 0.00406 kg fuel increment and early +1 m distance handling. These differences are documented and tested.
14. Sergio's PowerPoint contains several rounded/stale values that conflict with the Excel; the Excel is the numerical authority for full-mission equivalence.
15. Legacy KPI formulas such as `Fuel Energy Consumed = 80 * 12` are compatibility-only until their engineering basis is confirmed.
