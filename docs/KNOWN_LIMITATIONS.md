# Known Limitations

Version 1.3.0 is a deterministic profile-reporting release for validated Electric/Hybrid workflows, subject to the following boundaries.

1. Reporting profile fidelity is validated against the supplied Electric_05 and Hybrid_06 templates and representative Electric/Hybrid source CSVs.
2. Optional source channels, including `Track_Height`, are exported only when present in the uploaded data.
3. The current automatic header/unit detection is validated against normal CSV/XLSX layouts. Unusual multi-sheet or multi-level-header exports may require explicit import settings.
4. Math channels are configured expressions, not a free-form graphical formula editor.
5. Plots are configuration-driven rather than a fully interactive chart designer.
6. Excel charts are embedded deterministic PNG assets rather than native editable Excel chart objects.
7. PowerPoint report content is deterministic and template-driven. Manual annotations present in reference presentations are not automatically inferred.
8. `SETUP_VSM_TOOL.bat` performs an online first-time bootstrap when a validated Python 3.11.x runtime or Python packages are missing. The automatic runtime target is CPython 3.11.9 because it is the supported Python 3.11 Windows binary installer for this release. Company networks that block downloads may require manual Python installation or an offline wheelhouse prepared by the maintainer.
9. The release is not yet a fully standalone/offline Windows executable or MSI installer.
10. AI-assisted KPI/plot recommendations are not included in the deterministic release. If introduced later, AI must remain advisory and must not replace numerical calculations.
