# Known Limitations

Version 1.2.10 is a deterministic profile-reporting release for the validated RoboSprayer workflow, subject to the following boundaries.

1. Reporting profile fidelity is validated against the supplied Electric_03 and Hybrid_04 templates and representative RoboSprayer source CSVs.
2. Optional source channels, including `Track_Height`, are exported only when present in the uploaded data.
3. The current automatic header/unit detection is validated against normal CSV/XLSX layouts. Unusual multi-sheet or multi-level-header exports may require explicit import settings.
4. Math channels are configured expressions, not a free-form graphical formula editor.
5. Plots are configuration-driven rather than a fully interactive chart designer.
6. Excel charts are embedded deterministic PNG assets rather than native editable Excel chart objects.
7. PowerPoint report content is deterministic and template-driven. Manual annotations present in reference presentations are not automatically inferred.
8. `START_VSM_TOOL.bat` requires `uv` to be installed and available in PATH. The first setup may require internet access to obtain the locked Python environment/packages if they are not already cached.
9. The release is not yet a fully standalone/offline Windows executable or MSI installer.
10. AI-assisted KPI/plot recommendations are not included in the deterministic release. If introduced later, AI must remain advisory and must not replace numerical calculations.
