# Known Limitations

Version 1.3.0 provides deterministic profile reporting for Electric, Hybrid and
Full Size Sprayer Diesel, subject to the following boundaries.

1. Current profiles follow the Diesel `_03` and Electric/Hybrid `_07` reporting layouts; original profiles remain explicit legacy options. Release checks use the updated Diesel `_07` CSV and Electric UoL `_04` CSV, including nonzero gearbox torque and Sergio-confirmed HV power. Hybrid passes available older-reference and controlled integration tests, including generator activity. The latest real Hybrid export has not been individually acceptance-tested; this is not a release gate or a known software defect.
2. Optional source channels, including `Track_Height`, are exported only when present in the uploaded data.
3. The current automatic header/unit detection is validated against normal CSV/XLSX layouts. Unusual multi-sheet or multi-level-header exports may require explicit import settings.
4. Math channels are configured expressions, not a free-form graphical formula editor.
5. Plots are configuration-driven rather than a fully interactive chart designer.
6. Profile Excel reports use native editable scatter charts. PowerPoint and UI previews use deterministic PNG plots; legacy configurable report paths retain their own layout behavior.
7. PowerPoint report content is deterministic and template-driven. Manual annotations present in reference presentations are not automatically inferred.
8. `SETUP_VSM_TOOL.bat` performs an online first-time bootstrap when a validated Python 3.11.x runtime or Python packages are missing. The automatic runtime target is CPython 3.11.9 because it is the supported Python 3.11 Windows binary installer for this release. Company networks that block downloads may require manual Python installation or an offline wheelhouse prepared by the maintainer.
9. The release is not yet a fully standalone/offline Windows executable or MSI installer.
10. AI-assisted KPI/plot recommendations are not included in the deterministic release. If introduced later, AI must remain advisory and must not replace numerical calculations.

11. Diesel has the original Test_03 CSV validation and updated `_07` release
    validation, plus controlled boundary tests. D1/D2/D3 are implementation stages,
    not three independent field datasets.
12. Diesel rejects invalid/non-finite source data, resetting mass/distance counters,
    negative RPM/flow, and non-increasing time. Optional absence is allowed;
    current profiles explicitly mark invalid optional sources unavailable. Sampled integration estimates
    the continuous simulation and is not its internal integrator.
13. Diesel fuel mass and integrated volume are independent source measurements.
    No density, efficiency, cycle BSFC, fuel-per-distance or fuel-per-area metric
    is assumed. Composed duty cycles are outside the current Diesel profile.
    Auxiliary consumption, tyre losses and wheel power follow the configured
    Diesel template definitions. Signed torque/power and
    net energy retain overrun. Ordinary drive-cycle duration is last minus first time.
14. Optional road height and temperature availability depend on the uploaded
    source. Oil/water traces happen to be identical in the old Diesel CSV, but
    both are independently plotted. Missing temperature measurements are named
    on the current thermal slide; available temperature traces, including zeros,
    are retained. Other optional-only Diesel slides may be omitted when their
    plots are absent/all-zero. Display rounding does not alter underlying metrics.
15. Office rendering is checked with installed Windows Excel and PowerPoint,
    not every third-party viewer. Office is not required for normal generation.
16. Current profiles include live Excel integration/cumulative/statistic formulas.
    All applicable configured KPI Values have live dependency chains, with
    formula helpers and documented parameters on hidden Formula Support.
    Legacy visible MATH columns may retain Python values; KPI helpers preserve
    their original conventions. Regenerate after source changes to refresh all
    retained values, chart caches and PowerPoint. Excel and PowerPoint remain
    separate files; editing Excel does not update a prior PowerPoint.
