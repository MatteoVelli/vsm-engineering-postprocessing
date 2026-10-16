# Known Limitations

Version 1.3.0 provides deterministic profile reporting for Electric, Hybrid and
Full Size Sprayer Diesel, subject to the following boundaries.

1. Current profiles target Diesel `_03` and Electric/Hybrid `_07`; original `_05`/`_06` profiles remain explicit legacy options. Diesel and Electric have real-source validation, including nonzero Electric gearbox torque and Sergio-confirmed HV power. The updated Hybrid configuration is tested with controlled fixtures; complete acceptance still requires its updated raw CSV. See [current validation](HV_AUXILIARY_CONFIRMATION_2026_10_16.md).
2. Optional source channels, including `Track_Height`, are exported only when present in the uploaded data.
3. The current automatic header/unit detection is validated against normal CSV/XLSX layouts. Unusual multi-sheet or multi-level-header exports may require explicit import settings.
4. Math channels are configured expressions, not a free-form graphical formula editor.
5. Plots are configuration-driven rather than a fully interactive chart designer.
6. Profile Excel reports use native editable scatter charts. PowerPoint and UI previews use deterministic PNG plots; legacy configurable report paths retain their own layout behavior.
7. PowerPoint report content is deterministic and template-driven. Manual annotations present in reference presentations are not automatically inferred.
8. `SETUP_VSM_TOOL.bat` performs an online first-time bootstrap when a validated Python 3.11.x runtime or Python packages are missing. The automatic runtime target is CPython 3.11.9 because it is the supported Python 3.11 Windows binary installer for this release. Company networks that block downloads may require manual Python installation or an offline wheelhouse prepared by the maintainer.
9. The release is not yet a fully standalone/offline Windows executable or MSI installer.
10. AI-assisted KPI/plot recommendations are not included in the deterministic release. If introduced later, AI must remain advisory and must not replace numerical calculations.

11. Diesel is validated against one supplied Test_03 CSV (223 samples, 586 source
    channels) plus synthetic boundary cases. D1/D2/D3 are implementation stages,
    not three independent field datasets. See [Diesel final validation](DIESEL_FINAL_VALIDATION.md).
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
14. Road height is absent in the Diesel reference. Oil/coolant source traces are
    identical in the old CSV; separate oil and water plots are available without
    assuming they are independent physical measurements. A new run's coolant
    mapping still needs confirmation when its source is absent.
    Optional-only PowerPoint slides disappear when their plots are unavailable
    or all-zero. Display rounding does not alter underlying metrics.
15. Diesel Office rendering is validated on the installed Windows Excel and
    PowerPoint versions, not every viewer. Final delivery ZIP creation remains
    separate from validation-only packaging tests.
16. Revised profiles include live Excel integration/cumulative/statistic formulas
    when their source columns are exported. Legacy/unsupported expressions or
    explicitly unexported dependencies retain Python results with equation notes.
    Regenerate reports after source edits to refresh retained results and
    PowerPoint. Excel and PowerPoint remain separate output files.
