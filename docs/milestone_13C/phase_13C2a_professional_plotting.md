# Milestone 13C.2a - Professional Matplotlib Engineering Plotting

## Purpose

Milestone 13C.2a upgrades the Python plotting path into the static figure system used by standalone exports and future PowerPoint reporting. Matplotlib is used instead of MATLAB because the project is already Python-based, deterministic, portable and installable without a client MATLAB license.

## Plotting Architecture

The validated 70-channel duty-cycle dataset remains the source of truth. Plotting receives the same imported or duty-cycle-composed dataset already used by the pipeline and, when configured, the existing math-channel result.

Excel and Matplotlib now have separate responsibilities:

- Excel report: keeps native Excel scatter charts for interactive workbook use.
- Matplotlib renderer: produces presentation-quality PNG files and optional SVG files for PowerPoint and standalone engineering figures.

The plotting renderer does not modify deterministic values, duty-cycle composition, statistics or KPI calculations.

## Engineering Style Configuration

`config/plotting_example.yaml` defines default figure geometry and a `style:` block. The current engineering standard is:

- Figure size: 11.5 x 6.2 in
- DPI: 180
- White opaque background
- 15 pt titles, 11 pt axis labels, 9 pt ticks and legends
- 1.35 pt anti-aliased traces
- Light engineering grid with restrained alpha
- Combined legends with configurable location and columns
- Tight layout after axis/legend construction

Code defaults remain conservative so existing simple plotting configurations still load.

## Axis Formatting

The renderer applies engineering tick formatting from the visible axis labels:

- Time axes avoid scientific notation and use compact decimal formatting.
- Distance, speed, power, energy, torque, fuel and rpm labels use clean decimal formatting.
- SOC and percent-labelled axes use integer percent-style ticks.
- Zero lines are drawn only when the plotted range crosses zero.

Axis ranges are not hard-coded by default.

## Secondary Axes

Series configured with `axis: secondary` are rendered on a true Matplotlib right-hand Y axis via `twinx()`. The renderer does not normalize, rescale or otherwise alter source values to combine incompatible units.

Primary and secondary legend handles are merged into one readable legend. Secondary series use the configured secondary line style by default.

## Phase Awareness

Duty-cycle phase awareness is optional per plot:

```yaml
show_phase_boundaries: true
show_phase_labels: false
```

When the full duty-cycle pipeline runs, it passes `duty_cycle_provenance.csv` to the plotting stage. Phase boundaries are drawn as subtle vertical dotted lines. Labels can be enabled separately, but are off by default to avoid clutter.

The source-cycle workflow does not require duty-cycle provenance.

## Output Formats

PNG is always required and is written under:

```text
outputs/.../06_plots/png/
```

SVG can be enabled through:

```yaml
style:
  output_formats: [png, svg]
```

SVG files are written under:

```text
outputs/.../06_plots/svg/
```

The plot manifest records each PNG path, optional SVG path, dimensions, DPI, axes count, series IDs and visible legend labels.

## Performance and Memory

The renderer plots the full input trace, including the 17,418-sample Sergio duty cycle. It does not smooth, average, decimate, interpolate or filter values.

Each Matplotlib figure is explicitly closed after saving. Tests verify repeated generation does not accumulate open figures.

## Known Differences From Sergio

The Matplotlib figures are intentionally cleaner than the original Excel/PPT reference visuals. Titles are concise, raw internal channel IDs are removed from visible titles and legends, and units are placed primarily on axes rather than repeated in every legend label.

Sergio's orphan/scratch Excel charts remain excluded from the meaningful report plotting set.

## Future Style Changes

Future client/report styles should be introduced through the plotting YAML style block where possible. New constants should not be scattered through rendering code unless they are stable engine defaults.

Full PowerPoint layout fidelity remains reserved for Milestone 13C.2b / v1.3.0.
