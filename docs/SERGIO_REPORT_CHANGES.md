# Sergio report updates

## Editing and reusing Excel charts

Generated plots are native XY Scatter charts with straight lines, real numerical
X values, and no markers. Each chart includes an editable navy 14 pt title, axis
titles (11 pt), tick labels (10 pt), and a compact bottom legend (9 pt) when more
than one series is present. Traces use the existing workbook accent palette with
explicit RGB colors and 2.5 pt widths. White chart/plot areas and subtle primary
Y gridlines are part of the chart itself. Per-chart worksheet banners are removed.

Electric and Hybrid retain a two-column grid and existing chart order. Charts
measure 637 by 360 pixels (637/96 by 3.75 inches at 96 DPI), with one existing
column between them and two rows between chart rows. Engineering data, KPI cells,
worksheet column widths, and row heights are unchanged.

### Engineering axes and editing

Every visible numerical axis explicitly displays tick labels next to the axis in
10 pt dark text, with outward major ticks and its axis title retained. The duplicate
X axis used internally for dual-axis charts is hidden. Limits and major intervals
are calculated from finite plotted data, rounded outward using conventional
1/2/2.5/5/10 steps, targeting 5-7 tick labels. Near-zero speed/time/distance includes
zero; energy, height and SOC retain useful nonzero ranges. Constant data receives
small padding, with a symmetric range for all-zero values. Each Y axis is scaled
independently. Source data is never rounded or changed.

Charts remain editable using Excel's ordinary chart tools. After replacing a
series with substantially different data, use **Format Axis** to reset bounds and
major units to **Automatic**, or enter suitable limits and display precision.
The workbook contains only normal report/mapping/metadata sheets; no chart-template
or instructional worksheet is generated. Shared chart styling helpers remain.

Desktop Excel rendering is not automated in this environment. Saved chart XML,
reopening, source references, and numerical comparisons are checked; manual Excel
review remains necessary for text placement. openpyxl's reader drops plot-area fill
styling on readback, so that fill is verified in the original saved XML. PowerPoint
and UI previews retain their existing matplotlib plots.

## Adding branded PowerPoint slides

In PowerPoint, open the **New Slide** menu and select
**Astauto VSM - Additional content**. Enter a title and caption, then insert text,
tables, or pictures in the available central area. You can also duplicate a slide
created with this layout.

The original 12 report slides remain in place. The additional layout uses their
existing slide master and copies the normal content background, title/caption
styles, Astauto logo, and footer. Its page number is a native PowerPoint field.
The original template files are not modified. python-pptx has no public API for
creating layouts, so the generator isolates the necessary package relationship
and XML work in one helper. Branding remains native shapes/text and the original
logo asset, rather than a background screenshot.

Road Profile and Wheel Steering captions/backgrounds now come from the normal
content slide. Their cards, plots, explanatory note, and other geometry remain
unchanged.

## Shared Electric/Hybrid PowerPoint visual theme

Both profiles now use the approved Electric Astauto template for native slide
backgrounds, title/subtitle geometry, icons, KPI cards, plots, and the reusable
additional-content layout. Profiles still supply their own titles, metadata,
metric IDs, plots, summary content, and N/A/inactive behavior. The original template
files and engineering profile definitions are unchanged.

The cover uses Cambria Bold 46 pt; other titles use Cambria Bold 25 pt. Normal
captions are Calibri 12.5 pt, italic grey. Body text uses Calibri. Slides 1, 3 and 12
reuse Electric's original dark background images; content slides use F4F6FA.
All footers use the Electric wording/version and page-number spacing. Electric's
intentional corrections are the cover footer wording and replacing the historical
Georgia titles on Road Profile/Wheel Steering with Cambria.

Logo validation checks actual canonical image bytes, not merely a picture's
position. Plot-slot validation rejects header pictures. This prevents a stale
Hybrid slot index from replacing the logo with a plot. Shared KPI cards can expand
for content variants without dropping requested metrics. Hybrid's full configured
cover/executive KPI set is retained in the shared layout, whose capacity matches
those configurations.

Desktop PowerPoint export was unavailable during the original shared-theme validation. Review slides 1-12
in PowerPoint, especially the cover, Executive Results, generator/fuel content,
Road Profile N/A panel, steering note, and Summary. Structural validation covers
background media, fonts, coordinates, logos, plot images and metric text.

## Names and auxiliary statistics

Report filenames use the source stem with `.xlsx` and `.pptx` extensions. Machine
display-name overrides affect titles independently. Upload hashes live in parent
directories, preserving source names even when they start with hexadecimal text.
Invalid filename characters are replaced, and reserved Windows device names are
escaped. If an XLSX report would replace its input, it is placed in a `reports`
subdirectory with the same filename.

Both profiles use `MAX(auxiliary_energy_consumption_accumulated)` in kWh and
`MAX(total_auxiliary_power)` in kW. Hybrid inherits these definitions from the
Electric base. Total auxiliary power retains its existing low-voltage plus
high-voltage definition and sign convention. Excel and PowerPoint consume the
same canonical statistics. The combined tyres-plus-auxiliary KPI references the
corrected auxiliary result; tyre rolling-resistance calculations are unchanged.

The two statistics use the existing finite-sample omission policy. Other
statistics retain their existing policy, and source/math validation remains in
force. The auxiliary-energy statistic ID is now
`auxiliary_energy_accumulated_max`; bundled configuration references were updated
together.
