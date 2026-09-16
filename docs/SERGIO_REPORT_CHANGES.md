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

### Create an additional graph

1. Open the visible **Plot Templates** sheet.
2. Select the border of a Single Series, Multiple Series (four traces), or Dual Y
   Axes template. Copy with Ctrl+C and paste with Ctrl+V into your chosen worksheet.
3. Open **Chart Design > Select Data**. Select an existing series and choose
   **Edit**; set its name, numerical X range, and Y range. X/Y lengths must match.
4. Repeat for other series; remove unused series. Editing the existing series
   preserves its formatting. Additional series may need their color/width set;
   a four-series template avoids that setup for common wheel plots.
5. Edit the native chart title and axis titles/units. For a dual-axis chart, keep
   each channel on the correct primary or secondary Y axis.

To modify an existing report graph, follow steps 3-5 directly on that graph.
Templates use clearly labeled sample data on their own sheet, never production
KPI cells. They share the exact report chart builder and styling helpers.

### Scaling and Excel behavior

Axis bounds and major intervals are automatic so a copied/repointed chart can
adapt to new ranges. Engineering number formats retain sensible unit/range-based
precision for generated data. After a substantial scale or unit change, use
**Format Axis > Number** to adjust displayed precision. Excel chooses tick density;
5-7 labels are a readability target rather than a fixed constraint. No source data
is rounded. Both X axes in a dual-axis chart reference the same numerical values.

Inserting a brand-new chart through Excel's Insert menu still uses Excel defaults.
Use the supplied charts to reuse the complete style. A built-in style ID alone does
not include our explicit formatting or register a custom preset. Excel's optional
[Save as Template workflow](https://support.microsoft.com/en-au/excel/save-a-custom-chart-as-a-template)
creates a separate `.crtx` file; this project does not generate or install one.
No macros or external setup are required for the supplied template sheet.

Legacy configurable reports receive the same chart styling and template sheet.
Plot-only channels excluded from their report table remain on a hidden `Chart Data`
sheet. Existing PowerPoint and UI previews retain their matplotlib images.

Desktop Office rendering is not automated in this environment. Saved chart XML,
reopening, source references, and numerical comparisons are checked; manual Excel
review remains necessary for text placement and interaction. openpyxl's reader
drops plot-area fill styling on readback, so that fill is verified in the original
saved XML. Excel itself controls automatic plot/legend layout. Copy charts as
chart objects, not pictures, and use source formatting when pasting across workbooks.

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
