# PowerPoint plot-container correction — 4 November 2026

The shared Electric/Hybrid generator now applies the approved slide-9 style to
chart background containers. KPI cards and chart pictures are unchanged.

## Implementation

- [profile_powerpoint_report_engine.py](../src/vsm_postprocessing/profile_powerpoint_report_engine.py):
  `_add_custom_plot_panel` calls `_style_plot_background_card` for both single
  and dual layouts, including Motor Torque & Power and Torque–Speed Maps.
- The helper sets `roundRect` adjustment `1617` (0.01617), white fill, an explicit
  invisible outline, and the approved outer shadow. Shadow attributes are
  `blurRad=152400`, `dist=50800`, `dir=5400000`, `sx=sy=100000`, `kx=ky=0`,
  `algn=bl`, `rotWithShape=0`, colour `0A1524`, alpha `13000`.
- Theme `p:style` references are removed from these containers. Applying the
  helper again replaces the explicit effect rather than adding another shadow.
- The same existing shape is styled; picture insertion, coordinates, insets and
  stacking order are unchanged. Production code does not select slide numbers.
- [test_plot_background_cards.py](../tests/test_plot_background_cards.py) adds six
  regression cases covering exact XML style, both layouts/profiles, picture
  counts/bytes/bounds, slide boundaries, KPI preservation, notes and engineering
  data, idempotence, and unavailable-image handling.

The authoritative reference was the user's original Downloads deck:
`RoboSprayer_8500Kg_Electric_03_UoL_DC_20kph_DriverPC_25pcGrad_FL-RL_Rough (13).pptx`.
Its SHA-256 remains
`0c10a4ea062eb8144a17474bf1603fe2e06aa3a1eb02c480cdce5096fdb12ce1`.
The corrected ChatGPT review deck was not available locally; the original
slide-9 XML supplied the styling reference.

## Regenerated reports and visual evidence

The real Electric CSV was uploaded through the current Streamlit UI, with the
same 23-point peak/continuous motor curves, 1700 kg battery mass, 17.5 ratio and
0.97 efficiency. Validation and report generation completed without a UI
exception, using the workspace Python environment.

- [Fresh Electric PowerPoint](../outputs/plot_cards_2026_11_04/browser_reports/20261104_175651_530870/profile_powerpoint_report/RoboSprayer_8500Kg_Electric_03_UoL_DC_20kph_DriverPC_25pcGrad_FL-RL_Rough.pptx)
- [Fresh Electric Excel](../outputs/plot_cards_2026_11_04/browser_reports/20261104_175651_530870/profile_excel_report/RoboSprayer_8500Kg_Electric_03_UoL_DC_20kph_DriverPC_25pcGrad_FL-RL_Rough.xlsx)
- [Before/after visual comparison, slides 9–16](../outputs/plot_cards_2026_11_04/visual_comparison.html)
- [Native PowerPoint PDF export](../outputs/plot_cards_2026_11_04/native_review/after.pdf)
- [Road Profile screenshot](../outputs/plot_cards_2026_11_04/native_review/after/slide_10.png)
- [Motor Torque & Power screenshot](../outputs/plot_cards_2026_11_04/native_review/after/slide_15.png)
- [Synthetic Hybrid PowerPoint](../outputs/plot_cards_2026_11_04/SYNTHETIC_hybrid/ppt/SYNTHETIC_hybrid.pptx)
- [Synthetic Hybrid PDF](../outputs/plot_cards_2026_11_04/native_review/SYNTHETIC_hybrid.pdf)

All Electric slides 9–16 were visually inspected after native Microsoft
PowerPoint export. Road Profile and Motor Torque & Power have clean, small
corners and subtle shadows. The dual-chart layouts retain their original insets.
Hybrid Motor Torque & Power and Torque–Speed Maps were also visually inspected.
**Hybrid validation uses a synthetic fixture, not real Hybrid engineering data.**

## Preservation checks

The fresh Electric presentation retains 17 slides. Exactly 11 background cards
changed: one each on slides 10, 11 and 15, and two each on 12, 13, 14 and 16.
All shape counts, bounds, IDs, names and text are unchanged. All 58 embedded
pictures are byte-identical; no picture or card was duplicated. All other shape
XML is identical, including KPI cards, titles, placeholders, logos and footers.
The 30 native Excel chart definitions are also byte-identical. A comparison of
803,935 populated cells across all seven Excel sheets found only two differences:
the generation timestamp (`Metadata!B3`) and uploaded-source path
(`Motor Duty Summary!A152`). All engineering values, formulas and labels match;
see [cell comparison](../outputs/plot_cards_2026_11_04/excel_cell_comparison.json).

Native 1600×900 renders show slides 1–9 and 17 pixel-identical to the reference.
Differences on slides 10–16 are confined to the plot containers and their
shadows. See [pixel comparison](../outputs/plot_cards_2026_11_04/pixel_comparison.json)
and [artifact audit](../outputs/plot_cards_2026_11_04/artifact_audit.json).

Engineering content in notes is unchanged. Fresh-generation provenance paths
and the generated-workbook checksum naturally identify the new run; the audit
records these differences explicitly and checks that the uploaded CSV bytes
match. The isolated same-input before/after tests preserve notes exactly.
The synthetic Hybrid deck retains 16 slides and changes only nine plot cards;
its pictures and notes are unchanged.

## Tests and warnings

- Focused tests: **6 passed**.
- Full regression suite: **1,116 passed, 7 skipped, 0 failures/errors**, out of
  1,123 tests; 1,406.606 seconds (23 minutes 27 seconds). See
  [JUnit results](../outputs/plot_cards_2026_11_04/full_suite.xml) and
  [test log](../outputs/plot_cards_2026_11_04/full_suite.log).
- `git diff --check` passes. Git reports existing LF-to-CRLF notices for two
  previously modified test files outside this change.

The seven skips comprise six client-ZIP tests (disabled to respect the user's
instruction) and one existing Diesel soil-applicability case. Pytest reports
two expected warnings from an existing sparse synthetic motor-curve test: only
one nonzero same-speed check is available for each of peak and continuous
torque, whereas release acceptance requires three. These are not warnings from
the actual 23-point Electric report. Captured Streamlit bare-mode test logs also
contain the existing missing-ScriptRunContext notices.

The actual Electric report retains the existing warning
that density (kg/L) is required for mass-to-volume conversion; this styling
change does not alter that calculation or its inputs.

## Regenerating clean outputs

Restart the workspace UI with `START_VSM_TOOL.bat` so it loads the updated
generator. Upload the CSV,
select the Electric or Hybrid profile, and enter the original run's inputs and
motor peak/continuous reference curves. Validate, generate the Excel and
PowerPoint reports, and download the newly generated PowerPoint. Existing
downloaded decks are not modified automatically.

To rerun the focused tests from the repository root in PowerShell:

```powershell
.venv\Scripts\python.exe -m pytest tests/test_plot_background_cards.py -q
```

For the complete suite while preserving the no-client-ZIP constraint, use the
review runner at `outputs/plot_cards_2026_11_04/run_suite.py` with a fresh
`--basetemp` directory in its pytest arguments.

No commit, push, tag or client ZIP was created for this change.
