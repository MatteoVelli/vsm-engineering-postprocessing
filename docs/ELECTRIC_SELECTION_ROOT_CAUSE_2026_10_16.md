# Electric selection failure: engineering assessment, 16 October 2026

This records the selection investigation before Sergio's subsequent HV-unit
confirmation. Its HV-unavailable findings are historical; the confirmed
interpretation and restored outputs are documented in
[the later HV assessment](HV_AUXILIARY_CONFIRMATION_2026_10_16.md).

## A. Established root cause

The source CSV is valid. The failure is a loss of channel identity during the
Streamlit multiselect round trip, before report data resolution or calculations.
The local environment contains Streamlit **1.61.1**, outside the project's declared
`>=1.49.1,<1.50` range. Its `MultiSelectSerde.serialize()` transports formatted
strings; `deserialize()` resolves each string using the first matching formatted
option. `_render_profile_report_options()` previously formatted channels using
only the profile's report name and channel type.

The Electric and Hybrid profiles, including their inherited October revisions,
contain these four pairs of distinct, valid semantic identifiers with identical
presentation labels:

| Display label | Distinct semantic identifiers |
| --- | --- |
| Suspension Damper Force FL [VSM] | `suspension_damper_force_fl`, `suspension_damper_force_fr` |
| Suspension Damper Travel FL [VSM] | `suspension_damper_travel_fl`, `suspension_damper_travel_fr` |
| Tyre Rolling Resistance Power FR [VSM] | `tyre_rollingresistancepower_fr`, `tyre_rollingresistancepower_rr` |
| SoftSoil Shear Displacement Y IN FL [AVL] | `softsoil_sheardisplacementy_in_fl`, `softsoil_sheardisplacementy_in_rr` |

The initial export defaults contain 330 unique identifiers. After a UI round trip,
the list still has 330 entries but only 326 distinct identifiers: the first member
of each pair occurs twice and the second member is absent. No duplicate source
channel or duplicate semantic identifier in the profile caused this error.
Diesel's default channel labels are unique, explaining its successful workflow.
Custom duplicate plot titles expose the same weakness in plot selectors.

The captured real-file traceback follows this path:

1. `ui_config.generate_reporting_profile_engineering_report()`
2. `excel_report_engine.generate_profile_excel_report()`
3. `report_selection.select_profile()` / nested `chosen()`
4. `ConfigurationError: Invalid export channels selection: [] (duplicates are not allowed)`

Step 3 validates the source/profile resolution, which succeeds; it does not apply
the Step 4 report selection. The old exception printed only *unknown* identifiers,
so an empty unknown list concealed the actual duplicate identifiers.

Evidence: `outputs/electric_uol_2026_10_16/reproduction.json` contains all profile
label collisions, the actual corrupted selection, successful validation, the full
traceback, enabled Generate button, and the displayed failure. The real CSV has
608 channels, 3,954 samples and SHA-256
`773e632fcabad1759441de495167fd43b088c3819c3851533419637a84950681`.

## B. Alternatives and decision

| Alternative | Assessment |
| --- | --- |
| Deduplicate the backend selection | Reject: silently preserves the wrong member of each pair and loses four selected channels. |
| Disable duplicate validation | Reject: hides corruption and permits invalid API input. |
| Rename only the four profile labels | Incomplete: changes inherited report labels and leaves custom duplicate channel names or plot titles vulnerable. |
| Add widget keys/session-state repair only | Insufficient: widget identity does not make its formatted option strings unique; identity already lost cannot be reconstructed reliably. |
| Downgrade the active environment only | Restores the declared runtime but leaves selection identity dependent on implementation-specific widget serialization. |
| Include immutable identifiers in every formatted option | Selected: deterministic, traceable, compatible with both tested Streamlit versions, and independent of configurable presentation names. |

## C. Implementation and scope

- `src/vsm_postprocessing/ui_app.py`: channel labels include semantic identifiers;
  Excel and PowerPoint plot selectors include plot IDs. Precomputed label maps
  use the same identity convention in every affected widget.
- `src/vsm_postprocessing/report_selection.py`: strict validation remains in place;
  errors separately identify unknown and duplicate identifiers, including both
  when applicable.
- `tests/test_profile_selection_roundtrip.py`: six regression cases cover all three
  powertrains, repeated channel labels, custom repeated plot titles, repeated
  reruns, individual deselection, intentional empty selection and invalid IDs.
- `scripts/validate_electric_uol.py`: reproducible private-file UI acceptance and
  independent CSV/report audit; refuses existing output directories.

No changes were made for this correction to channel source mappings, formulas,
physical-unit validation, profile inheritance, report templates or dependency
bounds. Existing unrelated local changes were preserved. No commit or push.
Testing packages were installed only into ignored output folders; the application
environment was left unchanged.

## D. Validation evidence

Before the correction, the real upload reproduced the exact reported failure.
The six new regressions all failed against the previous implementation. After the
correction, UI reruns preserve all 330 requested identifiers and Step 4 produces
Excel and PowerPoint reports. Unavailable optional outputs are diagnosed rather
than fabricated.

The accepted real-file artifacts and machine-readable audit are under
`outputs/electric_uol_2026_10_16/corrected/`. Independent checks parse the CSV
directly, separately from the production importer and expression evaluator:

- All four `KERS_MotGenUnit_GearboxOutputTorque_1..4` arrays match their respective
  corners, and their sum matches total wheel torque.
- Signed corner wheel power matches `torque_Nm * rpm * 2*pi / 60000`; total wheel
  power ranges from **-5.3701880862 to 31.2520518996 kW**. Maximum independent
  wheel-power difference is approximately `1.07e-14 kW`.
- Rolling, bulldozing and compaction powers sum the four corners. Interval energy
  uses trapezoidal integration against actual CSV timestamps, with zero initial
  increment; accumulated energy is the cumulative sum. Total accumulated tyre
  resistance energy ends at **28.0835308011 kWh** (maximum difference `3.56e-15`).
- All 326 available selected channels are exported, including both members of
  each formerly colliding pair: 293 raw and 33 MATH columns.
- **1,289,004** numeric data cells match the calculation arrays; **130,482** live
  data formulas remain in the workbook. Twenty-six statistics, including both
  RMS values, were independently recomputed from raw or independently derived
  arrays. There are 33 available summary statistics, 11 KPIs and 18 Excel plots.
- Desktop Excel independently recalculated a disposable copy after all **131,856**
  formula caches were removed and verified empty. **1,290,838** numeric cells across
  all worksheets agree within **4.55e-13**; no Excel error cells were found.
- The generated PowerPoint contains 12 slides: the 10-slide reference template
  plus the existing Road Profile and Wheel Steering additions. The template
  fidelity check passes for dimensions, 10 text slots, 67 KPI cards and 13 plot
  slots. Desktop PowerPoint successfully rendered the entire deck to PDF; slide
  9 was visually inspected for layout, signed traction curves, typography,
  colours, units and the 28.08 kWh total tyre-energy card.

The complete 35-module suite collected 532 cases. All 531 executable cases pass;
one pre-existing unsupported Diesel soil-slide case is skipped. The initial run
had 530 passes and one failure while `venv` ran `ensurepip` in its deeply nested
temporary directory. That bootstrap test passes using a shorter temporary path;
the entire 20-case bootstrap module also passes there. The original failure and
the rerun were retained; no test was weakened. Suite evidence is in `full_suite/`
and `bootstrap_short_path.txt` under the evidence root.

Supported-runtime checks also pass: **eight UI/selection tests** under isolated
Streamlit 1.49.1, including actual Diesel report generation, and the complete real
Electric upload, generation and numerical audit. The latter produced another
326-column workbook and 12-slide deck under `supported_149/`, with evidence in
`supported_149_tests_complete.txt` and `supported_149_e2e_complete.txt`. The
application runtime remains 1.61.1; isolated testing uses the declared 1.49.1
runtime and its dependencies.
The generated primary worksheet XML, workbook styles, all 18 chart XML parts and
all 12 PowerPoint slide XML parts are byte-identical between the two runtime
versions. `verification_summary.json` records the comparison and test outcomes.

Reproduce the actual upload and independent audit (use a fresh destination):

```powershell
.venv\Scripts\python.exe scripts/validate_electric_uol.py `
  '<path-to-private-electric-vsm.csv>' `
  --output-dir outputs/electric_uol_new_validation
```

The upload is supplied to the real Streamlit application through AppTest's
file-uploader substitute; actual validation, widgets, report orchestration,
Excel generation and PowerPoint generation execute normally. No browser file
upload gesture is claimed.

## E. Outstanding source and environment concerns

The HV source explicitly declares `-`, not `kW`, and ranges from 0 to 10 in those
unconfirmed units. Its resemblance to older reference values is insufficient
evidence of physical units. Conservative handling is unchanged: the HV channel,
total auxiliary power, its two energy channels and their dependent results are
unavailable. The optional warning is present, while independent report sections
remain available. Sergio must confirm the physical units or provide a corrected
export before this auxiliary chain can be validated.

The existing battery-capacity warning also remains: the Energy/SOC-derived median
nominal capacity is 98.2263 kWh, with maximum deviation 1.86791 kWh exceeding the
configured 0.982263 kWh tolerance. No battery-capacity correction is inferred.

The four inherited display-name collisions remain in approved report labels;
the UI now disambiguates them with their semantic identifiers. Any correction to
those report-facing corner labels should be confirmed against the authoritative
template. The local 1.61.1 installation is outside the declared support range;
the compatibility test does not change or broaden that dependency contract.
