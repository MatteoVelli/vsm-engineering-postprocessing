from __future__ import annotations

import hashlib
import os
from dataclasses import asdict, replace
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

from vsm_postprocessing.errors import VSMPostProcessingError, DataValidationError
from vsm_postprocessing.diagnostics import as_diagnostic
from vsm_postprocessing.importer import ImportOptions, inspect_data_file
from vsm_postprocessing.pipeline_engine import run_pipeline
from vsm_postprocessing.report_metadata import resolve_report_metadata
from vsm_postprocessing.version import __version__
from vsm_postprocessing.utils import source_report_filename, normalize_display_unit
from vsm_postprocessing.report_profile import load_reporting_profile
from vsm_postprocessing.report_selection import ProfileReportSelection
from vsm_postprocessing.ui_config import (
    available_math_channel_ids,
    build_runtime_bundle,
    default_ui_profile,
    discover_reporting_profiles,
    generate_reporting_profile_engineering_report,
    get_reporting_profile_definition,
    load_ui_profile,
    load_ui_templates,
    supported_profile_upload_extensions,
    save_ui_profile,
    validate_profile_upload_extension,
    validate_reporting_profile_source,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
PROFILE_PATH = PROJECT_ROOT / "config" / "ui_saved_profile.yaml"
UI_WORKSPACE = PROJECT_ROOT / "outputs" / "ui_workspace"
UI_RUNS = PROJECT_ROOT / "outputs" / "ui_runs"


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="VSM Post-Processing", page_icon="📊", layout="wide")
    st.title("VSM Engineering Post-Processing")
    st.caption(f"v{__version__} | Deterministic processing: source → channels → math → statistics → plots → Excel + optional PowerPoint reports")

    templates = load_ui_templates(PROJECT_ROOT)
    defaults = default_ui_profile(templates)
    profile = load_ui_profile(PROFILE_PATH, fallback=defaults)

    processing_mode = st.radio(
        "Workflow",
        ["Engineering Report", "Custom Analysis"],
        horizontal=True,
    )
    if processing_mode == "Engineering Report":
        _render_engineering_report_workflow(st)
        return

    st.header("Custom Analysis")
    uploaded = st.file_uploader("1. Load VSM results", type=["xlsx", "csv"])
    st.caption("The original file is copied into the local outputs workspace; it is not modified.")
    if uploaded is None:
        st.info("Choose a VSM CSV/XLSX file to inspect its channels and configure the report.")
        return

    source_path = _persist_upload(uploaded)
    try:
        inspection = inspect_data_file(source_path, ImportOptions(strict=True))
    except VSMPostProcessingError as exc:
        st.error(f"Input inspection failed: {exc}")
        return

    quality = inspection.quality
    time_id = quality.time_channel_id
    if time_id is None:
        st.error("No time channel was detected. The current UI requires a valid time channel.")
        return

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Samples", quality.sample_count)
    c2.metric("Channels", quality.channel_count)
    c3.metric("Raw", quality.raw_channel_count)
    c4.metric("Imported math", quality.math_channel_count)
    st.caption(
        f"Time channel: {time_id} | {_fmt_number(quality.time_start)} to {_fmt_number(quality.time_end)} | nominal step {_fmt_number(quality.nominal_time_step)}"
    )

    source_channels = {channel.channel_id: channel for channel in inspection.channels}
    all_math_defs = {item["channel_id"]: item for item in templates.math_channels["math_channels"]}
    stat_defs = {item["statistic_id"]: item for item in templates.statistics["statistics"]}
    all_plot_defs = {item["plot_id"]: item for item in templates.plotting["plots"]}

    source_options = list(source_channels)
    available_math = available_math_channel_ids(templates.math_channels, source_options)
    math_defs = {channel_id: all_math_defs[channel_id] for channel_id in available_math}
    available_channel_ids = set(source_options) | set(available_math)
    available_stat_ids = [
        sid for sid, definition in stat_defs.items()
        if definition["channel_id"] in available_channel_ids
    ]
    plot_defs = {
        pid: definition for pid, definition in all_plot_defs.items()
        if _plot_is_available(definition, available_channel_ids)
    }
    report_options = [*source_options, *available_math]

    def channel_label(channel_id: str) -> str:
        if channel_id in source_channels:
            channel = source_channels[channel_id]
            unit = f" [{channel.unit}]" if channel.unit else ""
            return f"{channel.display_name}{unit} — {channel_id}"
        definition = math_defs[channel_id]
        unit = f" [{definition.get('unit')}]" if definition.get("unit") else ""
        return f"MATH: {definition['display_name']}{unit} — {channel_id}"

    tabs = st.tabs(["Export channels", "Math channels", "Statistics", "Plots", "Excel report", "PowerPoint"])
    with tabs[0]:
        export_ids = st.multiselect(
            "Channels to export",
            source_options,
            default=_valid_defaults(profile.get("export_channels", []), source_options),
            format_func=channel_label,
            help="The time channel is inserted automatically by the channel-selection engine.",
        )
    with tabs[1]:
        math_ids = st.multiselect(
            "Math channels to calculate",
            list(math_defs),
            default=_valid_defaults(profile.get("math_channels", []), list(math_defs)),
            format_func=lambda cid: f"{math_defs[cid]['display_name']} [{math_defs[cid].get('unit') or '-'}] — {cid}",
            help="Math dependencies required by selected statistics, plots or report columns are added automatically.",
        )
        if math_ids:
            st.dataframe(
                [
                    {
                        "channel_id": cid,
                        "name": math_defs[cid]["display_name"],
                        "unit": math_defs[cid].get("unit"),
                        "expression": math_defs[cid]["expression"],
                    }
                    for cid in math_ids
                ],
                width="stretch",
                hide_index=True,
            )
    with tabs[2]:
        statistic_ids = st.multiselect(
            "Statistics to calculate",
            available_stat_ids,
            default=_valid_defaults(profile.get("statistics", []), available_stat_ids),
            format_func=lambda sid: _stat_label(stat_defs[sid]),
        )
        kpi_candidates = list(statistic_ids)
        kpi_ids = st.multiselect(
            "Statistics to show in the KPI strip",
            kpi_candidates,
            default=_valid_defaults(profile.get("kpis", []), kpi_candidates),
            format_func=lambda sid: _stat_label(stat_defs[sid]),
        )
    with tabs[3]:
        plot_ids = st.multiselect(
            "Plots to generate",
            list(plot_defs),
            default=_valid_defaults(profile.get("plots", []), list(plot_defs)),
            format_func=lambda pid: f"{plot_defs[pid]['title']} — {pid}",
        )
    with tabs[4]:
        report_ids = st.multiselect(
            "Channels shown in the Excel report",
            report_options,
            default=_valid_defaults(profile.get("report_channels", []), report_options),
            format_func=channel_label,
            help="RAW and MATH channels remain visually distinct in the generated workbook.",
        )
        st.caption("RMS statistics are placed above their channel; MAX/MIN/last/SUM are placed below it when the target channel is visible.")

    with tabs[5]:
        generate_powerpoint = st.checkbox(
            "Generate PowerPoint report",
            value=bool(profile.get("generate_powerpoint", True)),
            help="Uses the selected statistics and plots. No AI selection is applied.",
        )
        st.caption(
            "The presentation follows the useful pattern in Sergio's reference deck: KPI strip plus one or two engineering plots per slide."
        )

    current_profile = {
        "version": 1,
        "export_channels": list(export_ids),
        "math_channels": list(math_ids),
        "report_channels": list(report_ids),
        "statistics": list(statistic_ids),
        "kpis": list(kpi_ids),
        "plots": list(plot_ids),
        "generate_powerpoint": bool(generate_powerpoint),
    }

    left, right = st.columns([1, 2])
    with left:
        if st.button("Save selections", width="stretch"):
            try:
                save_ui_profile(PROFILE_PATH, current_profile)
            except VSMPostProcessingError as exc:
                st.error(str(exc))
            else:
                st.success(f"Saved: {PROFILE_PATH.relative_to(PROJECT_ROOT)}")
    with right:
        run_clicked = st.button("Run complete pipeline", type="primary", width="stretch")

    if run_clicked:
        if not export_ids:
            st.error("Select at least one export channel.")
            return
        if not statistic_ids:
            st.error("Select at least one statistic.")
            return
        if not plot_ids:
            st.error("Select at least one plot.")
            return
        if not report_ids:
            st.error("Select at least one Excel report channel.")
            return

        run_dir = UI_RUNS / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        try:
            bundle = build_runtime_bundle(
                source_file=source_path,
                runtime_dir=run_dir,
                templates=templates,
                time_channel_id=time_id,
                export_channel_ids=export_ids,
                selected_math_channel_ids=math_ids,
                report_channel_ids=report_ids,
                selected_statistic_ids=statistic_ids,
                kpi_statistic_ids=kpi_ids,
                selected_plot_ids=plot_ids,
                include_powerpoint=generate_powerpoint,
            )
            with st.spinner("Running deterministic VSM pipeline..."):
                result = run_pipeline(bundle.pipeline_config)
        except VSMPostProcessingError as exc:
            st.error(f"Pipeline failed: {exc}")
            return

        st.session_state["last_report"] = str(result.report_path) if result.report_path else None
        st.session_state["last_powerpoint"] = str(result.powerpoint_path) if result.powerpoint_path else None
        st.session_state["last_manifest"] = str(result.manifest_path)
        st.session_state["last_run_dir"] = str(run_dir)
        st.success(f"Pipeline completed: {result.completed_stage_count}/{len(result.stages)} stages PASS")
        st.dataframe(
            [
                {
                    "stage": stage.name,
                    "status": stage.status,
                    **stage.metrics,
                }
                for stage in result.stages
            ],
            width="stretch",
            hide_index=True,
        )
        if set(bundle.effective_math_channel_ids) != set(math_ids):
            auto_added = [cid for cid in bundle.effective_math_channel_ids if cid not in math_ids]
            if auto_added:
                st.info("Math dependencies added automatically: " + ", ".join(auto_added))

    report_path_text = st.session_state.get("last_report")
    powerpoint_path_text = st.session_state.get("last_powerpoint")
    if report_path_text or powerpoint_path_text:
        st.subheader("Final reports")
    if report_path_text:
        report_path = Path(report_path_text)
        if report_path.exists():
            st.markdown("**Excel engineering report**")
            st.code(str(report_path))
            with report_path.open("rb") as handle:
                report_bytes = handle.read()
            col_a, col_b = st.columns(2)
            col_a.download_button(
                "Download Excel report",
                data=report_bytes,
                file_name=report_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )
            if os.name == "nt" and col_b.button("Open report in Excel", width="stretch"):
                os.startfile(report_path)  # type: ignore[attr-defined]
    if powerpoint_path_text:
        powerpoint_path = Path(powerpoint_path_text)
        if powerpoint_path.exists():
            st.markdown("**PowerPoint engineering report**")
            st.code(str(powerpoint_path))
            with powerpoint_path.open("rb") as handle:
                powerpoint_bytes = handle.read()
            col_c, col_d = st.columns(2)
            col_c.download_button(
                "Download PowerPoint report",
                data=powerpoint_bytes,
                file_name=powerpoint_path.name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                width="stretch",
            )
            if os.name == "nt" and col_d.button("Open report in PowerPoint", width="stretch"):
                os.startfile(powerpoint_path)  # type: ignore[attr-defined]


def _render_engineering_report_workflow(st: Any) -> None:
    st.header("Engineering Report")
    _render_profile_engineering_report_workflow(st)


def _render_profile_engineering_report_workflow(st: Any) -> None:
    st.write("Generate validated Electric, Hybrid or Diesel Excel and PowerPoint engineering reports from one VSM result file.")

    try:
        profiles = discover_reporting_profiles(PROJECT_ROOT)
    except VSMPostProcessingError as exc:
        st.error("Reporting profiles could not be loaded.")
        with st.expander("Technical details"):
            st.write(str(exc))
        return
    if not profiles:
        st.error("No reporting profiles are configured.")
        return

    uploaded = st.file_uploader(
        "STEP 1 - Upload VSM Result",
        type=[suffix.lstrip(".") for suffix in supported_profile_upload_extensions()],
        help="Supported formats: CSV and XLSX.",
        key="profile_report_source",
    )
    if uploaded is not None:
        upload_identity = (uploaded.name, len(uploaded.getvalue()))
        if st.session_state.get('auto_profile_source') != upload_identity:
            detected_powertrain = resolve_report_metadata(uploaded.name).powertrain_name.lower()
            candidates = [p.profile_id for p in profiles if p.powertrain == detected_powertrain]
            if len(candidates) == 1:
                st.session_state['engineering_profile_choice'] = candidates[0]
            st.session_state['auto_profile_source'] = upload_identity
    selected_profile_id = st.selectbox(
        "STEP 2 - Select Report Profile",
        [profile.profile_id for profile in profiles],
        format_func=lambda profile_id: next(profile.display_name for profile in profiles if profile.profile_id == profile_id),
        help="Detected from the source identity when available. Change it here if necessary; inactive channels never determine the powertrain.",
        key="engineering_profile_choice",
    )
    profile_definition = get_reporting_profile_definition(PROJECT_ROOT, selected_profile_id)

    source_path: Path | None = None
    validation_key: str | None = None
    machine_name_override: str | None = None
    selection = None
    report_import_options = ImportOptions(strict=True)
    reporting_inputs = None
    if uploaded is None:
        st.info("Upload a supported VSM CSV or XLSX file to validate it against the selected profile.")
    else:
        source_path = _persist_upload(uploaded)
        detected_metadata = resolve_report_metadata(source_path, profile_definition)
        st.caption(f"Detected machine: {detected_metadata.machine_name}")
        machine_name_override = st.text_input(
            "Machine / Vehicle Name",
            value=detected_metadata.machine_name,
            help="This display name is used in report titles. The selected powertrain profile remains separate.",
        )
        validation_key = f"{source_path}:{source_path.stat().st_size}:{profile_definition.profile_id}"
        try:
            validate_profile_upload_extension(source_path)
            inspection = inspect_data_file(source_path, ImportOptions(strict=False))
            if not inspection.quality.time_is_strictly_increasing or inspection.quality.sample_count < 2:
                raise DataValidationError("At least two finite, strictly increasing time samples are required")
        except VSMPostProcessingError as exc:
            st.error(_friendly_input_error(exc))
            with st.expander("Diagnostics"):
                st.write(str(exc))
            source_path = None
            validation_key = None
        else:
            quality = inspection.quality
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Samples", f"{quality.sample_count:,}")
            c2.metric("Source channels", f"{quality.channel_count:,}")
            c3.metric("Raw channels", f"{quality.raw_channel_count:,}")
            c4.metric("Duration", _format_duration_minutes(quality.time_start, quality.time_end, quality.time_unit))
            st.caption(f"Time channel: {quality.time_channel_name or quality.time_channel_id or 'not detected'}")
            revision = st.selectbox("Reporting definitions", ["Sergio - 15 October 2026", "Legacy - Electric_05 / Hybrid_06 / Diesel"],
                                    help="Legacy mappings and energy conventions are explicit; sources are never silently substituted.")
            if revision.startswith("Sergio"):
                revised = profile_definition.profile_path.parent / "sergio_2026_10" / profile_definition.profile_path.name
                profile_definition = replace(profile_definition, profile_path=revised)
                profile_definition, selection, report_import_options = _render_profile_report_options(st, profile_definition, inspection)
            else:
                st.info("Legacy reporting uses the previous source mappings and sample-energy convention. The legacy high-voltage auxiliary definition is a constant zero, not a measured high-voltage source.")
            reporting_inputs = _render_reporting_metadata_inputs(st, profile_definition, source_path)
            settings = repr((profile_definition.profile_path, selection, report_import_options, reporting_inputs, machine_name_override))
            validation_key += ":" + hashlib.sha256(settings.encode()).hexdigest()

    if st.session_state.get("profile_input_key") != validation_key:
        _clear_profile_report_state(st)
        st.session_state["profile_input_key"] = validation_key

    validate_clicked = st.button(
        "STEP 3 - Validate",
        width="stretch",
        disabled=source_path is None,
    )
    if validate_clicked and source_path is not None and validation_key is not None:
        _clear_profile_report_state(st)
        try:
            with st.spinner("Validating source channels against reporting profile..."):
                resolve_report_metadata(source_path, load_reporting_profile(profile_definition.profile_path),
                                        machine_name_override=machine_name_override, reporting_inputs=reporting_inputs)
                summary = validate_reporting_profile_source(source_path, profile_definition.profile_path, report_import_options)
        except VSMPostProcessingError as exc:
            st.error(_friendly_profile_error(exc))
            with st.expander("Diagnostics"):
                st.write(str(exc))
        else:
            st.session_state["profile_validation_key"] = validation_key
            st.session_state["profile_validation_summary"] = summary

    summary = st.session_state.get("profile_validation_summary")
    if summary is not None and st.session_state.get("profile_validation_key") == validation_key:
        _render_profile_validation_summary(st, summary)
    elif source_path is not None:
        st.info("Validate the uploaded file and selected profile before generating the report.")

    if profile_definition.powertrain == "diesel":
        _render_diesel_report_workflow(st, profile_definition, source_path, validation_key, summary, machine_name_override,
                                      selection=selection, import_options=report_import_options, reporting_inputs=reporting_inputs)
        return

    can_generate = (
        source_path is not None
        and summary is not None
        and st.session_state.get("profile_validation_key") == validation_key
        and summary.is_valid
    )
    run_clicked = st.button(
        "STEP 4 - Generate Engineering Report",
        type="primary",
        width="stretch",
        disabled=not can_generate,
    )
    if run_clicked and source_path is not None:
        if not machine_name_override or not machine_name_override.strip():
            st.error("Enter a machine / vehicle name before generating reports.")
            return
        run_dir = UI_RUNS / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        try:
            with st.spinner("Generating profile-driven Excel and PowerPoint engineering reports..."):
                result = generate_reporting_profile_engineering_report(
                    source_path,
                    profile_definition.profile_path,
                    run_dir,
                    report_import_options,
                    machine_name_override=machine_name_override,
                    selection=selection,
                    reporting_inputs=reporting_inputs,
                )
        except VSMPostProcessingError as exc:
            st.error(_friendly_profile_error(exc))
            with st.expander("Diagnostics"):
                st.write(str(exc))
            return

        st.session_state["profile_report_result"] = result
        st.session_state["profile_report_result_key"] = (validation_key, machine_name_override)
        st.success(f"Engineering reports generated: {result.report_path.name} and {result.presentation_path.name}")
        _render_profile_report_completion(st, result)

    result = st.session_state.get("profile_report_result")
    if result is not None and st.session_state.get("profile_report_result_key") == (validation_key, machine_name_override):
        st.subheader("STEP 5 - Access Output")
        _render_profile_report_download(st, result)


def _render_reporting_metadata_inputs(st, definition, source_path=None):
    profile = load_reporting_profile(definition.profile_path)
    detected = resolve_report_metadata(source_path or 'Unknown.csv', profile)
    defaults = asdict(detected.reporting_inputs)
    source = Path(source_path) if source_path else None
    revision = (source.stat().st_size, source.stat().st_mtime_ns) if source and source.is_file() else None
    scope = hashlib.sha256(repr((str(source.resolve()) if source else None, revision,
                                 str(definition.profile_path.resolve()))).encode()).hexdigest()[:16]
    def widget_key(name):
        return 'report_metadata_' + scope + '_' + name
    if definition.powertrain in {'electric', 'hybrid'}:
        mass_text = st.text_input('Battery pack mass (kg)', value='',
                                  key=widget_key('battery_mass'),
                                  help='Optional manual engineering input. Leave blank to show N/A.')
        try:
            mass = float(mass_text) if mass_text.strip() else None
        except ValueError:
            mass = mass_text  # ReportingInputs provides the actionable validation error.
        defaults.update(battery_mass_kg=mass,
                        battery_mass_status='user_provided' if mass is not None else 'not_provided',
                        battery_mass_provenance='Manual report input' if mass is not None else None)
    manual = {key: value for key, value in defaults.items() if value != getattr(detected.reporting_inputs, key)}
    st.caption('Agrochemical mass is calculated automatically using the selected reporting profile.')
    with st.expander('Advanced Configuration'):
        if not st.checkbox('Override automatically detected reporting information', key=widget_key('advanced')):
            return manual
        st.caption('Overrides affect reporting only. Chemical density has no assumed default.')
        identifiers = [None, *[item['configuration_id'] for item in profile.metadata.machine_configurations
                              if definition.powertrain in item['powertrains']]]
        selected = st.selectbox('Machine configuration override', identifiers,
                                key=widget_key('configuration'),
                                index=identifiers.index(defaults['machine_configuration_id']) if defaults['machine_configuration_id'] in identifiers else 0,
                                format_func=lambda value: value or 'Automatic / unspecified')
        if selected and selected != defaults['machine_configuration_id']:
            record = next(item for item in profile.metadata.machine_configurations if item['configuration_id'] == selected)
            defaults.update(record['reporting_inputs'])
            defaults['machine_configuration_id'] = selected
        density = st.checkbox('Specify agrochemical density [kg/L]', value=defaults['density_kg_per_L'] is not None, key=widget_key('density_enabled'))
        if density:
            density_text = st.text_input('Agrochemical density [kg/L] (strictly positive)', value=str(defaults['density_kg_per_L'] or ''), key=widget_key('density'))
            try:
                defaults['density_kg_per_L'] = float(density_text) if density_text.strip() else None
            except ValueError:
                defaults['density_kg_per_L'] = density_text  # validated by ReportingInputs before generation
            defaults['density_provenance'] = st.text_input('Density source / assumption', value=defaults['density_provenance'] or '', key=widget_key('density_provenance')) or None
        else:
            defaults.update(density_kg_per_L=None, density_provenance=None)
    return {key: value for key, value in defaults.items() if value != getattr(detected.reporting_inputs, key)}


def _render_profile_report_options(st, definition, inspection):
    """Configure the existing profile engine; persist only explicit input mappings."""
    profile = load_reporting_profile(definition.profile_path)
    options = ImportOptions(strict=False)
    with st.expander("Channels, statistics, plots and optional inputs"):
        if definition.powertrain == "diesel":
            names = [channel.source_name for channel in inspection.channels
                     if (normalize_display_unit(channel.unit) or "").lower() in {"°c", "°f", "k"}]
            configured = profile.raw_by_semantic_name()["engine_water_temperature"].source_name
            candidates = [None, *names]
            selected = st.selectbox("Engine Water Temperature source", candidates,
                                    index=candidates.index(configured) if configured in names else 0,
                                    format_func=lambda name: name or "Unavailable - select a measured coolant channel",
                                    help="Select the actual coolant measurement. Ambient temperature is never substituted.")
            if selected is not None and selected != configured:
                channel = next(c for c in inspection.channels if c.source_name == selected)
                if not channel.unit:
                    st.error("The selected temperature channel needs an explicit temperature unit in its source export.")
                else:
                    metadata = asdict(profile.metadata)
                    metadata["extends"] = str(definition.profile_path.resolve())
                    payload = {"version": 1, "profile": metadata, "overrides": {
                        "raw": {"engine_water_temperature": {"source_name": selected, "aliases": [], "unit": channel.unit}},
                        "plots": {"diesel_engine_water_temperature": {"primary_y_label": f"Water Temperature [{channel.unit}]"}},
                        "statistics": {s.statistic_id: {"unit": channel.unit} for s in profile.statistics if s.target == "engine_water_temperature"}}}
                    serialized = yaml.safe_dump(payload, sort_keys=False, allow_unicode=True)
                    folder = UI_WORKSPACE / "profiles"
                    folder.mkdir(parents=True, exist_ok=True)
                    path = folder / (hashlib.sha256(serialized.encode()).hexdigest() + ".yaml")
                    path.write_text(serialized, encoding="utf-8")
                    definition = replace(definition, profile_path=path)
                    profile = load_reporting_profile(path)
        # Streamlit round-trips formatted strings, so presentation names alone
        # cannot identify options: valid profiles may repeat names or titles.
        channel_labels = {c.semantic_name: f"{c.report_name} [{c.channel_type}] ({c.semantic_name})"
                          for c in (*profile.raw_channels, *profile.math_channels)}
        label = channel_labels.__getitem__
        math_names = [c.semantic_name for c in profile.math_channels]
        math = st.multiselect("Calculated MATH channels", math_names, default=math_names, format_func=label)
        channels = [c.semantic_name for c in profile.raw_channels if c.for_report] + math
        exported = st.multiselect("Channels exported to Excel", channels, default=channels, format_func=label)
        rms_names = [s.statistic_id for s in profile.statistics if s.operation in {"rms", "time_weighted_rms"}]
        other_names = [s.statistic_id for s in profile.statistics if s.statistic_id not in rms_names]
        rms = st.multiselect("RMS calculations", rms_names, default=rms_names)
        statistics = st.multiselect("Summary statistics", other_names, default=other_names)
        bottom = st.multiselect("Excel column statistics", ["max", "min", "first", "last"],
                                default=["max", "min", "first", "last"], format_func=str.upper)
        kpi_names = [k.kpi_id for k in profile.kpis]
        kpis = st.multiselect("Engineering KPIs", kpi_names, default=kpi_names)
        plot_names = [p.plot_id for p in profile.plots]
        plot_labels = {p.plot_id: f"{p.title} ({p.plot_id})" for p in profile.plots}
        plots = st.multiselect("Plots included in Excel", plot_names, default=plot_names,
                               format_func=plot_labels.__getitem__)
        plot_series = {}
        for plot in profile.plots:
            if plot.report_group == "torque_maps":
                series_labels = {s.semantic_name: s.label or s.component or s.semantic_name for s in plot.series}
                plot_series[plot.plot_id] = tuple(st.multiselect(
                    "Components - " + plot.title, list(series_labels), default=list(series_labels),
                    format_func=series_labels.__getitem__))
        if plot_series:
            st.caption("Torque Maps retain editable paired data independently of main-sheet exports. Legends identify motor, gearbox-output or driveshaft measurements; no conversion is assumed.")
        st.caption("Dependencies needed by selected plots, statistics and KPIs are calculated automatically. Plots whose source columns are not exported use the shared chart image in Excel.")
        ppt_plots = {}
        ppt_statistics = None
        if profile.presentation:
            metric_names = list(dict.fromkeys(s for slide in profile.presentation.slides for s in slide.statistics))
            ppt_statistics = st.multiselect("KPIs included in PowerPoint", metric_names, default=metric_names)
            for slide in profile.presentation.slides:
                if slide.plots:
                    ppt_plots[slide.slide_id] = tuple(st.multiselect(
                        f"PowerPoint plots - {slide.slide_id}", plot_names, default=list(slide.plots),
                        max_selections=len(slide.plots), format_func=plot_labels.__getitem__))
        if inspection.quality.sample_count > 1 and st.checkbox("Use a selected sample interval"):
            last = inspection.quality.sample_count - 1
            first_sample = int(st.number_input("First sample (zero based)", min_value=0, max_value=max(0, last - 1), value=0))
            last_sample = int(st.number_input("Last sample (inclusive)", min_value=first_sample + 1, max_value=last, value=last))
            start_row = inspection.quality.data_start_row
            options = replace(options, data_start_row=start_row + first_sample, data_end_row=start_row + last_sample)
    selection = ProfileReportSelection(tuple(exported), tuple(math), tuple(rms + statistics), tuple(kpis), tuple(plots),
                                       ppt_plots or None, tuple(ppt_statistics) if ppt_statistics is not None else None,
                                       tuple(bottom), plot_series=plot_series or None)
    return definition, selection, options


def _clear_profile_report_state(st: Any) -> None:
    """Invalidate validation and downloads when inputs change or validation restarts."""
    for key in ("profile_validation_key", "profile_validation_summary",
                "profile_report_result", "profile_report_result_key",
                "diesel_report_result", "diesel_report_result_key"):
        st.session_state.pop(key, None)


def _render_diesel_report_workflow(st: Any, profile_definition: Any, source_path: Path | None,
                                 validation_key: str | None, summary: Any, machine_name: str | None,
                                 *, selection=None, import_options=None, reporting_inputs=None) -> None:
    st.info("Diesel Excel and PowerPoint reporting are available.")
    can_generate = (source_path is not None and summary is not None and summary.is_valid
                    and st.session_state.get("profile_validation_key") == validation_key)
    if st.button("STEP 4 - Generate Engineering Report", type="primary", width="stretch", disabled=not can_generate):
        if not can_generate:
            return
        if not machine_name or not machine_name.strip():
            st.error("Enter a machine / vehicle name before generating reports.")
            return
        try:
            with st.spinner("Generating Diesel Excel and PowerPoint engineering reports..."):
                result = generate_reporting_profile_engineering_report(
                    source_path, profile_definition.profile_path,
                    UI_RUNS / datetime.now().strftime("%Y%m%d_%H%M%S_%f"),
                    import_options or ImportOptions(strict=True), machine_name_override=machine_name,
                    reporting_inputs=reporting_inputs,
                    **({"selection": selection} if selection is not None else {}),
                )
        except VSMPostProcessingError as exc:
            st.error(_friendly_profile_error(exc))
            with st.expander("Diagnostics"):
                st.write(str(exc))
            return
        st.session_state["diesel_report_result"] = result
        st.session_state["diesel_report_result_key"] = (validation_key, machine_name)
        st.success(f"Engineering reports generated: {result.report_path.name} and {result.presentation_path.name}")
        _render_profile_report_completion(st, result)
    result = st.session_state.get("diesel_report_result")
    if result is not None and st.session_state.get("diesel_report_result_key") == (validation_key, machine_name):
        st.subheader("STEP 5 - Access Output")
        _render_profile_report_download(st, result)


def _render_profile_validation_summary(st: Any, summary: Any) -> None:
    if summary.is_valid:
        st.success(f"{summary.profile_name} validation passed.")
    else:
        st.error(f"{summary.profile_name} validation did not pass.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Required raw channels", summary.required_raw_count)
    c2.metric("Resolved", summary.resolved_raw_count)
    c3.metric("Missing required", summary.missing_required_count)
    c4.metric("Optional channels unavailable", summary.missing_optional_count)
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("MATH channels", summary.math_count)
    c6.metric("Statistics", summary.statistic_count)
    c7.metric("KPIs", summary.kpi_count)
    c8.metric("Profile plots", summary.plot_count)

    if summary.all_zero_resolved_count:
        st.info(
            f"{summary.all_zero_resolved_count} resolved channel(s) are inactive/all-zero. "
            "This is informational and does not make the profile invalid."
        )
    _render_diagnostics(st, getattr(summary, "diagnostics", ()))
    if summary.missing_optional_names:
        st.info(
            "Optional source channels are not required for profile validation. "
            "Associated optional outputs are skipped when unavailable."
        )
        with st.expander("Optional channel details"):
            st.write("Unavailable optional channels:")
            for name in summary.missing_optional_names:
                st.write(f"- {name}")
    if summary.missing_required_names:
        st.warning("Missing required source/VSM channels: " + ", ".join(summary.missing_required_names[:20]))
    st.caption("MATH channels are calculated by the application and are not required in the source CSV.")
    with st.expander("Activity details"):
        st.write(f"Active resolved channels: {summary.active_resolved_count}")
        st.write(f"Constant resolved channels: {summary.constant_resolved_count}")
        st.write(f"All-zero resolved channels: {summary.all_zero_resolved_count}")
        if summary.all_zero_names:
            st.write(", ".join(summary.all_zero_names[:40]))


def _render_diagnostics(st: Any, messages) -> None:
    for message in messages:
        diagnostic = as_diagnostic(message)
        if diagnostic.user_visible:
            getattr(st, diagnostic.severity)(str(diagnostic))


def _render_profile_report_completion(st: Any, result: Any) -> None:
    _render_diagnostics(st, result.excel_result.statistics_result.diagnostics)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Machine", result.report_metadata.machine_name)
    c2.metric("Samples", f"{result.sample_count:,}")
    c3.metric("Report channels", result.report_channel_count)
    c4.metric("Plots", result.plot_count)
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("MATH", result.math_count)
    c6.metric("Statistics", result.statistic_count)
    c7.metric("KPIs", result.kpi_count)
    c8.metric("Output", result.report_path.name)


def _render_profile_report_download(st: Any, result: Any) -> None:
    report_path = result.report_path
    if report_path.exists():
        with report_path.open("rb") as handle:
            report_bytes = handle.read()
        st.download_button(
            "Download Excel Engineering Report",
            data=report_bytes,
            file_name=report_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    powerpoint_path = result.presentation_path
    if powerpoint_path.exists():
        with powerpoint_path.open("rb") as handle:
            powerpoint_bytes = handle.read()
        st.download_button(
            "Download PowerPoint Engineering Report",
            data=powerpoint_bytes,
            file_name=powerpoint_path.name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            width="stretch",
        )
    if result.excel_result.plotting_result.rendered_plots:
        with st.expander("Generated plot previews"):
            columns = st.columns(2)
            for index, plot in enumerate(result.excel_result.plotting_result.rendered_plots):
                columns[index % 2].image(plot.png_file, caption=plot.title, width="stretch")


def _friendly_input_error(exc: Exception) -> str:
    text = str(exc)
    if "Unsupported input type" in text or "Unsupported upload type" in text:
        return "Unsupported file format. Upload a CSV or XLSX VSM result file."
    if "No time channel" in text or "time channel" in text:
        return "No usable time channel was detected in the uploaded file."
    if "non-finite" in text or "invalid" in text.lower() or "numeric" in text.lower():
        return "The uploaded file contains invalid or nonnumeric data."
    return "The uploaded VSM file could not be inspected."


def _friendly_profile_error(exc: Exception) -> str:
    text = str(exc)
    if "missing" in text.lower() or "unavailable" in text.lower():
        return "The selected profile requires channels or dependencies that are not available in this file."
    if "ambiguous" in text.lower():
        return "The selected profile has ambiguous channel matches in this file."
    if "math" in text.lower():
        return "Profile math-channel validation failed."
    return "Profile-driven report generation failed."


def _format_duration_minutes(start: float | None, end: float | None, unit: str | None) -> str:
    value = None
    if start is not None and end is not None:
        duration = float(end) - float(start)
        normalized = (unit or "").strip().lower()
        if normalized in {"s", "sec", "second", "seconds"}:
            value = duration / 60
        elif normalized in {"min", "minute", "minutes"}:
            value = duration
    return f"{value:,.1f} min" if value is not None else "n/a"


def _persist_upload(uploaded: Any) -> Path:
    UI_WORKSPACE.mkdir(parents=True, exist_ok=True)
    data = uploaded.getvalue()
    digest = hashlib.sha256(data).hexdigest()[:12]
    safe_name = source_report_filename(uploaded.name, Path(uploaded.name).suffix)
    destination = UI_WORKSPACE / digest / safe_name
    destination.parent.mkdir(parents=True, exist_ok=True)
    if not destination.exists() or destination.stat().st_size != len(data):
        destination.write_bytes(data)
    return destination


def _valid_defaults(values: Any, allowed: list[str]) -> list[str]:
    if not isinstance(values, list):
        return []
    allowed_set = set(allowed)
    return [value for value in values if value in allowed_set]


def _stat_label(definition: dict[str, Any]) -> str:
    return f"{definition.get('display_name') or definition['statistic_id']} | {definition['operation']} | {definition['channel_id']}"


def _plot_is_available(definition: dict[str, Any], available_channel_ids: set[str]) -> bool:
    required = [definition["x_channel_id"], *[series["channel_id"] for series in definition["series"]]]
    return all(channel_id in available_channel_ids for channel_id in required)


def _fmt_number(value: Any) -> str:
    if value is None:
        return "n/a"
    try:
        return f"{float(value):g}"
    except (TypeError, ValueError):
        return str(value)


if __name__ == "__main__":
    main()
