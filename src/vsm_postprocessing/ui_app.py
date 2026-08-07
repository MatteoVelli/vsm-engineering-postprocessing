from __future__ import annotations

import hashlib
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from vsm_postprocessing.errors import VSMPostProcessingError
from vsm_postprocessing.importer import ImportOptions, inspect_data_file
from vsm_postprocessing.pipeline_engine import run_pipeline
from vsm_postprocessing.ui_config import (
    available_math_channel_ids,
    build_runtime_bundle,
    default_ui_profile,
    load_ui_profile,
    load_ui_templates,
    save_ui_profile,
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
PROFILE_PATH = PROJECT_ROOT / "config" / "ui_saved_profile.yaml"
UI_WORKSPACE = PROJECT_ROOT / "outputs" / "ui_workspace"
UI_RUNS = PROJECT_ROOT / "outputs" / "ui_runs"


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="VSM Post-Processing", page_icon="📊", layout="wide")
    st.title("VSM Engineering Post-Processing")
    st.caption("v1.1.0 | Deterministic processing: source → channels → math → statistics → plots → Excel + optional PowerPoint reports")

    templates = load_ui_templates(PROJECT_ROOT)
    defaults = default_ui_profile(templates)
    profile = load_ui_profile(PROFILE_PATH, fallback=defaults)

    uploaded = st.file_uploader("1. Load VSM results", type=["xlsx", "xlsm", "csv"])
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
                use_container_width=True,
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
        if st.button("Save selections", use_container_width=True):
            try:
                save_ui_profile(PROFILE_PATH, current_profile)
            except VSMPostProcessingError as exc:
                st.error(str(exc))
            else:
                st.success(f"Saved: {PROFILE_PATH.relative_to(PROJECT_ROOT)}")
    with right:
        run_clicked = st.button("Run complete pipeline", type="primary", use_container_width=True)

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
            use_container_width=True,
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
                use_container_width=True,
            )
            if os.name == "nt" and col_b.button("Open report in Excel", use_container_width=True):
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
                use_container_width=True,
            )
            if os.name == "nt" and col_d.button("Open report in PowerPoint", use_container_width=True):
                os.startfile(powerpoint_path)  # type: ignore[attr-defined]


def _persist_upload(uploaded: Any) -> Path:
    UI_WORKSPACE.mkdir(parents=True, exist_ok=True)
    data = uploaded.getvalue()
    digest = hashlib.sha256(data).hexdigest()[:12]
    safe_name = Path(uploaded.name).name
    destination = UI_WORKSPACE / f"{digest}_{safe_name}"
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
