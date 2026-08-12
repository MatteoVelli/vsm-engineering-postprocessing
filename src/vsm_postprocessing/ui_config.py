from __future__ import annotations

import ast
import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

import yaml

from .errors import ConfigurationError
from .duty_cycle import load_duty_cycle_config, validate_source_dataset
from .importer import ImportOptions, inspect_data_file, load_data_file


@dataclass(frozen=True)
class UiTemplateBundle:
    channel_selection: dict[str, Any]
    math_channels: dict[str, Any]
    statistics: dict[str, Any]
    plotting: dict[str, Any]
    excel_report: dict[str, Any]
    powerpoint_report: dict[str, Any]


@dataclass(frozen=True)
class UiRuntimeBundle:
    config_dir: Path
    pipeline_config: Path
    channel_selection_config: Path
    math_config: Path
    statistics_config: Path
    plotting_config: Path
    excel_statistics_config: Path
    excel_report_config: Path
    powerpoint_report_config: Path | None
    output_root: Path
    effective_math_channel_ids: tuple[str, ...]


@dataclass(frozen=True)
class FullDutyCycleScenarioDefinition:
    scenario_id: str
    display_name: str
    scenario_config: Path
    profile_provider_config: Path
    profile_workbook: Path
    math_config: Path
    statistics_config: Path
    plotting_config: Path
    excel_report_config: Path
    powerpoint_report_config: Path
    excel_download_filename: str
    powerpoint_download_filename: str


_PROFILE_VERSION = 1


def default_full_duty_cycle_scenario(project_root: str | Path) -> FullDutyCycleScenarioDefinition:
    root = Path(project_root).expanduser().resolve()
    config = root / "config"
    return FullDutyCycleScenarioDefinition(
        scenario_id="caiman_sp_hybrid_field_road_duty_cycle",
        display_name="Caiman SP Hybrid - 6 Field Cycles + Road Transfer",
        scenario_config=config / "duty_cycle_sergio_reference.yaml",
        profile_provider_config=root / "assets" / "scenarios" / "caiman_sp_hybrid" / "profile_provider.yaml",
        profile_workbook=root / "assets" / "scenarios" / "caiman_sp_hybrid" / "missing_phase_profiles.csv",
        math_config=config / "math_channels_example.yaml",
        statistics_config=config / "statistics_excel_report.yaml",
        plotting_config=config / "plotting_example.yaml",
        excel_report_config=config / "excel_report_duty_cycle.yaml",
        powerpoint_report_config=config / "powerpoint_report_duty_cycle.yaml",
        excel_download_filename="Caiman_SP_Hybrid_Engineering_Report.xlsx",
        powerpoint_download_filename="Caiman_SP_Hybrid_Engineering_Report.pptx",
    )


def build_full_duty_cycle_runtime_bundle(
    *,
    source_file: str | Path,
    profile_workbook: str | Path,
    profile_original_filename: str | None = None,
    runtime_dir: str | Path,
    scenario: FullDutyCycleScenarioDefinition,
    clean_before_run: bool = True,
) -> UiRuntimeBundle:
    source = Path(source_file).expanduser().resolve()
    profile = Path(profile_workbook).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise ConfigurationError(f"Full duty-cycle source file does not exist: {source}")
    if not profile.exists() or not profile.is_file():
        raise ConfigurationError(f"Full duty-cycle profile workbook does not exist: {profile}")

    channel_selection_config = scenario.scenario_config.parent / "channel_selection_example.yaml"
    for label, path in (
        ("scenario", scenario.scenario_config),
        ("Profile Provider", scenario.profile_provider_config),
        ("channel-selection", channel_selection_config),
        ("math", scenario.math_config),
        ("statistics", scenario.statistics_config),
        ("plotting", scenario.plotting_config),
        ("Excel report", scenario.excel_report_config),
        ("PowerPoint report", scenario.powerpoint_report_config),
    ):
        if not path.exists() or not path.is_file():
            raise ConfigurationError(f"Full duty-cycle {label} configuration does not exist: {path}")

    runtime = Path(runtime_dir).expanduser().resolve()
    config_dir = runtime / "runtime_config"
    output_root = runtime / "results"
    config_dir.mkdir(parents=True, exist_ok=True)

    pipeline_path = config_dir / "pipeline.yaml"
    pipeline = {
        "version": 1,
        "input": {"file": str(source), "strict": True},
        "duty_cycle": {
            "scenario": str(scenario.scenario_config.resolve()),
            "profile_provider": str(scenario.profile_provider_config.resolve()),
            "profile_workbook": str(profile),
            "profile_validation_mode": "compatible",
            "profile_original_filename": Path(profile_original_filename).name
            if profile_original_filename
            else profile.name,
        },
        "configs": {
            "channel_selection": str(channel_selection_config.resolve()),
            "math_channels": str(scenario.math_config.resolve()),
            "statistics": str(scenario.statistics_config.resolve()),
            "plotting": str(scenario.plotting_config.resolve()),
            "excel_statistics": str(scenario.statistics_config.resolve()),
            "excel_report": str(scenario.excel_report_config.resolve()),
            "powerpoint_report": str(scenario.powerpoint_report_config.resolve()),
        },
        "output": {"root_dir": str(output_root), "clean_before_run": bool(clean_before_run)},
    }
    _write_yaml(pipeline_path, pipeline)

    return UiRuntimeBundle(
        config_dir=config_dir,
        pipeline_config=pipeline_path,
        channel_selection_config=channel_selection_config.resolve(),
        math_config=scenario.math_config.resolve(),
        statistics_config=scenario.statistics_config.resolve(),
        plotting_config=scenario.plotting_config.resolve(),
        excel_statistics_config=scenario.statistics_config.resolve(),
        excel_report_config=scenario.excel_report_config.resolve(),
        powerpoint_report_config=scenario.powerpoint_report_config.resolve(),
        output_root=output_root,
        effective_math_channel_ids=(),
    )


def build_engineering_report_runtime_bundle(
    *,
    source_file: str | Path,
    runtime_dir: str | Path,
    project_root: str | Path,
    clean_before_run: bool = True,
) -> UiRuntimeBundle:
    """Build the primary single-file Engineering Report runtime configuration."""

    root = Path(project_root).expanduser().resolve()
    source = Path(source_file).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise ConfigurationError(f"Engineering Report source file does not exist: {source}")

    scenario = default_full_duty_cycle_scenario(root)
    channel_selection_config = scenario.scenario_config.parent / "channel_selection_example.yaml"
    for label, path in (
        ("scenario", scenario.scenario_config),
        ("profile-provider", scenario.profile_provider_config),
        ("profile asset", scenario.profile_workbook),
        ("channel-selection", channel_selection_config),
        ("math", scenario.math_config),
        ("statistics", scenario.statistics_config),
        ("plotting", scenario.plotting_config),
        ("Excel report", scenario.excel_report_config),
        ("PowerPoint report", scenario.powerpoint_report_config),
    ):
        if not path.exists() or not path.is_file():
            raise ConfigurationError(f"Engineering Report {label} configuration does not exist: {path}")

    inspection = inspect_data_file(source, ImportOptions(strict=True))
    source_channel_ids = {channel.channel_id for channel in inspection.channels}
    required = _engineering_report_required_channel_ids(
        channel_selection=_read_yaml_mapping(channel_selection_config),
        math_channels=_read_yaml_mapping(scenario.math_config),
        statistics=_read_yaml_mapping(scenario.statistics_config),
        plotting=_read_yaml_mapping(scenario.plotting_config),
        excel_report=_read_yaml_mapping(scenario.excel_report_config),
    )
    missing = sorted(required - source_channel_ids)
    if missing:
        preview = ", ".join(missing[:12])
        extra = "" if len(missing) <= 12 else f" and {len(missing) - 12} more"
        raise ConfigurationError(
            "Engineering Report profile requires channel(s) unavailable in the uploaded VSM file: "
            + preview
            + extra
        )

    scenario_config = load_duty_cycle_config(scenario.scenario_config)
    source_dataset = load_data_file(source, ImportOptions(strict=True))
    validate_source_dataset(scenario_config, source_dataset)

    runtime = Path(runtime_dir).expanduser().resolve()
    config_dir = runtime / "runtime_config"
    output_root = runtime / "results"
    config_dir.mkdir(parents=True, exist_ok=True)
    pipeline_path = config_dir / "pipeline.yaml"
    pipeline = {
        "version": 1,
        "input": {"file": str(source), "strict": True},
        "duty_cycle": {
            "scenario": str(scenario.scenario_config.resolve()),
            "profile_provider": str(scenario.profile_provider_config.resolve()),
            "profile_workbook": str(scenario.profile_workbook.resolve()),
            "profile_validation_mode": "strict",
            "profile_original_filename": scenario.profile_workbook.name,
        },
        "configs": {
            "channel_selection": str(channel_selection_config.resolve()),
            "math_channels": str(scenario.math_config.resolve()),
            "statistics": str(scenario.statistics_config.resolve()),
            "plotting": str(scenario.plotting_config.resolve()),
            "excel_statistics": str(scenario.statistics_config.resolve()),
            "excel_report": str(scenario.excel_report_config.resolve()),
            "powerpoint_report": str(scenario.powerpoint_report_config.resolve()),
        },
        "output": {"root_dir": str(output_root), "clean_before_run": bool(clean_before_run)},
    }
    _write_yaml(pipeline_path, pipeline)

    return UiRuntimeBundle(
        config_dir=config_dir,
        pipeline_config=pipeline_path,
        channel_selection_config=channel_selection_config.resolve(),
        math_config=scenario.math_config.resolve(),
        statistics_config=scenario.statistics_config.resolve(),
        plotting_config=scenario.plotting_config.resolve(),
        excel_statistics_config=scenario.statistics_config.resolve(),
        excel_report_config=scenario.excel_report_config.resolve(),
        powerpoint_report_config=scenario.powerpoint_report_config.resolve(),
        output_root=output_root,
        effective_math_channel_ids=(),
    )


def load_ui_templates(project_root: str | Path) -> UiTemplateBundle:
    root = Path(project_root).expanduser().resolve()
    config = root / "config"
    return UiTemplateBundle(
        channel_selection=_read_yaml_mapping(config / "channel_selection_example.yaml"),
        math_channels=_read_yaml_mapping(config / "math_channels_example.yaml"),
        statistics=_read_yaml_mapping(config / "statistics_excel_report.yaml"),
        plotting=_read_yaml_mapping(config / "plotting_example.yaml"),
        excel_report=_read_yaml_mapping(config / "excel_report_example.yaml"),
        powerpoint_report=_read_yaml_mapping(config / "powerpoint_report_example.yaml"),
    )



def available_math_channel_ids(
    math_template: Mapping[str, Any],
    source_channel_ids: Iterable[str],
) -> list[str]:
    """Return template math channels whose complete dependency graph exists in the source."""

    definitions = math_template.get("math_channels")
    constants = math_template.get("constants", {})
    if not isinstance(definitions, list) or not isinstance(constants, dict):
        raise ConfigurationError("Invalid math template for UI availability analysis")
    math_defs = {item["channel_id"]: item for item in definitions}
    available_names = set(source_channel_ids) | set(constants)
    available_math: set[str] = set()
    changed = True
    while changed:
        changed = False
        for channel_id, definition in math_defs.items():
            if channel_id in available_math:
                continue
            dependencies = set(_expression_names(str(definition["expression"])))
            if dependencies.issubset(available_names | available_math):
                available_math.add(channel_id)
                changed = True
    return [item["channel_id"] for item in definitions if item["channel_id"] in available_math]

def default_ui_profile(templates: UiTemplateBundle) -> dict[str, Any]:
    channel_selection = templates.channel_selection["selection"]["export_channels"]
    math_ids = [item["channel_id"] for item in templates.math_channels["math_channels"]]
    statistic_ids = [item["statistic_id"] for item in templates.statistics["statistics"]]
    plot_ids = [item["plot_id"] for item in templates.plotting["plots"]]
    report_channels = list(templates.excel_report["channels"])
    kpis = list(templates.excel_report.get("statistics", {}).get("kpis", []))
    return {
        "version": _PROFILE_VERSION,
        "export_channels": list(channel_selection),
        "math_channels": math_ids,
        "report_channels": report_channels,
        "statistics": statistic_ids,
        "kpis": kpis,
        "plots": plot_ids,
        "generate_powerpoint": True,
    }


def load_ui_profile(path: str | Path, *, fallback: Mapping[str, Any] | None = None) -> dict[str, Any]:
    profile_path = Path(path).expanduser().resolve()
    if not profile_path.exists():
        if fallback is None:
            raise ConfigurationError(f"UI profile does not exist: {profile_path}")
        return copy.deepcopy(dict(fallback))
    profile = _read_yaml_mapping(profile_path)
    _validate_profile(profile)
    return profile


def save_ui_profile(path: str | Path, profile: Mapping[str, Any]) -> Path:
    payload = copy.deepcopy(dict(profile))
    payload.setdefault("version", _PROFILE_VERSION)
    _validate_profile(payload)
    destination = Path(path).expanduser().resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    _write_yaml(destination, payload)
    return destination


def _engineering_report_required_channel_ids(
    *,
    channel_selection: Mapping[str, Any],
    math_channels: Mapping[str, Any],
    statistics: Mapping[str, Any],
    plotting: Mapping[str, Any],
    excel_report: Mapping[str, Any],
) -> set[str]:
    math_ids = {str(item["channel_id"]) for item in math_channels.get("math_channels", [])}
    constants = set(math_channels.get("constants", {}))
    required: set[str] = set()

    def add(channel_id: str) -> None:
        if channel_id not in math_ids and channel_id not in constants:
            required.add(channel_id)

    for channel_id in channel_selection.get("selection", {}).get("export_channels", []):
        add(str(channel_id))
    for channel_id in math_channels.get("selection", {}).get("export_source_channels", []):
        add(str(channel_id))
    for definition in math_channels.get("math_channels", []):
        for dependency in _expression_names(str(definition["expression"])):
            add(dependency)
        comparison = definition.get("compare_to")
        if isinstance(comparison, Mapping) and comparison.get("required", True):
            add(str(comparison["channel_id"]))
    for definition in statistics.get("statistics", []):
        add(str(definition["channel_id"]))
    for plot in plotting.get("plots", []):
        add(str(plot["x_channel_id"]))
        for series in plot.get("series", []):
            add(str(series["channel_id"]))
    for channel_id in excel_report.get("channels", []):
        add(str(channel_id))
    return required


def build_runtime_bundle(
    *,
    source_file: str | Path,
    runtime_dir: str | Path,
    templates: UiTemplateBundle,
    time_channel_id: str,
    export_channel_ids: Sequence[str],
    selected_math_channel_ids: Sequence[str],
    report_channel_ids: Sequence[str],
    selected_statistic_ids: Sequence[str],
    kpi_statistic_ids: Sequence[str],
    selected_plot_ids: Sequence[str],
    include_powerpoint: bool = True,
    clean_before_run: bool = True,
) -> UiRuntimeBundle:
    source = Path(source_file).expanduser().resolve()
    if not source.exists() or not source.is_file():
        raise ConfigurationError(f"UI source file does not exist: {source}")
    if not time_channel_id:
        raise ConfigurationError("A detected time channel is required for the UI pipeline")

    exports = _unique_nonempty(export_channel_ids, "export channels")
    report_channels = _unique_nonempty(report_channel_ids, "report channels")
    statistic_ids = _unique_nonempty(selected_statistic_ids, "statistics")
    plot_ids = _unique_nonempty(selected_plot_ids, "plots")
    kpi_ids = _unique(kpi_statistic_ids)

    runtime = Path(runtime_dir).expanduser().resolve()
    config_dir = runtime / "runtime_config"
    output_root = runtime / "results"
    config_dir.mkdir(parents=True, exist_ok=True)

    math_raw = copy.deepcopy(templates.math_channels)
    math_defs = {item["channel_id"]: item for item in math_raw["math_channels"]}
    stat_defs = {item["statistic_id"]: item for item in templates.statistics["statistics"]}
    plot_defs = {item["plot_id"]: item for item in templates.plotting["plots"]}

    _require_known(statistic_ids, stat_defs, "statistics")
    _require_known(kpi_ids, stat_defs, "KPI statistics")
    if not set(kpi_ids).issubset(statistic_ids):
        raise ConfigurationError("Every KPI statistic must also be selected in the statistics list")
    _require_known(plot_ids, plot_defs, "plots")
    _require_known(selected_math_channel_ids, math_defs, "math channels")

    required_math = set(selected_math_channel_ids)
    required_math.update(channel_id for channel_id in report_channels if channel_id in math_defs)
    for statistic_id in statistic_ids:
        channel_id = stat_defs[statistic_id]["channel_id"]
        if channel_id in math_defs:
            required_math.add(channel_id)
    for plot_id in plot_ids:
        plot = plot_defs[plot_id]
        x_id = plot["x_channel_id"]
        if x_id in math_defs:
            required_math.add(x_id)
        for series in plot["series"]:
            channel_id = series["channel_id"]
            if channel_id in math_defs:
                required_math.add(channel_id)

    effective_math = _math_dependency_closure(math_defs, required_math)
    if not effective_math:
        # The current deterministic pipeline has a mandatory math stage. Keep a
        # minimal infrastructure conversion so a "no custom math" UI choice is
        # still valid while leaving engineering outputs unchanged.
        fallback_id = "calc_time_minutes" if "calc_time_minutes" in math_defs else next(iter(math_defs), None)
        if fallback_id is None:
            raise ConfigurationError("The math template does not contain any math-channel definitions")
        effective_math = [fallback_id]

    channel_config = copy.deepcopy(templates.channel_selection)
    channel_config["selection"]["export_channels"] = [cid for cid in exports if cid != time_channel_id]
    if not channel_config["selection"]["export_channels"]:
        # Loader requires at least one explicit export channel in addition to auto time.
        raise ConfigurationError("Select at least one export channel in addition to the time channel")

    math_config = copy.deepcopy(templates.math_channels)
    math_config["math_channels"] = [copy.deepcopy(math_defs[channel_id]) for channel_id in effective_math]
    source_dependencies = _source_dependencies_for_math(math_defs, effective_math, math_config.get("constants", {}))
    math_exports = _unique([
        *[cid for cid in exports if cid != time_channel_id and cid not in math_defs],
        *source_dependencies,
    ])
    if not math_exports:
        math_exports = [time_channel_id]
    math_config["selection"]["export_source_channels"] = math_exports

    statistics_config = copy.deepcopy(templates.statistics)
    statistics_config["statistics"] = [copy.deepcopy(stat_defs[statistic_id]) for statistic_id in statistic_ids]

    plotting_config = copy.deepcopy(templates.plotting)
    plotting_config["plots"] = [copy.deepcopy(plot_defs[plot_id]) for plot_id in plot_ids]

    excel_config = copy.deepcopy(templates.excel_report)
    excel_channels = _unique([time_channel_id, *report_channels])
    excel_config["channels"] = excel_channels
    selected_stat_defs = {sid: stat_defs[sid] for sid in statistic_ids}
    top_rms: list[str] = []
    bottom_summary: list[str] = []
    excel_channel_set = set(excel_channels)
    for statistic_id, definition in selected_stat_defs.items():
        operation = definition["operation"]
        channel_id = definition["channel_id"]
        if operation in {"rms", "time_weighted_rms"}:
            if channel_id in excel_channel_set:
                top_rms.append(statistic_id)
        elif channel_id in excel_channel_set:
            bottom_summary.append(statistic_id)
    excel_config["statistics"]["top_rms"] = top_rms
    excel_config["statistics"]["kpis"] = list(kpi_ids)
    excel_config["statistics"]["bottom_summary"] = bottom_summary
    excel_config["plots"]["include"] = list(plot_ids)

    powerpoint_config: dict[str, Any] | None = None
    if include_powerpoint:
        powerpoint_config = copy.deepcopy(templates.powerpoint_report)
        filtered_slides: list[dict[str, Any]] = []
        selected_statistic_set = set(statistic_ids)
        selected_plot_set = set(plot_ids)
        for slide in powerpoint_config.get("slides", []):
            slide_copy = copy.deepcopy(slide)
            slide_copy["statistics"] = [
                sid for sid in slide_copy.get("statistics", []) if sid in selected_statistic_set
            ]
            slide_copy["plots"] = [
                pid for pid in slide_copy.get("plots", []) if pid in selected_plot_set
            ]
            if slide_copy.get("type") == "plot_pair" and not slide_copy["plots"]:
                continue
            filtered_slides.append(slide_copy)
        if not filtered_slides:
            raise ConfigurationError("The current UI selection does not leave any PowerPoint slides to generate")
        powerpoint_config["slides"] = filtered_slides

    channel_path = config_dir / "channel_selection.yaml"
    math_path = config_dir / "math_channels.yaml"
    statistics_path = config_dir / "statistics.yaml"
    excel_statistics_path = config_dir / "excel_statistics.yaml"
    plotting_path = config_dir / "plotting.yaml"
    excel_path = config_dir / "excel_report.yaml"
    powerpoint_path = config_dir / "powerpoint_report.yaml"
    pipeline_path = config_dir / "pipeline.yaml"

    _write_yaml(channel_path, channel_config)
    _write_yaml(math_path, math_config)
    _write_yaml(statistics_path, statistics_config)
    _write_yaml(excel_statistics_path, statistics_config)
    _write_yaml(plotting_path, plotting_config)
    _write_yaml(excel_path, excel_config)
    if powerpoint_config is not None:
        _write_yaml(powerpoint_path, powerpoint_config)

    pipeline = {
        "version": 1,
        "input": {"file": str(source), "strict": True},
        "configs": {
            "channel_selection": str(channel_path),
            "math_channels": str(math_path),
            "statistics": str(statistics_path),
            "plotting": str(plotting_path),
            "excel_statistics": str(excel_statistics_path),
            "excel_report": str(excel_path),
        },
        "output": {"root_dir": str(output_root), "clean_before_run": bool(clean_before_run)},
    }
    if include_powerpoint:
        pipeline["configs"]["powerpoint_report"] = str(powerpoint_path)
    _write_yaml(pipeline_path, pipeline)

    return UiRuntimeBundle(
        config_dir=config_dir,
        pipeline_config=pipeline_path,
        channel_selection_config=channel_path,
        math_config=math_path,
        statistics_config=statistics_path,
        plotting_config=plotting_path,
        excel_statistics_config=excel_statistics_path,
        excel_report_config=excel_path,
        powerpoint_report_config=powerpoint_path if include_powerpoint else None,
        output_root=output_root,
        effective_math_channel_ids=tuple(effective_math),
    )


def _math_dependency_closure(math_defs: Mapping[str, Mapping[str, Any]], requested: Iterable[str]) -> list[str]:
    requested_set = set(requested)
    _require_known(requested_set, math_defs, "math channels")
    order = list(math_defs)
    closure: set[str] = set()

    def add(channel_id: str) -> None:
        if channel_id in closure:
            return
        for name in _expression_names(str(math_defs[channel_id]["expression"])):
            if name in math_defs:
                add(name)
        closure.add(channel_id)

    for channel_id in order:
        if channel_id in requested_set:
            add(channel_id)
    return [channel_id for channel_id in order if channel_id in closure]


def _source_dependencies_for_math(
    math_defs: Mapping[str, Mapping[str, Any]],
    effective_math: Sequence[str],
    constants: Mapping[str, Any],
) -> list[str]:
    math_ids = set(math_defs)
    constant_ids = set(constants)
    dependencies: list[str] = []
    for channel_id in effective_math:
        for name in _expression_names(str(math_defs[channel_id]["expression"])):
            if name not in math_ids and name not in constant_ids and name not in dependencies:
                dependencies.append(name)
    return dependencies


def _expression_names(expression: str) -> list[str]:
    try:
        tree = ast.parse(expression, mode="eval")
    except SyntaxError as exc:
        raise ConfigurationError(f"Invalid math expression in UI template: {expression}") from exc
    function_names = {
        node.func.id
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
    }
    names: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and node.id not in function_names and node.id not in names:
            names.append(node.id)
    return names


def _validate_profile(profile: Mapping[str, Any]) -> None:
    allowed = {"version", "export_channels", "math_channels", "report_channels", "statistics", "kpis", "plots", "generate_powerpoint"}
    unknown = sorted(set(profile) - allowed)
    if unknown:
        raise ConfigurationError("Unknown UI profile key(s): " + ", ".join(unknown))
    if profile.get("version") != _PROFILE_VERSION:
        raise ConfigurationError(f"UI profile version must be {_PROFILE_VERSION}")
    if not isinstance(profile.get("generate_powerpoint", True), bool):
        raise ConfigurationError("UI profile 'generate_powerpoint' must be true or false")
    for key in ("export_channels", "math_channels", "report_channels", "statistics", "kpis", "plots"):
        value = profile.get(key)
        if not isinstance(value, list) or any(not isinstance(item, str) or not item for item in value):
            raise ConfigurationError(f"UI profile '{key}' must be a list of non-empty strings")
        if len(set(value)) != len(value):
            raise ConfigurationError(f"UI profile '{key}' contains duplicates")


def _read_yaml_mapping(path: Path) -> dict[str, Any]:
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, yaml.YAMLError) as exc:
        raise ConfigurationError(f"Could not read UI template '{path}': {exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigurationError(f"UI template must contain a YAML mapping: {path}")
    return raw


def _write_yaml(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(dict(payload), sort_keys=False, allow_unicode=True), encoding="utf-8")


def _unique(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        if value not in result:
            result.append(value)
    return result


def _unique_nonempty(values: Sequence[str], label: str) -> list[str]:
    result = _unique(value for value in values if isinstance(value, str) and value)
    if not result:
        raise ConfigurationError(f"Select at least one {label}")
    return result


def _require_known(values: Iterable[str], known: Mapping[str, Any], label: str) -> None:
    missing = sorted(set(values) - set(known))
    if missing:
        raise ConfigurationError(f"Unknown {label}: " + ", ".join(missing))
