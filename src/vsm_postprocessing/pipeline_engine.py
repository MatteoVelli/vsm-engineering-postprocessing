from __future__ import annotations

import json
import platform
import shutil
import sys
import time
from datetime import datetime, timezone
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from .channel_manager import export_channel_selection, select_channels
from .errors import ConfigurationError, PipelineError, VSMPostProcessingError
from .excel_report_engine import ExcelReportResult, generate_excel_report
from .importer import ImportOptions, export_inspection, inspect_data_file
from .math_engine import export_math_channels, calculate_math_channels
from .plotting_engine import PlottingResult, render_plots
from .powerpoint_report_engine import PowerPointReportResult, build_powerpoint_report
from .statistics_engine import export_statistics, calculate_statistics
from .utils import atomic_write_text, sha256_file
from .version import __version__


@dataclass(frozen=True)
class PipelineConfig:
    version: int
    input_file: Path
    import_options: ImportOptions
    channel_selection_config: Path
    math_config: Path
    statistics_config: Path
    plotting_config: Path
    excel_statistics_config: Path
    excel_report_config: Path
    powerpoint_report_config: Path | None
    output_root: Path
    clean_before_run: bool


@dataclass
class PipelineStage:
    name: str
    status: str
    output_dir: Path
    outputs: dict[str, Path] = field(default_factory=dict)
    metrics: dict[str, Any] = field(default_factory=dict)
    error: str | None = None
    started_at_utc: str | None = None
    completed_at_utc: str | None = None
    duration_seconds: float | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "status": self.status,
            "output_dir": str(self.output_dir),
            "outputs": {key: str(value) for key, value in self.outputs.items()},
            "metrics": self.metrics,
            "error": self.error,
            "started_at_utc": self.started_at_utc,
            "completed_at_utc": self.completed_at_utc,
            "duration_seconds": self.duration_seconds,
        }


@dataclass
class PipelineResult:
    config_path: Path
    config: PipelineConfig
    stages: list[PipelineStage]
    status: str
    report_path: Path | None
    powerpoint_path: Path | None
    manifest_path: Path
    summary_path: Path
    log_path: Path
    started_at_utc: str
    completed_at_utc: str
    duration_seconds: float
    processing_input_file: Path

    @property
    def completed_stage_count(self) -> int:
        return sum(stage.status == "PASS" for stage in self.stages)


_ALLOWED_ROOT_KEYS = {"version", "input", "configs", "output"}
_ALLOWED_INPUT_KEYS = {
    "file",
    "sheet_name",
    "header_row",
    "unit_row",
    "data_start_row",
    "data_end_row",
    "last_channel_column",
    "time_channel",
    "strict",
}
_ALLOWED_CONFIG_KEYS = {
    "channel_selection",
    "math_channels",
    "statistics",
    "plotting",
    "excel_statistics",
    "excel_report",
    "powerpoint_report",
}
_ALLOWED_OUTPUT_KEYS = {"root_dir", "clean_before_run"}
def load_pipeline_config(path: str | Path) -> PipelineConfig:
    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Pipeline configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Pipeline configuration is not a file: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Pipeline configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in pipeline configuration '{config_path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Pipeline configuration root must be a YAML mapping")
    _reject_unknown_keys(raw, _ALLOWED_ROOT_KEYS, "root")
    if raw.get("version") != 1:
        raise ConfigurationError("Pipeline configuration version must be 1")

    input_raw = raw.get("input")
    if not isinstance(input_raw, dict):
        raise ConfigurationError("input must be a YAML mapping")
    _reject_unknown_keys(input_raw, _ALLOWED_INPUT_KEYS, "input")
    input_file = _resolve_required_path(config_path, input_raw.get("file"), "input.file")

    strict = input_raw.get("strict", True)
    if not isinstance(strict, bool):
        raise ConfigurationError("input.strict must be true or false")
    import_options = ImportOptions(
        sheet_name=_optional_string(input_raw.get("sheet_name"), "input.sheet_name"),
        header_row=_optional_positive_int(input_raw.get("header_row"), "input.header_row"),
        unit_row=_optional_positive_int(input_raw.get("unit_row"), "input.unit_row"),
        data_start_row=_optional_positive_int(input_raw.get("data_start_row"), "input.data_start_row"),
        data_end_row=_optional_positive_int(input_raw.get("data_end_row"), "input.data_end_row"),
        last_channel_column=_optional_positive_int(
            input_raw.get("last_channel_column"), "input.last_channel_column"
        ),
        time_channel=_optional_string(input_raw.get("time_channel"), "input.time_channel"),
        strict=strict,
    )

    configs_raw = raw.get("configs")
    if not isinstance(configs_raw, dict):
        raise ConfigurationError("configs must be a YAML mapping")
    _reject_unknown_keys(configs_raw, _ALLOWED_CONFIG_KEYS, "configs")
    required_keys = _ALLOWED_CONFIG_KEYS - {"powerpoint_report"}
    required_configs = {
        key: _resolve_required_path(config_path, configs_raw.get(key), f"configs.{key}")
        for key in sorted(required_keys)
    }
    powerpoint_value = configs_raw.get("powerpoint_report")
    powerpoint_config = (
        _resolve_required_path(config_path, powerpoint_value, "configs.powerpoint_report")
        if powerpoint_value is not None
        else None
    )

    output_raw = raw.get("output", {})
    if not isinstance(output_raw, dict):
        raise ConfigurationError("output must be a YAML mapping")
    _reject_unknown_keys(output_raw, _ALLOWED_OUTPUT_KEYS, "output")
    output_root = _resolve_path_value(
        config_path,
        output_raw.get("root_dir", "../outputs/end_to_end"),
        "output.root_dir",
    )
    clean_before_run = output_raw.get("clean_before_run", True)
    if not isinstance(clean_before_run, bool):
        raise ConfigurationError("output.clean_before_run must be true or false")

    if output_root == config_path.parent or output_root == input_file.parent:
        raise ConfigurationError("output.root_dir must be a dedicated output directory")
    try:
        input_file.relative_to(output_root)
    except ValueError:
        pass
    else:
        raise ConfigurationError("input.file must not be located inside output.root_dir")

    return PipelineConfig(
        version=1,
        input_file=input_file,
        import_options=import_options,
        channel_selection_config=required_configs["channel_selection"],
        math_config=required_configs["math_channels"],
        statistics_config=required_configs["statistics"],
        plotting_config=required_configs["plotting"],
        excel_statistics_config=required_configs["excel_statistics"],
        excel_report_config=required_configs["excel_report"],
        powerpoint_report_config=powerpoint_config,
        output_root=output_root,
        clean_before_run=clean_before_run,
    )


def run_pipeline(config_file: str | Path) -> PipelineResult:
    config_path = Path(config_file).expanduser().resolve()
    config = load_pipeline_config(config_path)
    _prepare_output_root(config.output_root, config.clean_before_run)

    manifest_path = config.output_root / "pipeline_manifest.json"
    summary_path = config.output_root / "pipeline_summary.txt"
    log_path = config.output_root / "pipeline.log"
    run_started_at = _utc_now()
    run_started_perf = time.perf_counter()
    log_lines: list[str] = []
    _log(log_lines, "INFO", f"Pipeline started | software={_software_version()} | source={config.input_file}")

    stages: list[PipelineStage] = []
    report_path: Path | None = None
    powerpoint_path: Path | None = None
    current_stage_name = "initialization"
    current_stage_started_at: str | None = None
    current_stage_started_perf: float | None = None

    stage_specs = [
        ("inspection", "01_inspection"),
        ("channel_selection", "02_channel_selection"),
        ("math_channels", "03_math_channels"),
        ("statistics", "04_statistics"),
        ("plotting", "05_plots"),
        ("excel_report", "06_excel_report"),
    ]
    if config.powerpoint_report_config is not None:
        stage_specs.append(("powerpoint_report", "07_powerpoint_report"))
    stage_dirs = {name: config.output_root / dirname for name, dirname in stage_specs}

    def begin_stage(name: str) -> tuple[str, float]:
        nonlocal current_stage_name, current_stage_started_at, current_stage_started_perf
        current_stage_name = name
        current_stage_started_at = _utc_now()
        current_stage_started_perf = time.perf_counter()
        _log(log_lines, "INFO", f"Stage started: {name}")
        return current_stage_started_at, current_stage_started_perf

    def pass_stage(
        name: str,
        output_dir: Path,
        outputs: dict[str, Path],
        metrics: dict[str, Any],
        started_at: str,
        started_perf: float,
    ) -> None:
        completed_at = _utc_now()
        duration = time.perf_counter() - started_perf
        stages.append(
            PipelineStage(
                name=name,
                status="PASS",
                output_dir=output_dir,
                outputs=outputs,
                metrics=metrics,
                started_at_utc=started_at,
                completed_at_utc=completed_at,
                duration_seconds=round(duration, 6),
            )
        )
        _log(log_lines, "INFO", f"Stage passed: {name} | duration={duration:.3f}s")

    processing_input_file = config.input_file
    processing_import_options = config.import_options

    try:
        started_at, started_perf = begin_stage("inspection")
        inspection = inspect_data_file(config.input_file, config.import_options)
        inspection_outputs = export_inspection(inspection, stage_dirs["inspection"])
        pass_stage(
            "inspection",
            stage_dirs["inspection"],
            inspection_outputs,
            {
                "samples": inspection.quality.sample_count,
                "channels": inspection.quality.channel_count,
                "raw_channels": inspection.quality.raw_channel_count,
                "math_channels": inspection.quality.math_channel_count,
            },
            started_at,
            started_perf,
        )

        started_at, started_perf = begin_stage("channel_selection")
        selection = select_channels(
            processing_input_file,
            config.channel_selection_config,
            processing_import_options,
        )
        selection_outputs = export_channel_selection(selection, stage_dirs["channel_selection"])
        pass_stage(
            "channel_selection",
            stage_dirs["channel_selection"],
            selection_outputs,
            {"samples": selection.sample_count, "selected_channels": selection.channel_count},
            started_at,
            started_perf,
        )

        started_at, started_perf = begin_stage("math_channels")
        math_result = calculate_math_channels(
            processing_input_file,
            config.math_config,
            processing_import_options,
        )
        math_outputs = export_math_channels(math_result, stage_dirs["math_channels"])
        pass_stage(
            "math_channels",
            stage_dirs["math_channels"],
            math_outputs,
            {
                "samples": math_result.sample_count,
                "source_channels": math_result.source_channel_count,
                "math_channels": math_result.math_channel_count,
                "comparisons": len(math_result.comparisons),
            },
            started_at,
            started_perf,
        )

        started_at, started_perf = begin_stage("statistics")
        statistics = calculate_statistics(
            processing_input_file,
            config.statistics_config,
            processing_import_options,
            math_config_file=config.math_config,
        )
        statistics_outputs = export_statistics(statistics, stage_dirs["statistics"])
        pass_stage(
            "statistics",
            stage_dirs["statistics"],
            statistics_outputs,
            {
                "samples": statistics.sample_count,
                "statistics": statistics.statistic_count,
                "comparisons": statistics.comparison_count,
            },
            started_at,
            started_perf,
        )

        started_at, started_perf = begin_stage("plotting")
        plotting = render_plots(
            processing_input_file,
            config.plotting_config,
            stage_dirs["plotting"],
            processing_import_options,
            math_config_file=config.math_config,
        )
        plot_outputs = _plot_outputs(plotting, stage_dirs["plotting"])
        pass_stage(
            "plotting",
            stage_dirs["plotting"],
            plot_outputs,
            {
                "samples": plotting.sample_count,
                "plots": plotting.plot_count,
                "series": plotting.series_count,
                "secondary_axis_plots": plotting.secondary_axis_plot_count,
                "svg_plots": plotting.svg_count,
            },
            started_at,
            started_perf,
        )

        started_at, started_perf = begin_stage("excel_report")
        excel = generate_excel_report(
            processing_input_file,
            config.excel_report_config,
            config.excel_statistics_config,
            config.plotting_config,
            stage_dirs["excel_report"],
            processing_import_options,
            math_config_file=config.math_config,
            precomputed_statistics_result=(
                statistics
                if config.statistics_config.resolve() == config.excel_statistics_config.resolve()
                else None
            ),
            precomputed_plotting_result=plotting,
        )
        report_path = excel.report_path
        excel_outputs = _excel_outputs(excel)
        pass_stage(
            "excel_report",
            stage_dirs["excel_report"],
            excel_outputs,
            {
                "samples": excel.sample_count,
                "report_channels": excel.channel_count,
                "statistics": excel.statistic_count,
                "configured_plots": excel.configured_plot_count,
                "plot_series": excel.plotting_result.series_count,
                "native_excel_charts": excel.native_excel_chart_count,
                "plot_images": excel.embedded_plot_image_count,
            },
            started_at,
            started_perf,
        )

        if config.powerpoint_report_config is not None:
            started_at, started_perf = begin_stage("powerpoint_report")
            powerpoint = build_powerpoint_report(
                excel.statistics_result,
                excel.plotting_result,
                config.powerpoint_report_config,
                stage_dirs["powerpoint_report"],
                plot_assets_dir=excel.plot_assets_dir,
            )
            powerpoint_path = powerpoint.presentation_path
            powerpoint_outputs = _powerpoint_outputs(powerpoint)
            pass_stage(
                "powerpoint_report",
                stage_dirs["powerpoint_report"],
                powerpoint_outputs,
                {
                    "samples": powerpoint.sample_count,
                    "slides": powerpoint.slide_count,
                    "statistics": powerpoint.statistic_count,
                    "plots": powerpoint.plot_count,
                },
                started_at,
                started_perf,
            )
    except VSMPostProcessingError as exc:
        _append_failed_stage(
            stages,
            current_stage_name,
            stage_dirs.get(current_stage_name, config.output_root),
            str(exc),
            current_stage_started_at,
            current_stage_started_perf,
        )
        _log(log_lines, "ERROR", f"Stage failed: {current_stage_name} | {type(exc).__name__}: {exc}")
        result = _finalize_pipeline_result(
            config_path, config, stages, "FAIL", report_path, powerpoint_path, processing_input_file,
            manifest_path, summary_path, log_path, run_started_at, run_started_perf, log_lines,
        )
        _write_pipeline_metadata(result)
        raise PipelineError(f"End-to-end pipeline failed during stage '{current_stage_name}': {exc}") from exc
    except Exception as exc:
        # Preserve diagnostics even for unexpected implementation/runtime errors.
        message = f"Unexpected {type(exc).__name__}: {exc}"
        _append_failed_stage(
            stages,
            current_stage_name,
            stage_dirs.get(current_stage_name, config.output_root),
            message,
            current_stage_started_at,
            current_stage_started_perf,
        )
        _log(log_lines, "ERROR", f"Stage failed unexpectedly: {current_stage_name} | {message}")
        result = _finalize_pipeline_result(
            config_path, config, stages, "FAIL", report_path, powerpoint_path, processing_input_file,
            manifest_path, summary_path, log_path, run_started_at, run_started_perf, log_lines,
        )
        _write_pipeline_metadata(result)
        raise PipelineError(f"End-to-end pipeline failed during stage '{current_stage_name}': {message}") from exc

    _log(log_lines, "INFO", "Pipeline completed successfully")
    result = _finalize_pipeline_result(
        config_path, config, stages, "PASS", report_path, powerpoint_path, processing_input_file,
        manifest_path, summary_path, log_path, run_started_at, run_started_perf, log_lines,
    )
    _write_pipeline_metadata(result)
    return result


def _append_failed_stage(
    stages: list[PipelineStage],
    name: str,
    output_dir: Path,
    error: str,
    started_at: str | None,
    started_perf: float | None,
) -> None:
    completed_at = _utc_now()
    duration = (time.perf_counter() - started_perf) if started_perf is not None else None
    stages.append(
        PipelineStage(
            name=name,
            status="FAIL",
            output_dir=output_dir,
            error=error,
            started_at_utc=started_at,
            completed_at_utc=completed_at,
            duration_seconds=round(duration, 6) if duration is not None else None,
        )
    )


def _finalize_pipeline_result(
    config_path: Path,
    config: PipelineConfig,
    stages: list[PipelineStage],
    status: str,
    report_path: Path | None,
    powerpoint_path: Path | None,
    processing_input_file: Path,
    manifest_path: Path,
    summary_path: Path,
    log_path: Path,
    started_at_utc: str,
    started_perf: float,
    log_lines: list[str],
) -> PipelineResult:
    completed_at = _utc_now()
    duration = time.perf_counter() - started_perf
    atomic_write_text(log_path, "\n".join(log_lines) + "\n")
    return PipelineResult(
        config_path=config_path,
        config=config,
        stages=stages,
        status=status,
        report_path=report_path,
        powerpoint_path=powerpoint_path,
        manifest_path=manifest_path,
        summary_path=summary_path,
        log_path=log_path,
        started_at_utc=started_at_utc,
        completed_at_utc=completed_at,
        duration_seconds=round(duration, 6),
        processing_input_file=processing_input_file,
    )


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _log(lines: list[str], level: str, message: str) -> None:
    lines.append(f"{_utc_now()} | {level:<5} | {message}")


def _software_version() -> str:
    return __version__

def _write_pipeline_metadata(result: PipelineResult) -> None:
    config = result.config
    configs = {
        "channel_selection": config.channel_selection_config,
        "math_channels": config.math_config,
        "statistics": config.statistics_config,
        "plotting": config.plotting_config,
        "excel_statistics": config.excel_statistics_config,
        "excel_report": config.excel_report_config,
    }
    if config.powerpoint_report_config is not None:
        configs["powerpoint_report"] = config.powerpoint_report_config
    manifest = {
        "status": result.status,
        "software": {
            "name": "vsm-postprocessing",
            "version": _software_version(),
            "python": platform.python_version(),
            "python_executable": sys.executable,
            "platform": platform.platform(),
        },
        "started_at_utc": result.started_at_utc,
        "completed_at_utc": result.completed_at_utc,
        "duration_seconds": result.duration_seconds,
        "pipeline_configuration_version": config.version,
        "pipeline_configuration_file": str(result.config_path),
        "pipeline_configuration_sha256": sha256_file(result.config_path),
        "source_file": str(config.input_file),
        "source_sha256": sha256_file(config.input_file) if config.input_file.exists() else None,
        "processing_input_file": str(result.processing_input_file),
        "processing_input_sha256": (
            sha256_file(result.processing_input_file) if result.processing_input_file.exists() else None
        ),
        "output_root": str(config.output_root),
        "log_file": str(result.log_path),
        "import_options": {
            "sheet_name": config.import_options.sheet_name,
            "header_row": config.import_options.header_row,
            "unit_row": config.import_options.unit_row,
            "data_start_row": config.import_options.data_start_row,
            "data_end_row": config.import_options.data_end_row,
            "last_channel_column": config.import_options.last_channel_column,
            "time_channel": config.import_options.time_channel,
            "strict": config.import_options.strict,
        },
        "configs": {
            name: {"path": str(path), "sha256": sha256_file(path)}
            for name, path in configs.items()
        },
        "completed_stage_count": result.completed_stage_count,
        "stage_count": len(result.stages),
        "stages": [stage.to_dict() for stage in result.stages],
        "final_report": str(result.report_path) if result.report_path else None,
        "final_powerpoint": str(result.powerpoint_path) if result.powerpoint_path else None,
    }
    atomic_write_text(result.manifest_path, json.dumps(manifest, indent=2) + "\n")

    lines = [
        "VSM END-TO-END PIPELINE",
        "=======================",
        f"Status: {result.status}",
        f"Software: vsm-postprocessing {_software_version()}",
        f"Started (UTC): {result.started_at_utc}",
        f"Completed (UTC): {result.completed_at_utc}",
        f"Duration: {result.duration_seconds:.3f} s",
        f"Source: {config.input_file}",
        f"Processing input: {result.processing_input_file}",
        f"Output root: {config.output_root}",
        f"Stages passed: {result.completed_stage_count}/{len(result.stages)}",
        f"Log: {result.log_path}",
        "",
        "Stages:",
    ]
    for index, stage in enumerate(result.stages, start=1):
        metrics = ", ".join(f"{key}={value}" for key, value in stage.metrics.items())
        suffix = f" | {metrics}" if metrics else ""
        if stage.duration_seconds is not None:
            suffix += f" | duration={stage.duration_seconds:.3f}s"
        if stage.error:
            suffix += f" | error={stage.error}"
        lines.append(f"{index:02d}. {stage.name}: {stage.status}{suffix}")
    if result.report_path:
        lines.extend(["", f"Final Excel report: {result.report_path}"])
    if result.powerpoint_path:
        lines.append(f"Final PowerPoint report: {result.powerpoint_path}")
    atomic_write_text(result.summary_path, "\n".join(lines) + "\n")

def _plot_outputs(result: PlottingResult, output_dir: Path) -> dict[str, Path]:
    outputs: dict[str, Path] = {
        "plot_catalogue": output_dir / "plot_catalogue.csv",
        "plot_manifest": output_dir / "plot_manifest.json",
        "plotting_summary": output_dir / "plotting_summary.txt",
    }
    for item in result.rendered_plots:
        outputs[f"plot:{item.plot_id}"] = Path(item.output_file)
        if item.svg_file is not None:
            outputs[f"plot_svg:{item.plot_id}"] = Path(item.svg_file)
    return outputs


def _excel_outputs(result: ExcelReportResult) -> dict[str, Path]:
    outputs = {
        "report": result.report_path,
        "excel_report_manifest": result.manifest_path,
        "excel_report_summary": result.summary_path,
    }
    if result.plot_assets_dir.exists():
        outputs["plot_assets"] = result.plot_assets_dir
    return outputs


def _powerpoint_outputs(result: PowerPointReportResult) -> dict[str, Path]:
    return {
        "presentation": result.presentation_path,
        "powerpoint_report_manifest": result.manifest_path,
        "powerpoint_report_summary": result.summary_path,
    }


_GENERATED_OUTPUT_ENTRIES = {
    "01_inspection",
    "02_channel_selection",
    "03_math_channels",
    "04_statistics",
    "05_plots",
    "06_excel_report",
    "07_powerpoint_report",
    "pipeline_manifest.json",
    "pipeline_summary.txt",
    "pipeline.log",
    ".vsm_postprocessing_output",
}


def _prepare_output_root(path: Path, clean: bool) -> None:
    resolved = path.resolve()
    if resolved == Path(resolved.anchor):
        raise ConfigurationError("Refusing to use a filesystem root as output.root_dir")
    resolved.mkdir(parents=True, exist_ok=True)
    if clean:
        # Client-safety rule: never delete the whole output root. Only remove artifacts
        # whose names are owned by this pipeline, preserving any user-created files.
        for name in _GENERATED_OUTPUT_ENTRIES:
            item = resolved / name
            if item.is_dir():
                shutil.rmtree(item)
            elif item.exists():
                item.unlink()
    (resolved / ".vsm_postprocessing_output").write_text(
        "This directory contains generated VSM post-processing outputs.\n",
        encoding="utf-8",
    )

def _resolve_required_path(config_path: Path, value: Any, context: str) -> Path:
    path = _resolve_path_value(config_path, value, context)
    if not path.exists():
        raise ConfigurationError(f"{context} does not exist: {path}")
    if not path.is_file():
        raise ConfigurationError(f"{context} is not a file: {path}")
    return path


def _resolve_path_value(config_path: Path, value: Any, context: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be a non-empty path string")
    candidate = Path(value.strip()).expanduser()
    if not candidate.is_absolute():
        candidate = config_path.parent / candidate
    return candidate.resolve()


def _optional_string(value: Any, context: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise ConfigurationError(f"{context} must be null or a non-empty string")
    return value.strip()


def _optional_positive_int(value: Any, context: str) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise ConfigurationError(f"{context} must be null or an integer >= 1")
    return value


def _reject_unknown_keys(mapping: dict[str, Any], allowed: set[str], context: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: " + ", ".join(unknown))
