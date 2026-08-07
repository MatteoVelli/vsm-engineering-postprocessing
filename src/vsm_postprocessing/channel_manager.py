from __future__ import annotations

import csv
import difflib
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

import numpy as np
import yaml

from .errors import ChannelSelectionError, ConfigurationError
from .importer import ImportOptions, load_data_file
from .models import ChannelInfo, ImportedDataset


@dataclass(frozen=True)
class SelectionOutputOptions:
    """Output controls for a channel-selection run."""

    data_filename: str = "selected_channels.csv"
    include_units_row: bool = True
    float_precision: int = 12


@dataclass(frozen=True)
class ChannelSelectionConfig:
    """Validated configuration for selecting channels by stable channel ID."""

    version: int
    export_channels: tuple[str, ...]
    include_time: bool
    output: SelectionOutputOptions


@dataclass
class ChannelSelectionResult:
    """Selected channel metadata and numeric data in configured order."""

    dataset: ImportedDataset
    selected_channels: list[ChannelInfo]
    selected_values: np.ndarray
    config_path: Path
    config: ChannelSelectionConfig

    @property
    def sample_count(self) -> int:
        return int(self.selected_values.shape[0])

    @property
    def channel_count(self) -> int:
        return int(self.selected_values.shape[1])


def load_selection_config(path: str | Path) -> ChannelSelectionConfig:
    """Load and strictly validate a versioned YAML selection file."""

    config_path = Path(path).expanduser().resolve()
    if not config_path.exists():
        raise ConfigurationError(f"Selection configuration does not exist: {config_path}")
    if not config_path.is_file():
        raise ConfigurationError(f"Selection configuration is not a file: {config_path}")

    try:
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except UnicodeDecodeError as exc:
        raise ConfigurationError(f"Selection configuration must be UTF-8: {config_path}") from exc
    except yaml.YAMLError as exc:
        raise ConfigurationError(f"Invalid YAML in selection configuration '{config_path}': {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigurationError("Selection configuration root must be a YAML mapping")

    _reject_unknown_keys(raw, {"version", "selection", "output"}, "root")
    version = raw.get("version")
    if version != 1:
        raise ConfigurationError("Selection configuration 'version' must be 1")

    selection = raw.get("selection")
    if not isinstance(selection, dict):
        raise ConfigurationError("Selection configuration must contain a 'selection' mapping")
    _reject_unknown_keys(selection, {"include_time", "export_channels"}, "selection")

    include_time = selection.get("include_time", True)
    if not isinstance(include_time, bool):
        raise ConfigurationError("selection.include_time must be true or false")

    export_channels_raw = selection.get("export_channels")
    if not isinstance(export_channels_raw, list) or not export_channels_raw:
        raise ConfigurationError("selection.export_channels must be a non-empty YAML list")
    if any(not isinstance(value, str) or not value.strip() for value in export_channels_raw):
        raise ConfigurationError("Every selection.export_channels entry must be a non-empty channel_id string")

    export_channels = tuple(value.strip() for value in export_channels_raw)
    duplicates = _duplicates(export_channels)
    if duplicates:
        raise ConfigurationError(
            "selection.export_channels contains duplicate channel IDs: " + ", ".join(duplicates)
        )

    output_raw = raw.get("output", {})
    if output_raw is None:
        output_raw = {}
    if not isinstance(output_raw, dict):
        raise ConfigurationError("output must be a YAML mapping")
    _reject_unknown_keys(output_raw, {"data_filename", "include_units_row", "float_precision"}, "output")

    data_filename = output_raw.get("data_filename", "selected_channels.csv")
    if not isinstance(data_filename, str) or not data_filename.strip():
        raise ConfigurationError("output.data_filename must be a non-empty string")
    data_filename = data_filename.strip()
    filename_path = Path(data_filename)
    if filename_path.name != data_filename or filename_path.suffix.lower() != ".csv":
        raise ConfigurationError("output.data_filename must be a plain .csv filename without directories")

    include_units_row = output_raw.get("include_units_row", True)
    if not isinstance(include_units_row, bool):
        raise ConfigurationError("output.include_units_row must be true or false")

    float_precision = output_raw.get("float_precision", 12)
    if isinstance(float_precision, bool) or not isinstance(float_precision, int) or not 1 <= float_precision <= 17:
        raise ConfigurationError("output.float_precision must be an integer from 1 to 17")

    return ChannelSelectionConfig(
        version=version,
        export_channels=export_channels,
        include_time=include_time,
        output=SelectionOutputOptions(
            data_filename=data_filename,
            include_units_row=include_units_row,
            float_precision=float_precision,
        ),
    )


def select_channels(
    input_file: str | Path,
    config_file: str | Path,
    import_options: ImportOptions | None = None,
) -> ChannelSelectionResult:
    """Load a valid dataset and select configured channels in deterministic order."""

    config_path = Path(config_file).expanduser().resolve()
    config = load_selection_config(config_path)
    dataset = load_data_file(input_file, import_options)

    channels_by_id = {channel.channel_id: channel for channel in dataset.channels}
    requested_ids = list(config.export_channels)

    if config.include_time:
        time_id = dataset.quality.time_channel_id
        if time_id is None:
            raise ChannelSelectionError("The imported dataset does not define a time channel")
        requested_ids = [time_id, *[channel_id for channel_id in requested_ids if channel_id != time_id]]

    missing = [channel_id for channel_id in requested_ids if channel_id not in channels_by_id]
    if missing:
        raise ChannelSelectionError(_format_missing_channels(missing, dataset.channels))

    source_indices = [channels_by_id[channel_id].source_column_index - 1 for channel_id in requested_ids]
    selected_values = dataset.values[:, source_indices].copy()
    selected_channels = [channels_by_id[channel_id] for channel_id in requested_ids]

    if selected_values.shape != (dataset.quality.sample_count, len(selected_channels)):
        raise ChannelSelectionError(
            "Internal selection shape mismatch: "
            f"expected {(dataset.quality.sample_count, len(selected_channels))}, got {selected_values.shape}"
        )

    return ChannelSelectionResult(
        dataset=dataset,
        selected_channels=selected_channels,
        selected_values=selected_values,
        config_path=config_path,
        config=config,
    )


def export_channel_selection(result: ChannelSelectionResult, output_dir: str | Path) -> dict[str, Path]:
    """Export selected numeric data, metadata, manifest and a concise summary."""

    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    data_path = output_path / result.config.output.data_filename
    catalogue_path = output_path / "selected_channel_catalogue.csv"
    manifest_path = output_path / "selection_manifest.json"
    summary_path = output_path / "selection_summary.txt"

    _write_selected_data_csv(result, data_path)
    _write_selected_catalogue(result, catalogue_path)
    manifest_path.write_text(json.dumps(_build_manifest(result, data_path), indent=2), encoding="utf-8")
    summary_path.write_text(_format_selection_summary(result, data_path), encoding="utf-8")

    return {
        "selected_data": data_path,
        "selected_channel_catalogue": catalogue_path,
        "selection_manifest": manifest_path,
        "selection_summary": summary_path,
    }


def _write_selected_data_csv(result: ChannelSelectionResult, path: Path) -> None:
    precision = result.config.output.float_precision
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow([channel.channel_id for channel in result.selected_channels])
        if result.config.output.include_units_row:
            writer.writerow([channel.unit or "" for channel in result.selected_channels])
        for row in result.selected_values:
            writer.writerow([_format_float(float(value), precision) for value in row])


def _write_selected_catalogue(result: ChannelSelectionResult, path: Path) -> None:
    fieldnames = [
        "selection_order",
        "channel_id",
        "source_name",
        "display_name",
        "unit",
        "source_column_index",
        "source_column_label",
        "kind",
        "dtype",
        "provenance",
        "dependencies",
        "formula_example",
    ]
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for order, channel in enumerate(result.selected_channels, start=1):
            row = channel.to_dict()
            row["selection_order"] = order
            row["dependencies"] = ";".join(row["dependencies"])
            writer.writerow(row)


def _build_manifest(result: ChannelSelectionResult, data_path: Path) -> dict[str, Any]:
    quality = result.dataset.quality
    return {
        "configuration_version": result.config.version,
        "source_file": str(result.dataset.source_path),
        "source_sha256": quality.source_sha256,
        "configuration_file": str(result.config_path),
        "sample_count": result.sample_count,
        "selected_channel_count": result.channel_count,
        "time_channel_id": quality.time_channel_id,
        "time_start": quality.time_start,
        "time_end": quality.time_end,
        "nominal_time_step": quality.nominal_time_step,
        "selected_channel_ids": [channel.channel_id for channel in result.selected_channels],
        "selected_channels": [channel.to_dict() for channel in result.selected_channels],
        "output_data_file": str(data_path),
        "include_units_row": result.config.output.include_units_row,
        "float_precision": result.config.output.float_precision,
    }


def _format_selection_summary(result: ChannelSelectionResult, data_path: Path) -> str:
    quality = result.dataset.quality
    lines = [
        "VSM CHANNEL SELECTION",
        "=====================",
        "Status: PASS",
        f"Source: {result.dataset.source_path}",
        f"Source SHA-256: {quality.source_sha256}",
        f"Configuration: {result.config_path}",
        f"Samples: {result.sample_count}",
        f"Selected channels: {result.channel_count}",
        f"Time range: {quality.time_start} to {quality.time_end} {quality.time_unit or ''}".rstrip(),
        f"Output data: {data_path}",
        "",
        "Selection order:",
    ]
    lines.extend(
        f"{index:02d}. {channel.channel_id} | {channel.display_name} [{channel.unit or '-'}] | {channel.kind}"
        for index, channel in enumerate(result.selected_channels, start=1)
    )
    return "\n".join(lines) + "\n"


def _format_missing_channels(missing: Sequence[str], channels: Sequence[ChannelInfo]) -> str:
    available_ids = [channel.channel_id for channel in channels]
    available_names: dict[str, list[str]] = {}
    for channel in channels:
        available_names.setdefault(channel.source_name.lower(), []).append(channel.channel_id)
        available_names.setdefault(channel.display_name.lower(), []).append(channel.channel_id)

    lines = ["Requested channel IDs were not found:"]
    for missing_id in missing:
        suggestions = difflib.get_close_matches(missing_id, available_ids, n=3, cutoff=0.45)
        exact_name_ids = available_names.get(missing_id.lower(), [])
        if exact_name_ids:
            suggestions = list(dict.fromkeys([*exact_name_ids, *suggestions]))[:3]
        if suggestions:
            lines.append(f"- {missing_id} (possible channel_id: {', '.join(suggestions)})")
        else:
            lines.append(f"- {missing_id}")
    lines.append("Use channel_id values from channel_catalogue.csv; display names are not unique keys.")
    return "\n".join(lines)


def _format_float(value: float, precision: int) -> str:
    if not np.isfinite(value):
        raise ChannelSelectionError("Selected data contains a non-finite value after successful validation")
    return format(value, f".{precision}g")


def _duplicates(values: Sequence[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for value in values:
        if value in seen and value not in duplicates:
            duplicates.append(value)
        seen.add(value)
    return duplicates


def _reject_unknown_keys(mapping: dict[str, Any], allowed: set[str], context: str) -> None:
    unknown = sorted(set(mapping) - allowed)
    if unknown:
        raise ConfigurationError(f"Unknown key(s) in {context}: {', '.join(unknown)}")
