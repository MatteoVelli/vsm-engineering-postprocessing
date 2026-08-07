from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import numpy as np


@dataclass(frozen=True)
class ChannelInfo:
    """Metadata for one imported channel."""

    channel_id: str
    source_name: str
    display_name: str
    unit: str | None
    source_column_index: int
    source_column_label: str
    kind: str
    dtype: str
    provenance: str
    dependencies: tuple[str, ...] = ()
    formula_example: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["dependencies"] = list(self.dependencies)
        return data


@dataclass
class DataQualityReport:
    """Deterministic validation results for an imported dataset."""

    source_file: str
    source_sha256: str
    file_type: str
    sheet_name: str | None
    header_row: int
    unit_row: int | None
    data_start_row: int
    data_end_row: int
    sample_count: int
    channel_count: int
    raw_channel_count: int
    math_channel_count: int
    time_channel_id: str | None
    time_channel_name: str | None
    time_unit: str | None
    time_start: float | None
    time_end: float | None
    nominal_time_step: float | None
    time_is_strictly_increasing: bool
    duplicate_timestamp_count: int
    missing_cell_count: int
    invalid_numeric_cell_count: int
    non_finite_cell_count: int
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.errors

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["is_valid"] = self.is_valid
        return data


@dataclass
class DataInspectionResult:
    """Metadata and validation output produced by the import milestone."""

    source_path: Path
    channels: list[ChannelInfo]
    quality: DataQualityReport

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_path": str(self.source_path),
            "channels": [channel.to_dict() for channel in self.channels],
            "quality": self.quality.to_dict(),
        }


@dataclass
class ImportedDataset:
    """Validated numeric dataset with channel metadata in source-column order."""

    source_path: Path
    channels: list[ChannelInfo]
    quality: DataQualityReport
    values: np.ndarray

    def __post_init__(self) -> None:
        expected_shape = (self.quality.sample_count, self.quality.channel_count)
        if self.values.shape != expected_shape:
            raise ValueError(f"Dataset shape {self.values.shape} does not match metadata {expected_shape}")

    def channel_index(self, channel_id: str) -> int:
        for index, channel in enumerate(self.channels):
            if channel.channel_id == channel_id:
                return index
        raise KeyError(channel_id)
