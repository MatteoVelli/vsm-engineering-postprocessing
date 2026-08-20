from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from .report_profile import ReportingProfile
from .utils import client_display_filename

_SEPARATOR_PATTERN = re.compile(r"[_\-\s]+")
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
_WEIGHT_PATTERN = re.compile(r"^\d+(?:[.,]\d+)?kg$", re.IGNORECASE)
_POWER_PATTERN = re.compile(r"^\d+(?:[.,]\d+)?kw$", re.IGNORECASE)
_SPEED_PATTERN = re.compile(r"^\d+(?:[.,]\d+)?kph$", re.IGNORECASE)
_RPM_PATTERN = re.compile(r"^\d+(?:[.,]\d+)?rpm$", re.IGNORECASE)
_DISTANCE_PATTERN = re.compile(r"^\d+x\d+km$", re.IGNORECASE)

_POWERTRAIN_ALIASES = {
    "electric": "Electric",
    "hybrid": "Hybrid",
    "diesel": "Diesel",
}
_KNOWN_MACHINE_ALIASES = {
    "robosprayer": "RoboSprayer",
    "robo sprayer": "RoboSprayer",
    "caiman sp": "Caiman SP",
}
_GENERIC_PREFIXES = {"sprayer", "vehicle", "machine", "vsm"}
_NON_IDENTITY_TOKENS = {
    "batt",
    "battery",
    "chem",
    "cool",
    "crop",
    "discharge",
    "field",
    "gen",
    "grad",
    "gradient",
    "ha",
    "ht",
    "mot",
    "motor",
    "rough",
    "soc",
    "susp",
    "test",
}


@dataclass(frozen=True)
class ReportMetadata:
    machine_name: str
    powertrain_name: str
    report_title: str
    safe_output_stem: str
    source_filename: str
    detection_source: str
    simulation_description: str | None = None


def resolve_report_metadata(
    source_file: str | Path,
    profile: ReportingProfile | Any | None = None,
    *,
    machine_name_override: str | None = None,
    source_metadata: Mapping[str, Any] | None = None,
) -> ReportMetadata:
    source_filename = client_display_filename(source_file)
    powertrain_name = _powertrain_name(profile, source_filename)
    machine_name, detection_source = _resolve_machine_name(
        source_filename,
        machine_name_override=machine_name_override,
        source_metadata=source_metadata,
    )
    report_title = f"{machine_name} {powertrain_name}".strip()
    return ReportMetadata(
        machine_name=machine_name,
        powertrain_name=powertrain_name,
        report_title=report_title,
        safe_output_stem=safe_filename_stem(report_title),
        source_filename=source_filename,
        detection_source=detection_source,
        simulation_description=_simulation_description(source_filename, machine_name, powertrain_name),
    )


def safe_filename_stem(value: str) -> str:
    stem = _INVALID_FILENAME_CHARS.sub("_", value.strip())
    stem = _SEPARATOR_PATTERN.sub("_", stem).strip(" ._")
    while "__" in stem:
        stem = stem.replace("__", "_")
    return stem or "VSM_Profile"


def _resolve_machine_name(
    source_filename: str,
    *,
    machine_name_override: str | None,
    source_metadata: Mapping[str, Any] | None,
) -> tuple[str, str]:
    override = _clean_display_text(machine_name_override)
    if override:
        return _normalize_known_machine(override), "user_override"

    metadata_name = _machine_name_from_source_metadata(source_metadata)
    if metadata_name:
        return _normalize_known_machine(metadata_name), "source_metadata"

    filename_name = _machine_name_from_filename(source_filename)
    if filename_name:
        return filename_name, "source_filename"

    return _fallback_from_source_stem(source_filename), "source_filename_fallback"


def _machine_name_from_source_metadata(source_metadata: Mapping[str, Any] | None) -> str | None:
    if not source_metadata:
        return None
    for key in ("machine_name", "vehicle_name", "machine", "vehicle"):
        value = _clean_display_text(source_metadata.get(key))
        if value:
            return value
    return None


def _machine_name_from_filename(source_filename: str) -> str | None:
    stem = Path(source_filename).stem
    tokens = _filename_tokens(stem)
    if len(tokens) > 1 and tokens[0].lower() in _GENERIC_PREFIXES:
        tokens = tokens[1:]

    identity: list[str] = []
    for token in tokens:
        if _is_parameter_or_powertrain_token(token):
            break
        identity.append(token)

    if not identity:
        return None
    return _normalize_known_machine(_display_name_from_tokens(identity))


def _fallback_from_source_stem(source_filename: str) -> str:
    tokens = _filename_tokens(Path(source_filename).stem)
    if not tokens:
        return "Unidentified Machine"
    return _normalize_known_machine(_display_name_from_tokens(tokens[:4]))


def _filename_tokens(stem: str) -> list[str]:
    return [token for token in _SEPARATOR_PATTERN.split(stem.strip()) if token]


def _is_parameter_or_powertrain_token(token: str) -> bool:
    normalized = token.strip().lower()
    if normalized in _POWERTRAIN_ALIASES or normalized in _NON_IDENTITY_TOKENS:
        return True
    return bool(
        _WEIGHT_PATTERN.match(normalized)
        or _POWER_PATTERN.match(normalized)
        or _SPEED_PATTERN.match(normalized)
        or _RPM_PATTERN.match(normalized)
        or _DISTANCE_PATTERN.match(normalized)
    )


def _display_name_from_tokens(tokens: list[str]) -> str:
    return " ".join(_display_token(token) for token in tokens).strip()


def _display_token(token: str) -> str:
    if token.isupper() and len(token) <= 4:
        return token
    if any(character.islower() for character in token) and any(character.isupper() for character in token):
        return token
    if len(token) <= 3 and token.isalpha():
        return token.upper()
    return token[:1].upper() + token[1:].lower()


def _normalize_known_machine(value: str) -> str:
    cleaned = _SEPARATOR_PATTERN.sub(" ", value.strip())
    alias = _KNOWN_MACHINE_ALIASES.get(cleaned.lower())
    return alias or cleaned


def _powertrain_name(profile: ReportingProfile | Any | None, source_filename: str) -> str:
    raw = None
    if profile is not None:
        metadata = getattr(profile, "metadata", None)
        raw = getattr(metadata, "powertrain", None) if metadata is not None else getattr(profile, "powertrain", None)
    normalized = (raw or "").strip().lower()
    if normalized in _POWERTRAIN_ALIASES:
        return _POWERTRAIN_ALIASES[normalized]

    for token in _filename_tokens(Path(source_filename).stem):
        candidate = token.lower()
        if candidate in _POWERTRAIN_ALIASES:
            return _POWERTRAIN_ALIASES[candidate]
    return "Engineering"


def _simulation_description(source_filename: str, machine_name: str, powertrain_name: str) -> str | None:
    stem = Path(source_filename).stem
    machine_stem = safe_filename_stem(machine_name).replace("_", " ")
    description = _SEPARATOR_PATTERN.sub(" ", stem).strip()
    for text in (machine_stem, machine_name, powertrain_name):
        description = re.sub(re.escape(text), "", description, flags=re.IGNORECASE).strip()
    return description or None


def _clean_display_text(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    cleaned = _SEPARATOR_PATTERN.sub(" ", value.strip())
    return cleaned or None
