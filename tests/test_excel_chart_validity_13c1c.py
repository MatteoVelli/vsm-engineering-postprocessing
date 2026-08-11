from __future__ import annotations

import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest


PROJECT_ROOT = Path(__file__).resolve().parents[1]
REPORT = PROJECT_ROOT / "outputs" / "end_to_end_sergio_duty_cycle" / "07_excel_report" / "vsm_engineering_report.xlsx"
NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _chart_names(archive: zipfile.ZipFile) -> list[str]:
    return sorted(
        (name for name in archive.namelist() if name.startswith("xl/charts/chart") and name.endswith(".xml")),
        key=lambda value: int(re.search(r"chart(\d+)\.xml", value).group(1)),
    )


def _axis_id(axis: ET.Element) -> int:
    return int(axis.find("c:axId", NS).get("val"))


def _axis_pos(axis: ET.Element) -> str:
    return axis.find("c:axPos", NS).get("val")


def _cross_axis_id(axis: ET.Element) -> int:
    return int(axis.find("c:crossAx", NS).get("val"))


def _crosses(axis: ET.Element) -> str | None:
    node = axis.find("c:crosses", NS)
    return None if node is None else node.get("val")


def _is_deleted(axis: ET.Element) -> bool:
    node = axis.find("c:delete", NS)
    return node is not None and node.get("val") in {"1", "true"}


@pytest.fixture(scope="module")
def report_archive():
    if not REPORT.exists():
        pytest.skip("Generated Sergio duty-cycle workbook is not present")
    with zipfile.ZipFile(REPORT) as archive:
        yield archive


def test_13c1c_workbook_zip_integrity_passes(report_archive: zipfile.ZipFile) -> None:
    assert report_archive.testzip() is None


def test_13c1c_chart_and_drawing_relationship_targets_exist(report_archive: zipfile.ZipFile) -> None:
    names = set(report_archive.namelist())
    rel_files = [name for name in names if name.startswith("xl/") and "/_rels/" in name and name.endswith(".rels")]
    assert rel_files
    for rel_file in rel_files:
        rel_root = ET.fromstring(report_archive.read(rel_file))
        source_dir = rel_file.replace("/_rels/", "/").removesuffix(".rels").rsplit("/", 1)[0]
        for rel in rel_root.findall("rel:Relationship", NS):
            target = rel.get("Target")
            if not target or "://" in target or target.startswith("/"):
                continue
            resolved = posixpath.normpath(posixpath.join(source_dir, target))
            assert resolved in names, f"{rel_file} points to missing {resolved}"


def test_13c1c_all_native_chart_axis_graphs_are_excel_valid(report_archive: zipfile.ZipFile) -> None:
    chart_names = _chart_names(report_archive)
    assert len(chart_names) == 18

    total_axes = 0
    secondary_chart_count = 0
    unresolved_cross_axes: list[tuple[str, int]] = []

    for chart_name in chart_names:
        root = ET.fromstring(report_archive.read(chart_name))
        scatter_groups = root.findall(".//c:scatterChart", NS)
        axes = root.findall(".//c:valAx", NS)
        axis_by_id = {_axis_id(axis): axis for axis in axes}
        axis_ids = set(axis_by_id)
        total_axes += len(axis_ids)

        assert scatter_groups, chart_name
        assert len(axis_ids) == len(axes), chart_name
        assert len(axis_ids) in {2, 4}, chart_name

        referenced_by_groups: set[int] = set()
        for group in scatter_groups:
            group_axis_ids = [int(node.get("val")) for node in group.findall("c:axId", NS)]
            assert len(group_axis_ids) == 2, chart_name
            assert set(group_axis_ids) <= axis_ids, chart_name
            x_id, y_id = group_axis_ids
            assert _cross_axis_id(axis_by_id[x_id]) == y_id, chart_name
            assert _cross_axis_id(axis_by_id[y_id]) == x_id, chart_name
            referenced_by_groups.update(group_axis_ids)

        assert referenced_by_groups == axis_ids, chart_name

        for axis_id, axis in axis_by_id.items():
            cross_id = _cross_axis_id(axis)
            if cross_id not in axis_ids:
                unresolved_cross_axes.append((chart_name, cross_id))
                continue
            assert _cross_axis_id(axis_by_id[cross_id]) == axis_id, chart_name

        first_group_ids = [int(node.get("val")) for node in scatter_groups[0].findall("c:axId", NS)]
        primary_x = axis_by_id[first_group_ids[0]]
        primary_y = axis_by_id[first_group_ids[1]]
        assert _axis_pos(primary_x) == "b", chart_name
        assert _axis_pos(primary_y) == "l", chart_name
        assert _crosses(primary_x) == "autoZero", chart_name
        assert _crosses(primary_y) == "autoZero", chart_name

        if len(scatter_groups) > 1:
            secondary_chart_count += 1
            second_group_ids = [int(node.get("val")) for node in scatter_groups[1].findall("c:axId", NS)]
            secondary_x = axis_by_id[second_group_ids[0]]
            secondary_y = axis_by_id[second_group_ids[1]]
            assert _axis_pos(secondary_x) == "b", chart_name
            assert _axis_pos(secondary_y) == "r", chart_name
            assert _is_deleted(secondary_x), chart_name
            assert _crosses(secondary_y) == "max", chart_name
        else:
            assert len(axis_ids) == 2, chart_name

    assert secondary_chart_count == 12
    assert unresolved_cross_axes == []
    assert total_axes == 60
