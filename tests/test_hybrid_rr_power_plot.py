"""Hybrid reuses the Electric presentation override with actual RR samples."""
from pathlib import Path

import pytest

from test_electric_rr_power_plot import (
    build_asymmetric_rr_reports,
    test_all_six_cards_preserved_and_rr_sources_correct as _assert_cards,
    test_only_target_image_changes_and_slide_xml_preserves_layout as _assert_layout,
    test_real_renderer_receives_asymmetric_rr_samples as _assert_samples,
    test_rendered_plot_and_metadata_select_actual_rr as _assert_rendered_plot,
)
from vsm_postprocessing.report_profile import load_reporting_profile


PROFILE = Path(__file__).resolve().parents[1] / "config/report_profiles/robosprayer_hybrid.yaml"


@pytest.fixture(scope="module")
def reports(tmp_path_factory):
    return build_asymmetric_rr_reports(tmp_path_factory.mktemp("asymmetric_hybrid_rr_power"), PROFILE)


def test_hybrid_presentation_override_keeps_shared_excel_rl_plot():
    profile = load_reporting_profile(PROFILE)
    plot = "power_at_wheels_and_edu"
    assert tuple(s.semantic_name for s in profile.plots_by_id()[plot].series) == ("wheel_power_rl", "edu_mech_power_rl")
    assert tuple(s.semantic_name for s in profile.presentation.plot_series_overrides[plot]) == ("wheel_power_rr", "edu_mech_power_rr")
    assert tuple(s.label for s in profile.presentation.plot_series_overrides[plot]) == ("Wheel Power RR", "EDU Mech Power RR")


def test_hybrid_rendered_plot_embeds_rr_and_excludes_rl(reports):
    _assert_rendered_plot(reports)


def test_hybrid_actual_asymmetric_rr_arrays_reach_renderer(reports, tmp_path, monkeypatch):
    _assert_samples(reports, tmp_path, monkeypatch)


def test_hybrid_rr_power_cards_and_other_four_cards_preserved(reports):
    _assert_cards(reports)


def test_hybrid_only_power_graph_changes_with_all_slide_layouts_preserved(reports):
    _assert_layout(reports)
