"""Sergio's exact-source unit declaration must preserve raw values and metadata."""
from dataclasses import replace
import json

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
import pytest
import yaml

from test_sergio_october_15 import dataset_for, profile_for, profile_path, csv_for, remove_channel
from vsm_postprocessing.errors import ConfigurationError
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from vsm_postprocessing.profile_math import calculate_profile_math_channels
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.profile_powerpoint_report_engine import build_profile_powerpoint_report
from vsm_postprocessing.report_profile import load_reporting_profile, resolve_profile

HV = 'auxiliary_high_voltage_power_consumption'
LV = 'electricsystem_auxiliary_lowvoltagepowerconsumption'
SOURCE = 'User_defined_Data_High_Voltage_Auxiliary_signal1_0'


def fixture(kind, unit='-'):
    data = dataset_for(kind, (0, 2, 5, 9))
    index = data.channel_index(HV)
    data.channels[index] = replace(data.channels[index], unit=unit)
    data.values[:, index] = [8., -4., -10., 10.]
    data.values[:, data.channel_index(LV)] = [2., 1., 3., 2.]
    return data


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
@pytest.mark.parametrize('unit', ['-', 'kW'])
def test_confirmed_unit_preserves_raw_metadata_values_and_signed_irregular_energy(kind, unit):
    data, profile = fixture(kind, unit), profile_for(kind)
    original = data.channels[data.channel_index(HV)]
    measurements = data.values.copy()
    result = calculate_profile_math_channels(data, profile)
    resolved = result.resolution.resolved[HV]
    assert result.resolution.is_valid and result.is_complete
    assert resolved.channel is original and original.unit == unit
    assert resolved.effective_unit == resolved.engineering_channel.unit == 'kW'
    np.testing.assert_array_equal(data.values, measurements)
    if unit == '-':
        assert resolved.unit_interpretation.authority.startswith('Sergio, Astauto, 16 October 2026')
        assert "source unit '-'" in resolved.unit_interpretation_note
        assert 'factor 1, offset 0' in resolved.unit_interpretation_note
    else:
        assert resolved.unit_interpretation is None
    assert not any('unit mismatch' in d for d in result.resolution.diagnostics)
    np.testing.assert_array_equal(result.values_by_semantic_name['total_auxiliary_power'], [10., -3., -7., 12.])
    # Independent trapezoids: 7, -15, 10 kW*s. LAST differs from MAX,
    # so this also detects clamping, double counting and a stale MAX statistic.
    np.testing.assert_allclose(result.values_by_semantic_name['auxiliary_energy_consumption'], np.array([0., 7., -15., 10.]) / 3600)
    np.testing.assert_allclose(result.values_by_semantic_name['auxiliary_energy_consumption_accumulated'], np.array([0., 7., -8., 2.]) / 3600)
    stats = calculate_profile_statistics(data, profile, math_result=result)
    by_id = {s.definition.statistic_id: s.value for s in stats.statistics}
    assert by_id['total_auxiliary_power_max'] == 12.
    assert by_id['auxiliary_energy_accumulated_max'] == pytest.approx(2 / 3600)


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
@pytest.mark.parametrize('unit', ['W', 'N', None, ''])
def test_other_or_absent_source_units_still_fail_strict_validation(kind, unit):
    resolution = resolve_profile(fixture(kind, unit), profile_for(kind))
    assert HV not in resolution.resolved
    assert any(i.definition.semantic_name == HV and 'unit mismatch' in i.reason for i in resolution.missing_optional)


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
def test_declaration_never_leaks_to_other_channels_aliases_or_normalized_names(kind):
    profile = profile_for(kind)
    data = fixture(kind)
    index = data.channel_index(LV)
    data.channels[index] = replace(data.channels[index], unit='-')
    resolution = resolve_profile(data, profile)
    assert HV in resolution.resolved and LV not in resolution.resolved
    assert any(i.definition.semantic_name == LV for i in resolution.unit_mismatches)
    for source in ('Other_User_Defined_Channel', SOURCE.lower()):
        data = fixture(kind)
        index = data.channel_index(HV)
        data.channels[index] = replace(data.channels[index], source_name=source)
        custom = replace(profile, raw_channels=tuple(replace(c, aliases=(source,)) if c.semantic_name == HV else c for c in profile.raw_channels))
        assert HV not in resolve_profile(data, custom).resolved


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
def test_missing_confirmed_source_keeps_optional_dependency_diagnostics(kind):
    data = fixture(kind)
    remove_channel(data, HV)
    result = calculate_profile_math_channels(data, profile_for(kind))
    assert result.is_complete and result.resolution.is_valid
    assert 'total_auxiliary_power' not in result.values_by_semantic_name
    assert any(SOURCE in d and 'missing' in d for d in result.resolution.diagnostics)
    assert not any('Engineering unit interpretation' in d for d in result.resolution.diagnostics)


@pytest.mark.parametrize('patch', [
    {'engineering_unit': 'W'}, {'authority': ''}, {'source_unit': ''}, {'scale': 1000},
])
def test_malformed_or_scaling_declarations_are_rejected(tmp_path, patch):
    interpretation = {'source_unit': '-', 'engineering_unit': 'kW', 'authority': 'Confirmed engineering authority'} | patch
    payload = {'version': 1, 'profile': {'profile_id': 'test', 'name': 'Test', 'extends': str(profile_path('electric').resolve())},
               'overrides': {'raw': {HV: {'unit_interpretation': interpretation}}}}
    path = tmp_path / 'invalid.yaml'
    path.write_text(yaml.safe_dump(payload), encoding='utf-8')
    with pytest.raises(ConfigurationError, match='unit_interpretation'):
        load_reporting_profile(path)


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
def test_restored_report_formulas_provenance_and_powerpoint_chart(kind, tmp_path):
    data = fixture(kind)
    report = generate_profile_excel_report(csv_for(data, tmp_path / 'confirmed.csv'), profile_path(kind), tmp_path / 'excel')
    assert 'auxiliaries_energy_consumption' in {p.plot_id for p in report.plotting_result.rendered_plots}
    assert report.plotting_result.channels_by_semantic_name[HV].unit == 'kW'
    manifest = json.loads(report.manifest_path.read_text())
    interpretation = manifest['source_unit_interpretations']
    assert len(interpretation) == 1 and interpretation[0]['source_name'] == SOURCE
    assert interpretation[0]['source_unit'] == '-' and interpretation[0]['engineering_unit'] == 'kW'
    assert interpretation[0]['values_unchanged']
    book, cached = load_workbook(report.report_path), load_workbook(report.report_path, data_only=True)
    try:
        columns = {c.channel_id: i for i, c in enumerate(report.report_channels, 1)}
        assert book.worksheets[0].cell(4, columns[HV]).value == 'kW'
        for name in ('total_auxiliary_power', 'auxiliary_energy_consumption', 'auxiliary_energy_consumption_accumulated'):
            assert book.worksheets[0].cell(6, columns[name]).data_type == 'f'
        for offset in range(4):
            assert cached.worksheets[0].cell(5 + offset, columns[HV]).value == data.values[offset, data.channel_index(HV)]
        for row in book['Rename From VSM to Astauto'].iter_rows(min_row=3):
            if row[5].value == HV:
                assert "source unit '-'" in row[7].value and 'Sergio' in row[7].value
        # Bottom statistics include MAX, MIN and LAST using the correct raw kW.
        main = cached.worksheets[0]
        column = columns[HV]
        numbers = [main.cell(row, column).value for row in range(9, main.max_row + 1)]
        assert 10. in numbers and -10. in numbers
    finally:
        book.close()
        cached.close()
    ppt = build_profile_powerpoint_report(report, tmp_path / 'ppt')
    deck = Presentation(ppt.presentation_path)
    if kind == 'electric':
        assert not any('PLOT DATA NOT AVAILABLE' in s.text for s in deck.slides[6].shapes if s.has_text_frame)
        assert sum(s.shape_type == 13 for s in deck.slides[6].shapes) >= 2


def test_diesel_and_legacy_profiles_have_no_unit_interpretations():
    assert not any(c.unit_interpretation for c in profile_for('diesel').raw_channels)
    for kind in ('electric', 'hybrid'):
        profile = load_reporting_profile(profile_path(kind).parent.parent / profile_path(kind).name)
        assert not any(c.unit_interpretation for c in profile.raw_channels)
