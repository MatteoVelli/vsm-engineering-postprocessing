"""Declared nominal energy, qualified inference and shared range reporting."""
from dataclasses import replace
import hashlib
import json

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
import pytest

from vsm_postprocessing.battery import nominal_capacity_estimates_kwh
from vsm_postprocessing.diagnostics import diagnostic_records
from vsm_postprocessing.errors import ConfigurationError, StatisticsError
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report
from vsm_postprocessing.profile_powerpoint_report_engine import build_profile_powerpoint_report
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.report_metadata import resolve_report_metadata
from vsm_postprocessing.reporting_inputs import ReportingInputs
from test_profile_statistics import _battery_profile_dataset, _battery_range_profile
from test_sergio_october_15 import dataset_for, csv_for, profile_for, profile_path


def confirmed():
    return ReportingInputs(battery_nominal_capacity_kwh=350,
        battery_nominal_capacity_status='confirmed',
        battery_nominal_capacity_provenance='Sergio, ASTAUTO, 2 November 2026; temperature-dependent nominal reference, law unspecified')


def values(result):
    return {k.definition.kpi_id: k.value for k in result.kpis}


def test_confirmed_reference_controls_existing_range_policy_without_changing_raw_data():
    data = _battery_profile_dataset([297.78, 221.265], [85, 64.5782], distance_km=10.7397)
    before = data.values.copy()
    result = calculate_profile_statistics(data, _battery_range_profile(), reporting_inputs=confirmed())
    k = values(result)
    assert k['battery_capacity_100'] == 350
    assert k['usable_battery_capacity_85'] == 297.5
    assert k['battery_capacity_used'] == pytest.approx(76.515)
    assert k['battery_energy_consumption_wh_per_km'] == pytest.approx(7124.500684376657)
    assert k['range_85_battery_km'] == pytest.approx(41.757312291707514)
    assert result.battery_capacity['reference_status'] == 'confirmed'
    np.testing.assert_array_equal(data.values, before)


def test_unknown_nominal_capacity_is_explicitly_inferred_and_preserves_legacy_range():
    data = _battery_profile_dataset([40, 35], [80, 70], distance_km=1)
    result = calculate_profile_statistics(data, _battery_range_profile())
    assert result.battery_capacity['nominal_capacity_kwh'] is None
    assert result.battery_capacity['reference_status'] == 'inferred'
    assert values(result)['range_85_battery_km'] == pytest.approx(8.5)
    assert not result.diagnostics


def test_nonlinear_energy_soc_is_qualified_not_claimed_as_nominal_capacity_change():
    data = _battery_profile_dataset([297.78, 255, 221.265], [85, 75, 64.5782], distance_km=1)
    result = calculate_profile_statistics(data, _battery_range_profile(), reporting_inputs=confirmed())
    warning = next(d for d in result.diagnostics if d.code == 'battery_energy_soc_variation')
    assert 'not a confirmed nominal specification' in warning
    assert 'cause is not established' in warning and 'no temperature law is applied' in warning
    assert warning.severity == 'info' and not warning.user_visible
    assert 'diagnostic heuristic' in warning
    assert result.battery_capacity['energy_soc_equivalent_kwh']['tolerance'] == pytest.approx(
        result.battery_capacity['energy_soc_equivalent_kwh']['median']*.01)
    assert values(result)['battery_capacity_100'] == 350
    assert 'temperature' in result.battery_capacity['nominal_capacity_provenance']


def test_small_initial_difference_is_recorded_without_a_false_mismatch_warning():
    result = calculate_profile_statistics(_battery_profile_dataset([297.78, 262.5], [85, 75]),
        _battery_range_profile(), reporting_inputs=confirmed())
    assert result.battery_capacity['initial_energy_minus_nominal_times_soc_kwh'] == pytest.approx(.28)
    assert all(d.severity == 'info' for d in result.diagnostics)


def test_invalid_source_pairs_are_flagged_and_not_used_for_inference():
    result = calculate_profile_statistics(_battery_profile_dataset([40, -1, 30], [80, 110, 60]),
        _battery_range_profile())
    assert values(result)['battery_capacity_100'] == 50
    assert any(d.code == 'battery_source_data' for d in result.diagnostics)
    assert result.battery_capacity['energy_soc_equivalent_kwh']['excluded_samples'] == 1
    np.testing.assert_array_equal(nominal_capacity_estimates_kwh([40, -1, 30], [80, 110, 60]), [50, 50])


def test_initial_energy_soc_comparison_is_technical_without_a_validated_physical_limit():
    result = calculate_profile_statistics(_battery_profile_dataset([170, 150], [85, 75]),
        _battery_range_profile(), reporting_inputs=confirmed())
    assert any(d.code == 'battery_initial_reference_mismatch' and not d.user_visible
               and d.severity == 'info' for d in result.diagnostics)
    assert values(result)['battery_capacity_100'] == 350


@pytest.mark.parametrize('inputs', [ReportingInputs(), confirmed()])
def test_ratio_deviation_alone_is_not_a_kpi_reliability_test(inputs):
    data = _battery_profile_dataset([40, 30, 20], [80, 40, 20], distance_km=1)
    result = calculate_profile_statistics(data, _battery_range_profile(), reporting_inputs=inputs)
    assert result.is_complete
    assert not any(d.user_visible for d in result.diagnostics)
    estimate = result.battery_capacity['energy_soc_equivalent_kwh']
    assert estimate['max_deviation'] > estimate['tolerance']
    assert 'not a physical acceptance limit' in estimate['tolerance_basis']
    assert values(result)['battery_capacity_100'] == (350 if inputs.battery_nominal_capacity_kwh else 75)


def test_unusable_energy_soc_reference_keeps_required_range_unavailable():
    with pytest.raises(StatisticsError, match='battery_capacity_100.*no finite nonnegative energy'):
        calculate_profile_statistics(_battery_profile_dataset([40, 30], [0, 0]),
                                     _battery_range_profile())


@pytest.mark.parametrize('value', [0, -1, True, float('nan'), float('inf'), '350'])
def test_nominal_capacity_requires_positive_finite_numeric_input(value):
    with pytest.raises(ConfigurationError):
        replace(confirmed(), battery_nominal_capacity_kwh=value)


@pytest.mark.parametrize('patch', [dict(battery_nominal_capacity_provenance=None),
    dict(battery_nominal_capacity_status='not_provided'), dict(battery_nominal_capacity_kwh=None)])
def test_nominal_capacity_requires_consistent_status_and_provenance(patch):
    with pytest.raises(ConfigurationError):
        replace(confirmed(), **patch)


def test_simulation_confirmation_requires_exact_content_hash(tmp_path):
    source = tmp_path/'RoboSprayer_8500Kg_Electric_03.csv'
    source.write_text('Synthetic source for configuration matching only')
    profile = profile_for('electric')
    rule = next(c for c in profile.metadata.machine_configurations if c.get('source_sha256'))
    rule = dict(rule, source_sha256=hashlib.sha256(source.read_bytes()).hexdigest())
    profile = replace(profile, metadata=replace(profile.metadata, machine_configurations=(rule,)))
    assert resolve_report_metadata(source, profile).reporting_inputs.battery_nominal_capacity_kwh == 350
    source.write_text('Different simulation with the same filename')
    assert resolve_report_metadata(source, profile).reporting_inputs.battery_nominal_capacity_kwh is None
    with pytest.raises(ConfigurationError):
        resolve_report_metadata(source, profile, reporting_inputs={'machine_configuration_id': rule['configuration_id']})


def test_scoped_confirmation_is_checked_against_the_imported_dataset():
    data = _battery_profile_dataset([40, 35], [80, 70])
    profile = _battery_range_profile()
    rule = next(c for c in profile_for('electric').metadata.machine_configurations if c.get('source_sha256'))
    profile = replace(profile, metadata=replace(profile.metadata, machine_configurations=(rule,)))
    inputs = replace(confirmed(), machine_configuration_id=rule['configuration_id'])
    with pytest.raises(StatisticsError, match='imported source SHA-256'):
        calculate_profile_statistics(data, profile, reporting_inputs=inputs)


@pytest.mark.parametrize('kind', ['electric', 'hybrid'])
def test_no_universal_nominal_capacity_for_other_configurations(kind):
    metadata = resolve_report_metadata('RoboSprayer_3500Kg_'+kind+'_Other.csv', profile_for(kind))
    assert metadata.reporting_inputs.battery_nominal_capacity_kwh is None


@pytest.fixture(scope='module')
def exported_reference(tmp_path_factory):
    folder = tmp_path_factory.mktemp('capacity_report')
    data = dataset_for('electric')
    data.values[:, data.channel_index('electricsystem_battery_energy')] = [297.78, 270, 240, 221.265]
    data.values[:, data.channel_index('electricsystem_battery_soc')] = [85, 78, 70, 64.5782]
    from dataclasses import asdict
    report = generate_profile_excel_report(csv_for(data, folder/'source.csv'), profile_path('electric'),
        folder/'excel', reporting_inputs=asdict(confirmed()))
    ppt = build_profile_powerpoint_report(report, folder/'powerpoint')
    return report, ppt


def test_excel_and_powerpoint_share_the_confirmed_reference_and_range(exported_reference):
    report, ppt = exported_reference
    expected = values(report.statistics_result)
    book = load_workbook(report.report_path, data_only=True)
    try:
        kpis = {r[0]: r[2] for r in book['Statistics'].values}
        assert kpis['battery_capacity_100'] == 350
        assert kpis['range_85_battery_km'] == pytest.approx(expected['range_85_battery_km'])
        assert 'Motor Curves' not in book and len(book['Torque Maps']._charts) == 2
        metadata = {r[0]: r[1] for r in book['Metadata'].values}
        assert json.loads(metadata['Battery capacity interpretation'])['reference_status'] == 'confirmed'
    finally:
        book.close()
    prs = Presentation(ppt.presentation_path)
    assert len(prs.slides) == 16
    notes = [json.loads(s.notes_slide.notes_text_frame.text) for s in prs.slides]
    assert any('350' in json.dumps(n) and 'battery_capacity_reference' in json.dumps(n) for n in notes)
    # The PowerPoint adapter receives exactly the shared canonical KPI value.
    actual = {s.statistic_id: s.value for s in ppt.powerpoint_result.statistics_result.statistics}
    assert actual['range_85_battery_km'] == pytest.approx(expected['range_85_battery_km'])


def test_excel_formula_parameter_retains_confirmation_provenance(exported_reference):
    report, _ = exported_reference
    book = load_workbook(report.report_path)
    try:
        row = next(r for r in book['Statistics'] if r[0].value == 'battery_capacity_100')
        assert row[2].data_type == 'f' and 'Formula Support' in row[2].value
        assert 'Sergio' in row[2].comment.text
        parameter = next(r for r in book['Formula Support'] if 'Nominal reference: confirmed' in str(r[0].value))
        assert parameter[1].value == 350 and '2 November 2026' in parameter[0].value
    finally:
        book.close()


def test_streamlit_and_exported_diagnostics_share_messages(exported_reference):
    from vsm_postprocessing.ui_app import _render_diagnostics
    report, _ = exported_reference
    class Display:
        def __init__(self): self.messages = []
        def warning(self, message): self.messages.append(('warning', message))
        def info(self, message): self.messages.append(('info', message))
        def error(self, message): self.messages.append(('error', message))
    display = Display()
    _render_diagnostics(display, report.statistics_result.diagnostics)
    expected = [(d['severity'], d['message']) for d in diagnostic_records(report.statistics_result.diagnostics) if d['user_visible']]
    assert display.messages == expected


def test_streamlit_results_keep_battery_provenance_on_demand_and_real_warnings_visible(exported_reference):
    from streamlit.testing.v1 import AppTest
    from types import SimpleNamespace
    from vsm_postprocessing.diagnostics import Diagnostic
    report, _ = exported_reference
    result = SimpleNamespace(excel_result=report, report_metadata=report.report_metadata,
        **{key: getattr(report, key) for key in ('sample_count', 'report_channel_count',
            'plot_count', 'math_count', 'statistic_count', 'kpi_count', 'report_path')})
    app = AppTest.from_string('''
import streamlit as st
from vsm_postprocessing.ui_app import _render_profile_report_completion
st.success('Engineering reports generated')
_render_profile_report_completion(st, st.session_state['result'])
''')
    app.session_state['result'] = result
    app.run()
    assert not app.exception and app.success
    assert not any('Battery range reference' in m.value or 'Energy/SOC' in m.value for m in app.info)
    assert not any('Energy/SOC' in m.value for m in app.warning)
    assert any(e.label == 'Diagnostics' and not e.proto.expanded for e in app.expander)
    assert 'battery_capacity_reference' in app.json[0].value
    assert 'Sergio' in app.json[0].value and 'source_sha256' in app.json[0].value
    problem = Diagnostic('Battery source data require review: invalid Energy/SOC pair', code='battery_source_data')
    report.statistics_result.diagnostics.append(problem)
    try:
        app.run()
        assert not app.exception
        assert any(m.value == problem for m in app.warning)
    finally:
        report.statistics_result.diagnostics.remove(problem)
