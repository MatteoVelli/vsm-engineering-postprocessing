"""Exercise the real Electric upload through Streamlit and audit raw CSV results.

Usage: python scripts/validate_electric_uol.py SOURCE --output-dir NEW_DIRECTORY
The private source is read only; existing evidence directories are refused.
"""
import argparse
import csv
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation
import streamlit as st
from streamlit.testing.v1 import AppTest

from vsm_postprocessing import ui_app
from vsm_postprocessing.report_profile import load_reporting_profile


def audit(result, source, selection, destination):
    # Parse the raw export independently of the production importer/math engine.
    with source.open(encoding='utf-8-sig', newline='') as stream:
        reader = csv.reader(stream)
        names, units = next(reader), next(reader)
        matrix = np.array([[float(v) for v in row] for row in reader])
    raw = dict(zip(names, matrix.T))
    assert matrix.shape == (3954, 608)
    report = result.excel_result
    values = report.plotting_result.values_by_semantic_name
    checks = {}

    def check(name, expected):
        actual = values[name]
        np.testing.assert_allclose(actual, expected, rtol=2e-14, atol=1e-12)
        checks[name] = {'samples': len(actual), 'max_absolute_error': float(np.max(np.abs(actual - expected))),
                        'minimum': float(np.min(actual)), 'maximum': float(np.max(actual)), 'last': float(actual[-1])}
        return expected

    expected = {c.semantic_name: raw[c.source_name] for c in report.profile.raw_channels
                if c.source_name in raw and c.semantic_name in values}
    time = raw['Track_Time']
    torque, power = np.zeros(len(time)), np.zeros(len(time))
    for i, wheel in enumerate(('fl', 'fr', 'rl', 'rr'), 1):
        corner_torque = check('driveshaft_torque_' + wheel, raw[f'KERS_MotGenUnit_GearboxOutputTorque_{i}'])
        speed = raw[f'Wheel_RotationalSpeed_{wheel.upper()}']
        corner_power = check('wheel_power_' + wheel, corner_torque * speed * (2 * np.pi / 60000))
        expected['wheel_power_' + wheel] = corner_power
        torque += corner_torque
        power += corner_power
    expected['wheel_total_torque'] = check('wheel_total_torque', torque)
    expected['wheel_power_total'] = check('wheel_power_total', power)
    resistance, interval = np.zeros(len(time)), np.zeros(len(time))
    for component in ('rolling', 'bulldozing', 'compaction'):
        total = sum(raw[f'Tyre_{component.capitalize()}ResistancePower_{w}'] for w in ('FL', 'FR', 'RL', 'RR'))
        expected[f'total_{component}_resistance_power'] = check(f'total_{component}_resistance_power', total)
        increments = np.r_[0., (total[:-1] + total[1:]) * np.diff(time) / 7200]
        for suffix, vector in (('kwh', increments), ('wh', increments * 1000), ('accumulated', np.cumsum(increments))):
            name = f'tyre_{component}_resistance_energy_{suffix}'
            expected[name] = check(name, vector)
        resistance += total
        interval += increments
    expected['tyre_total_resistance_power'] = check('tyre_total_resistance_power', resistance)
    for suffix, vector in (('kwh', interval), ('wh', interval * 1000), ('accumulated', np.cumsum(interval))):
        name = 'tyre_total_resistance_energy_' + suffix
        expected[name] = check(name, vector)
    for name, vector in expected.items():
        np.testing.assert_allclose(values[name], vector, rtol=2e-14, atol=1e-12)

    hv = 'auxiliary_high_voltage_power_consumption'
    hv_source = 'User_defined_Data_High_Voltage_Auxiliary_signal1_0'
    assert units[names.index(hv_source)] == '-'
    resolved_hv = report.resolution.resolved[hv]
    assert resolved_hv.channel.unit == '-' and resolved_hv.effective_unit == 'kW'
    assert resolved_hv.unit_interpretation.authority.startswith('Sergio, Astauto, 16 October 2026')
    assert not any('unit mismatch' in d for d in report.statistics_result.diagnostics)
    expected[hv] = check(hv, raw[hv_source])
    low = expected['electricsystem_auxiliary_lowvoltagepowerconsumption']
    total = expected[hv] + low
    expected['total_auxiliary_power'] = check('total_auxiliary_power', total)
    increments = np.r_[0., (total[:-1] + total[1:]) * np.diff(time) / 7200]
    expected['auxiliary_energy_consumption'] = check('auxiliary_energy_consumption', increments)
    expected['auxiliary_energy_consumption_accumulated'] = check('auxiliary_energy_consumption_accumulated', np.cumsum(increments))
    expected_exports = set(selection.export_channels) & set(values)
    assert {c.channel_id for c in report.report_channels} == expected_exports
    # Both previously colliding channels must be retained, with their own raw data.
    for pair in (('suspension_damper_force_fl', 'suspension_damper_force_fr'),
                 ('suspension_damper_travel_fl', 'suspension_damper_travel_fr'),
                 ('tyre_rollingresistancepower_fr', 'tyre_rollingresistancepower_rr'),
                 ('softsoil_sheardisplacementy_in_fl', 'softsoil_sheardisplacementy_in_rr')):
        assert set(pair) <= expected_exports

    operations = {'max': np.max, 'min': np.min, 'first': lambda v: v[0], 'last': lambda v: v[-1],
                  'rms': lambda v: np.sqrt(np.mean(v ** 2)), 'positive_max': lambda v: max(0., np.max(v)),
                  'absolute_max': lambda v: np.max(np.abs(v))}
    statistics = []
    for statistic in report.statistics_result.statistics:
        definition = statistic.definition
        if definition.target in expected:
            independent = float(operations[definition.operation](expected[definition.target]))
            np.testing.assert_allclose(statistic.value, independent, rtol=2e-14, atol=1e-12)
            statistics.append(definition.statistic_id)

    # Streaming readers avoid holding multiple million-cell workbook copies.
    cached = load_workbook(report.report_path, read_only=True, data_only=True)
    formulas = load_workbook(report.report_path, read_only=True, data_only=False)
    live = set()
    formula_count = numeric_count = 0
    try:
        channels = report.report_channels
        sheet = cached.worksheets[0]
        for index, row in enumerate(sheet.iter_rows(min_row=5, max_row=len(time) + 4, max_col=len(channels), values_only=True)):
            target = np.array([values[c.channel_id][index] for c in channels])
            np.testing.assert_allclose(np.asarray(row, dtype=float), target, rtol=2e-14, atol=1e-12)
            numeric_count += len(row)
        for row in formulas.worksheets[0].iter_rows(min_row=5, max_row=len(time) + 4, max_col=len(channels)):
            for channel, cell in zip(channels, row):
                if cell.data_type == 'f':
                    live.add(channel.channel_id)
                    formula_count += 1
        assert {'wheel_power_total', 'tyre_total_resistance_energy_accumulated',
                'total_auxiliary_power', 'auxiliary_energy_consumption', 'auxiliary_energy_consumption_accumulated'} <= live
        assert 'Statistics' in cached.sheetnames and 'Rename From VSM to Astauto' in cached.sheetnames
    finally:
        cached.close()
        formulas.close()
    deck = Presentation(result.presentation_path)
    template = Presentation(report.profile.presentation.powerpoint_template)
    assert len(deck.slides) == len(report.profile.presentation.slides)
    assert (deck.slide_width, deck.slide_height) == (template.slide_width, template.slide_height)
    assert not any('PLOT DATA NOT AVAILABLE' in s.text for slide in deck.slides for s in slide.shapes if s.has_text_frame)
    assert 'auxiliaries_energy_consumption' in {p.plot_id for p in report.plotting_result.rendered_plots}
    audit_entry = {'profile': 'electric_uol', 'streamlit_version': st.__version__, 'source': str(source),
                   'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                   'workbook': str(report.report_path), 'presentation': str(result.presentation_path),
                   'samples': len(time), 'source_channels': len(names), 'exported_channels': len(expected_exports),
                   'live_math': sorted(live), 'equation_inventory': list(report.math_result.calculation_order),
                   'numeric_cells_compared': numeric_count, 'data_formulas': formula_count,
                   'independent_checks': checks, 'independent_statistics': statistics,
                   'plots': [p.plot_id for p in report.plotting_result.rendered_plots], 'slides': len(deck.slides),
                   'diagnostics': list(report.statistics_result.diagnostics),
                   'hv_source_unit': units[names.index(hv_source)],
                   'hv_effective_unit': resolved_hv.effective_unit,
                   'hv_unit_interpretation': asdict(resolved_hv.unit_interpretation),
                   'hv_source_range_kw': [float(np.min(raw[hv_source])), float(np.max(raw[hv_source]))],
                   'slide_7_auxiliary_plot': 'PASS: restored from actual HV + LV measurements; no placeholder'}
    (destination / 'office_audit.json').write_text(json.dumps([audit_entry], indent=2), encoding='utf-8')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--output-dir', type=Path, required=True)
    args = parser.parse_args()
    source, destination = args.source.resolve(), args.output_dir.resolve()
    destination.mkdir(parents=True, exist_ok=False)
    results, selections = [], []
    original_options, original_generate = ui_app._render_profile_report_options, ui_app.generate_reporting_profile_engineering_report

    def options(*a, **kw):
        result = original_options(*a, **kw)
        selections.append(result[1])
        return result

    def generate(*a, **kw):
        result = original_generate(*a, **kw)
        results.append(result)
        return result

    uploaded = SimpleNamespace(name=source.name, getvalue=source.read_bytes)
    with patch.object(st, 'file_uploader', return_value=uploaded), \
         patch.object(ui_app, 'UI_WORKSPACE', destination / 'workspace'), \
         patch.object(ui_app, 'UI_RUNS', destination / 'ui_runs'), \
         patch.object(ui_app, '_render_profile_report_options', options), \
         patch.object(ui_app, 'generate_reporting_profile_engineering_report', generate):
        at = AppTest.from_string('from vsm_postprocessing.ui_app import main\nmain()', default_timeout=600).run()
        at.selectbox[0].set_value('robosprayer_electric').run()
        initial = selections[-1]
        at.button[0].click().run()
        assert not at.exception and not at.error
        assert at.session_state['profile_validation_summary'].is_valid
        assert selections[-1] == initial and not at.button[1].disabled
        summary = asdict(at.session_state['profile_validation_summary'])
        at.button[1].click().run()
        assert not at.exception and not at.error, [e.value for e in at.error]
        assert results and selections[-1] == initial
    (destination / 'ui_evidence.json').write_text(json.dumps({'validation': summary, 'selection': asdict(initial)}, indent=2), encoding='utf-8')
    print('Streamlit validation and Excel/PowerPoint generation passed', flush=True)
    audit(results[-1], source, initial, destination)
    print(destination / 'office_audit.json', flush=True)


if __name__ == '__main__':
    main()
