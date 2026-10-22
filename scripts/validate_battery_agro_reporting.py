"""Generate reviewable local samples and inspect XLSX/PPTX structure and values.

Run with .venv/Scripts/python.exe scripts/validate_battery_agro_reporting.py.
Does not commit, push or produce a release archive.
"""
from dataclasses import asdict
import json
from pathlib import Path
import subprocess
import sys
import types

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation

from vsm_postprocessing.agro_reporting import calculate_reporting_kpis
from vsm_postprocessing.importer import ImportOptions
from vsm_postprocessing.profile_powerpoint_report_engine import build_profile_powerpoint_report
from vsm_postprocessing.excel_report_engine import generate_profile_excel_report


ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / 'outputs' / 'battery_agro_validation'
ELECTRIC = ROOT / 'reference_files/RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Mot_63RPM_Susp_Cool_Rough_Grad_Discharge.csv'
HYBRID = ROOT / 'reference_files/Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem.csv'
ZERO = ROOT / 'reference_files/RoboSprayer_3500Kg_Electric_12kph_Batt_50kW_Motor_63RPM_Susp_Cool_Rough_Crop_Field_05.csv'
DIESEL = ROOT / 'reference_files/Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient.csv'


def _baseline_module():
    source = subprocess.run(['git', 'show', 'HEAD:src/vsm_postprocessing/profile_statistics.py'],
                            cwd=ROOT, text=True, capture_output=True, check=True).stdout
    module = types.ModuleType('vsm_postprocessing._battery_agro_baseline')
    module.__package__ = 'vsm_postprocessing'
    sys.modules[module.__name__] = module
    exec(compile(source, '<HEAD profile_statistics.py>', 'exec'), module.__dict__)
    return module


def inspect(excel, ppt, expected_slides):
    prs = Presentation(ppt.presentation_path)
    assert len(prs.slides) == expected_slides
    page = prs.slides[2 if excel.profile.metadata.powertrain == 'diesel' else 7]
    texts = [s.text for s in page.shapes if s.has_text_frame]
    assert 'AGRO CAPACITY' in texts and 'AGRO DISCHARGE RATE' in texts
    assert 'AGRO DISCHARGE' not in texts
    # Native shapes, with separate label/value frames inside styled card backgrounds.
    cards = [s for s in page.shapes if s.shape_type == 1 and s.top > .9 * 914400
             and s.top < 3 * 914400 and s.width > 1 * 914400 and s.height > .6 * 914400]
    assert len(cards) >= 2
    for i, first in enumerate(cards):
        assert first.left >= 0 and first.top >= 0
        assert first.left + first.width <= prs.slide_width
        for second in cards[i + 1:]:
            overlap_x = min(first.left + first.width, second.left + second.width) - max(first.left, second.left)
            overlap_y = min(first.top + first.height, second.top + second.height) - max(first.top, second.top)
            assert overlap_x <= 0 or overlap_y <= 0
    formula = load_workbook(excel.report_path)
    cached = load_workbook(excel.report_path, data_only=True)
    try:
        summary = formula.worksheets[0]
        labels = [summary.cell(3, c).value for c in range(excel.report_channel_count + 2, summary.max_column + 1)]
        assert any('Agro Capacity' in str(v) for v in labels)
        assert any('Agro Discharge Rate' in str(v) for v in labels)
        assert len(summary._charts) > 0
        start = len(excel.statistics_result.statistics) + len(excel.statistics_result.kpis) + 8
        rows = list(cached['Statistics'].iter_rows(min_row=start, values_only=True))
        assert any(row[0] == 'agro_discharge_rate_L_per_m' and row[3] == 'L/m' for row in rows)
        if excel.profile.metadata.powertrain in {'electric', 'hybrid'}:
            assert any(row[0] == 'battery_mass_kg' for row in rows)
            battery_texts = [s.text for s in prs.slides[4].shapes if s.has_text_frame]
            assert 'BATTERY PACK MASS: ' + excel.report_metadata.reporting_inputs.battery_display in battery_texts
        # All top summary cells, including text/N/A values, remain live references
        # and have readable initial caches matching their target Statistics cells.
        for col in range(excel.report_channel_count + 2, summary.max_column + 1):
            cell = summary.cell(4, col)
            if cell.value is None:
                continue
            assert cell.data_type == 'f'
            target = cell.value[1:].split('!')
            sheet = cached[target[0].strip("'")] if len(target) == 2 else cached.worksheets[0]
            actual, expected = cached.worksheets[0][cell.coordinate].value, sheet[target[-1]].value
            if isinstance(expected, (int, float)):
                np.testing.assert_allclose(actual, expected, rtol=2e-15, atol=1e-14)
            else:
                assert actual == expected
        return {'slide_count': len(prs.slides), 'excel_chart_count': len(summary._charts),
                'editable_agro_card_count': len(cards), 'summary_formula_caches': 'PASS',
                'agro_card_geometry': 'PASS'}
    finally:
        formula.close()
        cached.close()


def main():
    DEST.mkdir(parents=True, exist_ok=True)
    baseline = _baseline_module()
    audit = []
    scenarios = [
        ('electric_confirmed', ELECTRIC, 'robosprayer_electric', False, 12,
         {'machine_configuration_id': 'robosprayer_v5_3500kg', 'agro_density_kg_per_L': 1.0,
          'density_provenance': 'HYPOTHETICAL VALIDATION ONLY; not approved chemical density'}),
        ('hybrid_invalid_load', HYBRID, 'robosprayer_hybrid', False, 12, {}),
        ('electric_zero_load', ZERO, 'robosprayer_electric', False, 12, {}),
        ('october_electric', ELECTRIC, 'robosprayer_electric', True, 15, {}),
        ('october_hybrid', HYBRID, 'robosprayer_hybrid', True, 14, {}),
        ('full_size_diesel', DIESEL, 'full_size_sprayer_diesel', False, 10, {}),
    ]
    for name, source, profile_id, october, slides, settings in scenarios:
        print('Generating ' + name, flush=True)
        path = ROOT / 'config/report_profiles'
        if october:
            path /= 'sergio_2026_10'
        profile = path / (profile_id + '.yaml')
        settings = {'hitch_force_is_remaining_agro_mass': True, **settings}
        excel = generate_profile_excel_report(source, profile, DEST / name / 'excel',
                                             ImportOptions(strict=not october), reporting_inputs=settings)
        # Original HEAD statistics/registry plus exact math-array comparison.
        previous = baseline.calculate_profile_statistics(excel.dataset, excel.profile)
        assert {s.definition.statistic_id: s.value for s in previous.statistics} == {s.definition.statistic_id: s.value for s in excel.statistics_result.statistics}
        assert {k.definition.kpi_id: k.value for k in previous.kpis} == {k.definition.kpi_id: k.value for k in excel.statistics_result.kpis}
        for key, values in previous.math_result.values_by_semantic_name.items():
            np.testing.assert_array_equal(values, excel.math_result.values_by_semantic_name[key])
        old_metrics = {m.metric_id: m.value for m in baseline.build_canonical_report_metrics(previous)}
        assert old_metrics == {m.metric_id: m.value for m in excel.statistics_result.canonical_metrics}
        ppt = build_profile_powerpoint_report(excel, DEST / name / 'powerpoint')
        details = inspect(excel, ppt, slides)
        agro = excel.statistics_result.reporting_agro
        if name == 'electric_confirmed':
            assert abs(agro.rate.value - .499832) < 1e-6
            # A second reviewable report with the identical data and no density.
            from dataclasses import replace
            inputs = replace(excel.report_metadata.reporting_inputs, density_kg_per_L=None, density_provenance=None)
            metadata = replace(excel.report_metadata, reporting_inputs=inputs)
            no_density = calculate_reporting_kpis(excel.statistics_result, metadata)
            assert no_density.rate.value is None and 'density required' in no_density.rate.reason
        if 'hybrid' in name:
            assert agro.rate.value is None and 'negative' in agro.rate.reason
        if name == 'electric_zero_load':
            assert agro.capacity.value is None and agro.rate.value is None
        audit.append({'scenario': name, 'excel': str(excel.report_path), 'powerpoint': str(ppt.presentation_path),
                      'unchanged_legacy_arrays_statistics_kpis': 'PASS', **details,
                      'reporting_inputs': asdict(excel.report_metadata.reporting_inputs), 'agro': agro.to_dict()})
        (DEST / 'audit.json').write_text(json.dumps(audit, indent=2, allow_nan=False), encoding='utf-8')
        print(name + ' PASS', flush=True)


if __name__ == '__main__':
    main()
