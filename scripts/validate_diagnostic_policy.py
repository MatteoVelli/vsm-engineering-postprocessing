"""Real-source policy audit; snapshot numbers before applying presentation changes."""
from __future__ import annotations

import argparse
import csv
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import sys
from zipfile import ZipFile

import numpy as np
from openpyxl import load_workbook
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from vsm_postprocessing.importer import load_data_file
from vsm_postprocessing.profile_statistics import calculate_profile_statistics
from vsm_postprocessing.report_profile import load_reporting_profile
from vsm_postprocessing.report_metadata import resolve_report_metadata
from vsm_postprocessing.diagnostics import diagnostic_records
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report, validate_reporting_profile_source
from vsm_postprocessing.errors import VSMPostProcessingError

ROOT = Path('outputs/diagnostic_policy_2026_11_02')
PROFILE = Path('config/report_profiles/sergio_2026_10/robosprayer_electric.yaml')
KNOWN = Path.home() / 'Downloads/RoboSprayer_8500Kg_Electric_03_UoL_DC_20kph_DriverPC_25pcGrad_FL-RL_Rough.csv'
OTHER = Path('reference_files/RoboSprayer_3500Kg_Electric_UoL_DriveCycle_12kph_DriverPC_20pcGrad_04.csv')
EXAMPLE = Path('reference_files/RoboSprayer_3500Kg_Electric_UoL_Drive_Cycle_12kph_DriverPC_12-5pcGrad_01.csv')
UNSEEN = ROOT / 'inputs/Unregistered_Electric_Prototype.csv'


def sha(data):
    return hashlib.sha256(data).hexdigest()


def numerics(data, stats):
    return {'source_sha256': sha(data.source_path.read_bytes()), 'raw_sha256': sha(data.values.tobytes()),
            'math': {k: sha(v.tobytes()) for k, v in stats.math_result.values_by_semantic_name.items()},
            'statistics': {s.definition.statistic_id: s.value for s in stats.statistics},
            'kpis': {k.definition.kpi_id: k.value for k in stats.kpis}}


def options(index):
    if index:
        return {}
    old = json.loads(Path('outputs/bcapvalid/validation.json').read_text())[KNOWN.name]
    book = load_workbook(old['excel'], read_only=True, data_only=True)
    fields = {r[0]: r[1] for r in book['Metadata'].values}
    book.close()
    return {'reporting_inputs': json.loads(fields['Reporting inputs']),
            'machine_name_override': fields['Machine / Vehicle name']}


def write_csv(data, path):
    with path.open('w', newline='', encoding='utf-8') as stream:
        writer = csv.writer(stream)
        writer.writerow([c.source_name for c in data.channels])
        writer.writerow([c.unit or '' for c in data.channels])
        writer.writerows(data.values)


def snapshot():
    UNSEEN.parent.mkdir(exist_ok=False)
    data = load_data_file(EXAMPLE)
    energy = next(i for i, c in enumerate(data.channels) if c.source_name == 'ElectricSystem_Battery_Energy')
    data.values[len(data.values)//2, energy] += .0001
    write_csv(data, UNSEEN)
    records = {}
    profile = load_reporting_profile(PROFILE)
    for index, source in enumerate((KNOWN, OTHER, EXAMPLE, UNSEEN)):
        data = load_data_file(source)
        metadata = resolve_report_metadata(source, profile, **options(index))
        stats = calculate_profile_statistics(data, profile, reporting_inputs=metadata.reporting_inputs)
        records[source.name] = numerics(data, stats)
    (ROOT/'before.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
    protected = json.loads(Path('outputs/battery_capacity_2026_11_02/protected_hashes.json').read_text())
    for old in json.loads(Path('outputs/bcapvalid/validation.json').read_text()).values():
        for key in ('excel', 'powerpoint'):
            path = Path(old[key]); protected[str(path)] = sha(path.read_bytes())
    (ROOT/'protected_hashes.json').write_text(json.dumps(protected, indent=2), encoding='utf-8')
    print('Snapshot: four sources, all raw/math/statistic/KPI values; protected original reports', flush=True)


def preserve_parts(old_record, result):
    with ZipFile(old_record['excel']) as a, ZipFile(result.report_path) as b:
        # Metadata holds the corrected policy; every other worksheet and chart is protected.
        for part in a.namelist():
            if (part.startswith(('xl/charts/', 'xl/drawings/')) and part.endswith('.xml')
                    or part.startswith('xl/worksheets/') and part.endswith('.xml') and part != 'xl/worksheets/sheet3.xml'):
                assert a.read(part) == b.read(part), part
    a, b = Presentation(old_record['powerpoint']), Presentation(result.presentation_path)
    assert len(a.slides) == len(b.slides) == 16
    for index, (x, y) in enumerate(zip(a.slides, b.slides), 1):
        assert x.element.xml == y.element.xml, index
        # PNG encoders can produce different byte streams: compare rendered pixels.
        import io
        from PIL import Image
        left = [Image.open(io.BytesIO(s.image.blob)).convert('RGBA') for s in x.shapes if s.shape_type == 13]
        right = [Image.open(io.BytesIO(s.image.blob)).convert('RGBA') for s in y.shapes if s.shape_type == 13]
        assert len(left) == len(right)
        assert all(p.size == q.size and p.tobytes() == q.tobytes() for p, q in zip(left, right)), index


def validate():
    before = json.loads((ROOT/'before.json').read_text())
    old = json.loads(Path('outputs/bcapvalid/validation.json').read_text())
    records = {}
    for index, source in enumerate((KNOWN, OTHER, EXAMPLE, UNSEEN)):
        print('Generating', source.name, flush=True)
        result = generate_reporting_profile_engineering_report(source, PROFILE, ROOT/f'case_{index}', **options(index))
        report = result.excel_result
        stats = report.statistics_result
        actual = numerics(report.dataset, stats)
        assert actual == before[source.name], 'Engineering numerical change: ' + source.name
        records[source.name] = {'kpis': actual['kpis'], 'capacity_interpretation': stats.battery_capacity,
            'diagnostics': diagnostic_records(stats.diagnostics), 'excel': str(result.report_path),
            'powerpoint': str(result.presentation_path), 'numerics_unchanged': True}
        diagnostics = records[source.name]['diagnostics']
        for d in diagnostics:
            if d['code'] in {'battery_energy_soc_variation', 'battery_capacity_reference',
                             'battery_initial_reference_mismatch', 'machine_identity_unmatched',
                             'optional_battery_mass_missing', 'force_interpretation_unconfirmed'}:
                assert d['severity'] == 'info' and not d['user_visible'], d
        assert any(d['code'] == 'battery_energy_soc_variation' for d in diagnostics)
        assert stats.battery_capacity['reference_status'] == ('confirmed' if index == 0 else 'inferred')
        book = load_workbook(result.report_path, read_only=True, data_only=True)
        fields = {r[0]: r[1] for r in book['Metadata'].values}
        kpis = {r[0]: r[2] for r in book['Statistics'].values if r[0]}
        for key, value in actual['kpis'].items():
            np.testing.assert_allclose(kpis[key], value, rtol=1e-14)
        assert json.loads(fields['Diagnostics']) == diagnostics
        assert json.loads(fields['Battery capacity interpretation']) == stats.battery_capacity
        book.close()
        for slide in Presentation(result.presentation_path).slides:
            notes = json.loads(slide.notes_slide.notes_text_frame.text)
            assert notes['diagnostic_records'] == diagnostics
            assert notes['reporting_inputs'] == asdict(report.report_metadata.reporting_inputs)
        if source.name in old:
            preserve_parts(old[source.name], result)
            records[source.name]['protected_worksheets_charts_slide_xml_and_pixels_equal'] = True
        (ROOT/'validation.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
        print('PASS unchanged numerics / diagnostics / exports:', source.name, flush=True)
    validate_faults()


def validate_faults():
    faults = {}
    for fault in ('missing_required', 'nonnumeric', 'nonfinite', 'invalid_unit', 'unaligned', 'unusable_battery'):
        data = load_data_file(UNSEEN)
        energy = next(i for i, c in enumerate(data.channels) if c.source_name == 'ElectricSystem_Battery_Energy')
        soc = next(i for i, c in enumerate(data.channels) if c.source_name == 'ElectricSystem_Battery_SOC')
        if fault == 'missing_required':
            data.channels.pop(energy); data.values = np.delete(data.values, energy, axis=1)
        elif fault == 'nonfinite': data.values[20, energy] = np.inf
        elif fault == 'unusable_battery': data.values[:, soc] = 0
        elif fault == 'invalid_unit':
            from dataclasses import replace
            data.channels[energy] = replace(data.channels[energy], unit='kg')
        path = UNSEEN.parent/(fault+'.csv'); write_csv(data, path)
        if fault in {'nonnumeric', 'unaligned'}:
            with path.open(newline='', encoding='utf-8') as f: rows = list(csv.reader(f))
            if fault == 'nonnumeric': rows[22][energy] = 'not-a-number'
            else: rows[22] = rows[22][:energy]  # Required samples absent from this row.
            with path.open('w', newline='', encoding='utf-8') as f: csv.writer(f).writerows(rows)
        try:
            summary = validate_reporting_profile_source(path, PROFILE)
            if summary.is_valid:
                generate_reporting_profile_engineering_report(path, PROFILE, ROOT/'invalid'/fault)
            else:
                faults[fault] = {'blocked': True, 'missing': summary.missing_required_names,
                                 'unit_mismatches': summary.unit_mismatch_names}
                continue
        except VSMPostProcessingError as exc:
            faults[fault] = {'blocked': True, 'error': str(exc)}
            continue
        raise AssertionError('Invalid source accepted without blocking: ' + fault)
    (ROOT/'invalid_sources.json').write_text(json.dumps(faults, indent=2), encoding='utf-8')
    for path, expected in json.loads((ROOT/'protected_hashes.json').read_text()).items():
        assert sha(Path(path).read_bytes()) == expected, path
    print('PASS all six invalid files blocked; protected originals unchanged', flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--snapshot', action='store_true')
    parser.add_argument('--faults', action='store_true')
    args = parser.parse_args()
    snapshot() if args.snapshot else validate_faults() if args.faults else validate()
