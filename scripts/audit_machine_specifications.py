"""Reproduce the machine audit and retain reviewable reports, without Office.

Run with the repository Python from its root. Original reports are read only.
"""
from pathlib import Path
import json
import math

from openpyxl import load_workbook
from pptx import Presentation

from vsm_postprocessing.importer import ImportOptions
from vsm_postprocessing.ui_config import generate_reporting_profile_engineering_report
from vsm_postprocessing.utils import sha256_file
from vsm_postprocessing.report_selection import ProfileReportSelection


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / 'outputs/machine_audit_2026_10_23'
CASES = [
    ('electric3500', 'robosprayer_electric', 'RoboSprayer_3500Kg_Electric_UoL_Drive_Cycle_12kph_DriverPC_12-5pcGrad_01',
     '20261023_202309_023085', {'battery_mass_kg': 450}),
    ('electric8500', 'robosprayer_electric', 'RoboSprayer_8500Kg_Electric_02_UoL_Drive_Cycle_20kph_DriverPC_Flat_02',
     '20261023_202617_048134', {'battery_mass_kg': 1000}),
    ('hybrid', 'robosprayer_hybrid', 'Sprayer_Caiman_SP_9300Kg_Hybrid_12x1Km_4000Kg_Chem', None, {}),
    ('diesel', 'full_size_sprayer_diesel', 'Full_Size_Sprayer_8500Kg_Diesel_Test_03_Crop_Field_MatLab_20kph_Gradient', None, {}),
]


def workbook_metrics(path, result):
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        rows = {row[0]: row for row in wb['Statistics'].iter_rows(values_only=True) if row[0]}
        values = {s.definition.statistic_id: rows[s.definition.statistic_id][4]
                  for s in result.statistics_result.statistics}
        values.update({k.definition.kpi_id: rows[k.definition.kpi_id][2] for k in result.statistics_result.kpis})
        return values
    finally:
        wb.close()


def main():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    audit = []
    for name, profile_id, stem, run, inputs in CASES:
        source = ROOT / 'reference_files' / (stem + '.csv')
        profile_file = ROOT / 'config/report_profiles/sergio_2026_10' / (profile_id + '.yaml')
        manifest = None
        selection = None
        if run:
            old = ROOT / 'outputs/ui_runs' / run
            manifest = json.loads((old / 'profile_excel_report/profile_excel_report_manifest.json').read_text())
            assert manifest['source_sha256'] == sha256_file(source)
            assert manifest['profile_sha256'] == sha256_file(profile_file)
            selection = ProfileReportSelection(**manifest['selection']) if manifest['selection'] else None
        result = generate_reporting_profile_engineering_report(source,
            profile_file, OUTPUT / name,
            ImportOptions(strict=False), reporting_inputs=inputs, selection=selection)
        prs = Presentation(result.presentation_path)
        metrics = workbook_metrics(result.report_path, result.excel_result)
        entry = {'case': name, 'profile': profile_id, 'source': str(source),
                 'source_sha256': sha256_file(source), 'samples': result.sample_count,
                 'excel': str(result.report_path), 'presentation': str(result.presentation_path),
                 'slide_ids': [s.slide_id for s in result.powerpoint_result.powerpoint_result.config.slides],
                 'slide_count': len(prs.slides), 'metadata': result.report_metadata.__dict__.copy(),
                 'agro': result.excel_result.statistics_result.reporting_agro.to_dict(),
                 'canonical_workbook_values': metrics,
                 'profile_file': str(profile_file),
                 'cover_pills': [prs.slides[0].shapes[i].text for i in (6, 8, 10, 12, 14)]}
        entry['metadata']['reporting_inputs'] = result.report_metadata.reporting_inputs.__dict__
        if run:
            old = ROOT / 'outputs/ui_runs' / run
            old_excel = old / 'profile_excel_report' / (stem + '.xlsx')
            old_ppt = old / 'profile_powerpoint_report' / (stem + '.pptx')
            old_values = workbook_metrics(old_excel, result.excel_result)
            differences = {key: (old_values[key], value) for key, value in metrics.items()
                           if not math.isclose(old_values[key], value, rel_tol=1e-12, abs_tol=1e-12)}
            assert not differences, 'Pre-existing numeric statistic or KPI differences: ' + repr(differences)
            entry['original_run'] = run
            entry['original_excel'] = str(old_excel)
            entry['original_presentation'] = str(old_ppt)
            entry['original_excel_sha256'] = sha256_file(old_excel)
            entry['original_presentation_sha256'] = sha256_file(old_ppt)
            entry['all_original_statistics_and_kpis_unchanged'] = True
            entry['excel_cache_comparison_tolerance'] = {'relative': 1e-12, 'absolute': 1e-12}
            original = Presentation(old_ppt)
            entry['original_cover_pills'] = [original.slides[0].shapes[i].text for i in (6, 8, 10, 12, 14)]
        audit.append(entry)
        (OUTPUT / 'audit.json').write_text(json.dumps(audit, indent=2, allow_nan=False), encoding='utf-8')
        print(name + ': report generation and numerical checks PASS', flush=True)


if __name__ == '__main__':
    main()
