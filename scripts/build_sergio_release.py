"""Build the established portable Sergio layout with the repository ZIP writer.

Run from the repository root with the validated environment. The final build
requires clean tracked source; --candidate permits pre-commit validation only.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
from zipfile import ZipFile

from vsm_postprocessing.release_builder import _collect_client_files, _zip_write_bytes
from vsm_postprocessing.version import __version__


def build(root: Path, destination: Path, release_date: str, candidate: bool = False):
    root = root.resolve()
    def git(*args):
        return subprocess.check_output(['git', *args], cwd=root, text=True).strip()
    if not candidate and git('status', '--porcelain', '--untracked-files=no'):
        raise RuntimeError('Commit validated source before final packaging.')
    name = f'VSM_Engineering_PostProcessing_v{__version__}_Sergio_Requested_Final'
    # Keep real VSM basenames below Windows path limits in normal extraction paths.
    package_root = f'VSM_Sergio_v{__version__}'
    destination.mkdir(parents=True, exist_ok=True)
    archive_path = destination / (name + '.zip')
    if archive_path.exists():
        raise FileExistsError(f'Refusing to overwrite an existing release: {archive_path}')
    files, excluded = {}, []
    for source in _collect_client_files(root):
        relative = source.relative_to(root).as_posix()
        if ('.egg-info/' in relative or
                relative.startswith('scripts/') and source.name.startswith(
                    ('validate_', 'render_sergio_', 'audit_', 'inspect_', 'build_sergio_')) or
                relative in {
                    'docs/GIT_RELEASE.md', 'docs/SERGIO_2026_10_15_IMPLEMENTATION.md',
                    'docs/TYRE_ROLLING_CORRECTION_2026_10_18.txt',
                    'reference_files/Diesel Template_03.xlsx',
                    'reference_files/Robo_Sprayer_Electrification_Tamplate_Electric_07.xlsx',
                    'reference_files/Robo_Sprayer_Electrification_Tamplate_Hybrid_07.xlsx'}):
            excluded.append(relative)
            continue
        target = relative if relative in ('SETUP_VSM_TOOL.bat', 'START_VSM_TOOL.bat') else 'Tool/' + relative
        files[target] = source.read_bytes()
    files['Tool/outputs/.gitkeep'] = b''
    files['README_SERGIO.txt'] = (root / 'docs/SERGIO_CLIENT_README.md').read_bytes()
    for relative, data in files.items():
        parts = PurePosixPath(relative).parts
        assert not any(p in {'.git', '.venv', '.python', 'setup_cache', '__pycache__',
                            'tests', '.aws', '.codex', '.agents'} for p in parts), relative
        assert not relative.lower().endswith(('.csv', '.zip', '.pyc', '.pyo', '.env',
                                              'secrets.toml', '.pem', '.key')), relative
        assert not relative.startswith('Tool/outputs/') or relative == 'Tool/outputs/.gitkeep'
        if Path(relative).suffix.lower() in {'.py', '.ps1', '.bat', '.yaml', '.toml', '.md', '.txt', '.json'}:
            assert not re.search(rb'[A-Za-z]:[\\/]+Users[\\/]+', data, re.IGNORECASE), relative
    for relative in ('SETUP_VSM_TOOL.bat', 'START_VSM_TOOL.bat', 'Tool/scripts/client_setup.ps1',
                     'Tool/scripts/client_start.ps1', 'Tool/src/vsm_postprocessing/motor_capability_excel.py',
                     'Tool/src/vsm_postprocessing/profile_powerpoint_report_engine.py',
                     'Tool/config/report_profiles/sergio_2026_10/robosprayer_electric.yaml',
                     'Tool/config/report_profiles/sergio_2026_10/robosprayer_hybrid.yaml'):
        assert relative in files, relative
    manifest = {
        'product': 'VSM Engineering Post-Processing', 'version': __version__,
        'release_date': release_date, 'package_root': package_root,
        'source_commit': git('rev-parse', 'HEAD'), 'source_tree': git('rev-parse', 'HEAD^{tree}'),
        'source_state': 'validation candidate' if candidate else 'clean committed source',
        'python_range': '>=3.11,<3.12', 'bootstrap_python': '3.11.9',
        'build_script_sha256': hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
        'client_reference_data_included': False, 'generated_reports_included': False,
        'hybrid_acceptance': 'Latest motor-reporting real Hybrid acceptance pending; synthetic checks are not real-data validation.',
        'excluded_development_files': excluded,
        'files': [{'path': p, 'size_bytes': len(d), 'sha256': hashlib.sha256(d).hexdigest()}
                  for p, d in sorted(files.items())],
    }
    manifest['file_count'] = len(files)
    files['RELEASE_MANIFEST.json'] = (json.dumps(manifest, indent=2, sort_keys=True) + '\n').encode()
    with ZipFile(archive_path, 'w') as archive:
        for relative, data in sorted(files.items()):
            _zip_write_bytes(archive, package_root + '/' + relative, data)
    with ZipFile(archive_path) as archive:
        assert archive.testzip() is None and len(archive.namelist()) == len(files)
        for relative, data in files.items():
            assert archive.read(package_root + '/' + relative) == data
    digest = hashlib.sha256(archive_path.read_bytes()).hexdigest()
    archive_path.with_suffix('.sha256').write_text(digest + '  ' + archive_path.name + '\n', encoding='ascii')
    record = {'archive': str(archive_path.resolve()), 'sha256': digest,
              'size_bytes': archive_path.stat().st_size, 'file_count': len(files),
              'package_root': package_root, 'version': __version__, 'release_date': release_date,
              'source_commit': manifest['source_commit'], 'candidate': candidate}
    (destination / 'build.json').write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    return record


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project-root', type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--release-date', required=True, help='Release date, YYYY-MM-DD')
    parser.add_argument('--candidate', action='store_true')
    args = parser.parse_args()
    print(json.dumps(build(args.project_root, args.output_dir, args.release_date, args.candidate), indent=2))
