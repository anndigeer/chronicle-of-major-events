#!/usr/bin/env python3
"""No-network preflight for GitHub Actions and local development."""
from __future__ import annotations
import importlib.metadata
import platform
import sys
from pathlib import Path

REQUIRED = (
    'jrjj_v2.py',
    'jrjj_config.json',
    'requirements_v2.txt',
    'test_jrjj_v2.py',
    'scripts/government_xml_smoke.py',
    '.github/workflows/news-collection.yml',
    'scripts/test_government_xml.py',
    '.github/workflows/python-ci.yml',
)

def main() -> int:
    print(f'Python: {sys.version}')
    print(f'OS: {platform.platform()}')
    print(f'Working directory: {Path.cwd()}')
    missing = []
    for path in REQUIRED:
        exists = Path(path).is_file()
        print(f'{"OK" if exists else "MISSING"}: {path}')
        if not exists:
            missing.append(path)
    for pkg in ('beautifulsoup4', 'openpyxl', 'python-docx', 'requests', 'urllib3'):
        try:
            print(f'{pkg}: {importlib.metadata.version(pkg)}')
        except importlib.metadata.PackageNotFoundError:
            print(f'MISSING dependency: {pkg}')
            missing.append(pkg)
    return int(bool(missing))

if __name__ == '__main__':
    raise SystemExit(main())
