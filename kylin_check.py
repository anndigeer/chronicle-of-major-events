#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""银河麒麟/龙芯本机部署诊断。只用 Python 标准库，绝不修改系统配置。"""
from __future__ import annotations

import argparse
import importlib
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys

REQUIRED = (
    ('requests', 'requests', (2, 32)),
    ('beautifulsoup4', 'bs4', (4, 13)),
    ('openpyxl', 'openpyxl', (3, 1)),
    ('python-docx', 'docx', (1, 2)),
    ('urllib3', 'urllib3', (1, 26)),
)
ROOT = Path(__file__).resolve().parent


def parse_version(text):
    match = re.match(r'(\d+(?:\.\d+)*)', str(text or ''))
    return tuple(int(x) for x in match.group(1).split('.')) if match else (0,)


def os_release(path=Path('/etc/os-release')):
    """Only the public OS name/version; never include user paths or secrets."""
    values = {}
    try:
        for line in path.read_text(encoding='utf-8').splitlines():
            if '=' not in line or line.startswith('#'):
                continue
            key, value = line.split('=', 1)
            if key in ('PRETTY_NAME', 'ID', 'VERSION_ID'):
                values[key] = value.strip().strip('"')
    except OSError:
        pass
    return values


def check_package(distribution, module, minimum):
    status = {'name': distribution, 'required': '.'.join(map(str, minimum)),
              'installed': '', 'ok': False, 'error': ''}
    try:
        # Some distributions ship modules through the system package manager;
        # package metadata may not exist, while imports still work.
        imported = importlib.import_module(module)
        raw = str(getattr(imported, '__version__', '') or '')
        try:
            from importlib import metadata
            raw = metadata.version(distribution)
        except (ImportError, metadata.PackageNotFoundError):
            pass
        status['installed'] = raw or '已安装（无法确认版本）'
        if not raw:
            status['error'] = '已检测到模块，但无法确认版本；请人工核对'
        else:
            status['ok'] = parse_version(raw) >= minimum
            if not status['ok']:
                status['error'] = '版本过低'
    except Exception as exc:
        status['error'] = type(exc).__name__ + ': ' + str(exc)[:140]
    return status


def diagnose(root=ROOT):
    root = Path(root)
    architecture = platform.machine().lower()
    dpkg_arch = ''
    if shutil.which('dpkg'):
        try:
            dpkg_arch = subprocess.check_output(
                ['dpkg', '--print-architecture'], text=True, timeout=3
            ).strip()
        except (OSError, subprocess.SubprocessError):
            pass
    version = platform.python_version()
    packages = [check_package(*args) for args in REQUIRED]
    python_ok = sys.version_info >= (3, 10)
    core_ready = python_ok and all(p['ok'] for p in packages)
    arch_kind = ('MIPS64（3A4000 兼容架构）'
                 if 'mips64' in architecture or 'mips64' in dpkg_arch
                 else 'LoongArch（不是 3A4000 的 MIPS64）'
                 if 'loongarch' in architecture or 'loongarch' in dpkg_arch
                 else '其他架构')
    suggestions = []
    if not python_ok:
        suggestions.append('当前 Python 仅可用于打开网页工作台及环境诊断；采集核心要求 Python 3.10+。')
    if not all(p['ok'] for p in packages):
        suggestions.append('检查银河麒麟原生软件源是否包含所需模块；若使用另装的 Python，请确保包安装在同一个解释器下。')
    if 'mips64' in architecture or 'mips64' in dpkg_arch:
        suggestions.append('龙芯 3A4000 是 MIPS64，不要安装标记为 x86_64、arm64、loongarch64 的软件包。')
    if not shutil.which('xdg-open') and sys.platform.startswith('linux'):
        suggestions.append('未发现 xdg-open：可以手动复制浏览器本机地址；直接打开 Excel/Word 按钮可能不可用。')
    if not (root / 'names.txt').exists() and not (root / 'names.xlsx').exists():
        suggestions.append('项目目录未发现 names.txt 或 names.xlsx，请检查是否完整解压了仓库。')
    return {
        'os': os_release(),
        'machine': architecture,
        'dpkg_arch': dpkg_arch,
        'arch_description': arch_kind,
        'python': version,
        'python_ok': python_ok,
        'core_ready': core_ready,
        'packages': packages,
        'xdg_open': bool(shutil.which('xdg-open')),
        'has_git': bool(shutil.which('git')),
        'has_names': (root / 'names.txt').exists() or (root / 'names.xlsx').exists(),
        'suggestions': suggestions,
    }


def format_diagnostics(result):
    lines = [
        '银河麒麟 / 龙芯 3A4000 环境诊断（只读，不修改系统）',
        '系统：' + result['os'].get('PRETTY_NAME', '未知'),
        'CPU 架构：' + result['machine'] + '；软件包架构：' + (result['dpkg_arch'] or '未知'),
        '架构说明：' + result['arch_description'],
        'Python 版本：' + result['python'],
        '采集条件：' + ('全部满足' if result['core_ready'] else '尚未满足'),
        '依赖检测：',
    ]
    for p in result['packages']:
        lines.append(
            '  ' + ('[正常] ' if p['ok'] else '[待处理] ') +
            p['name'] + ' >= ' + p['required'] + '；当前：' +
            (p['installed'] or '未能导入') +
            (('；' + p['error']) if p['error'] else '')
        )
    lines.append('其他：系统文件打开命令 xdg-open ' +
                 ('可用' if result['xdg_open'] else '不可用'))
    lines.append('其他：names.txt / names.xlsx ' +
                 ('已找到' if result['has_names'] else '未找到'))
    if result['suggestions']:
        lines.append('建议：')
        lines.extend('  - ' + line for line in result['suggestions'])
    lines.append('诊断结束：本命令没有安装软件，也没有修改用户文件。')
    return '\n'.join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description='银河麒麟/龙芯3A4000环境只读诊断')
    parser.add_argument('--root', type=Path, default=ROOT, help='项目目录')
    parser.add_argument('--json', action='store_true', help='输出机器可读的诊断 JSON')
    args = parser.parse_args(argv)
    result = diagnose(args.root)
    print(json.dumps(result, ensure_ascii=False, indent=2)
          if args.json else format_diagnostics(result))
    return 0


if __name__ == '__main__':
    sys.exit(main())
