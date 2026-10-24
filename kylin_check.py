#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""跨平台环境诊断（Windows、银河麒麟、Linux）；仅使用 Python 标准库。"""
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
MAXIMUM_EXCLUSIVE = {
    'requests': (3,),
    'beautifulsoup4': (5,),
    'openpyxl': (4,),
    'python-docx': (2,),
    'urllib3': (3,),
}
ROOT = Path(__file__).resolve().parent


def package_upgrade_command(distribution, specification, system=None):
    """Return a native-shell command limited to the project's virtual env."""
    system = system or sys.platform
    if system == 'win32':
        return (r'.\.venv\Scripts\python.exe -m pip install --upgrade "' +
                specification + '"')
    return ("./.venv/bin/python -m pip install --upgrade '" +
            specification + "'")


def system_description(system=None):
    system = system or sys.platform
    if system == 'win32':
        return 'Windows'
    if system.startswith('linux'):
        return '银河麒麟 / Linux'
    if system == 'darwin':
        return 'macOS'
    return platform.system() or '其他系统'


def file_open_description(system=None, xdg_open=None):
    """Do not demand xdg-open on Windows or macOS."""
    system = system or sys.platform
    if system == 'win32':
        return 'Windows 文件关联（os.startfile）' + (
            '可用' if hasattr(os, 'startfile') else '不可用')
    if system == 'darwin':
        return 'macOS open 命令' + (
            '可用' if shutil.which('open') else '不可用')
    found = bool(shutil.which('xdg-open')) if xdg_open is None else bool(xdg_open)
    return 'Linux xdg-open 命令' + ('可用' if found else '不可用')


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
    maximum = MAXIMUM_EXCLUSIVE[distribution]
    specification = (distribution + '>=' + '.'.join(map(str, minimum)) +
                     ',<' + '.'.join(map(str, maximum)))
    upgrade = package_upgrade_command(distribution, specification)
    status = {'name': distribution, 'required': specification,
              'max_exclusive': '.'.join(map(str, maximum)),
              'installed': '', 'ok': False, 'error': '',
              'upgrade_command': upgrade,
              'install_command': (
                  upgrade if sys.platform == 'win32'
                  else 'sh ./市委大事记.sh install ' + distribution)}
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
            version = parse_version(raw)
            status['ok'] = minimum <= version < maximum
            if version < minimum:
                status['error'] = '版本过低'
            elif version >= maximum:
                status['error'] = '版本过高，超出项目支持范围'
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
                 else 'x86-64 / AMD64'
                 if architecture in ('amd64', 'x86_64') else '其他架构')
    suggestions = []
    if not python_ok:
        suggestions.append('当前 Python 仅可用于打开网页工作台及环境诊断；采集核心要求 Python 3.10+。')
    if not all(p['ok'] for p in packages):
        suggestions.append('请使用匹配当前 Python 的虚拟环境，不建议直接向系统 Python 强制安装 pip 包。')
        if sys.platform == 'win32':
            suggestions.append(
                r'Windows 可用 .\.venv\Scripts\python.exe -m pip install -r requirements.txt'
                ' 更新全部依赖。')
        else:
            suggestions.append('统一安装不合格依赖：sh ./市委大事记.sh install all（需确认）')
        for p in packages:
            if not p['ok']:
                suggestions.append(
                    p['name'] + ' 需要 ' + p['required'] + '；'
                    '单独安装/更新：' + p['install_command'] +
                    ('' if sys.platform == 'win32'
                     else '；等价的虚拟环境命令：' + p['upgrade_command']))
    if 'mips64' in architecture or 'mips64' in dpkg_arch:
        suggestions.append('龙芯 3A4000 是 MIPS64，不要安装标记为 x86_64、arm64、loongarch64 的软件包。')
    if not shutil.which('xdg-open') and sys.platform.startswith('linux'):
        suggestions.append('未发现 xdg-open：可以手动复制浏览器本机地址；直接打开 Excel/Word 按钮可能不可用。')
    if not (root / 'names.txt').exists() and not (root / 'names.xlsx').exists():
        suggestions.append('项目目录未发现 names.txt 或 names.xlsx，请检查是否完整解压了仓库。')
    return {
        'os': os_release() if sys.platform.startswith('linux') else {},
        'system': sys.platform,
        'system_name': system_description(),
        'os_display': (platform.platform()
                       if not sys.platform.startswith('linux')
                       else os_release().get('PRETTY_NAME', 'Linux')),
        'file_open': file_open_description(),
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
        result.get('system_name', system_description()) +
        ' 环境诊断（只读，不修改系统）',
        '系统：' + result.get('os_display',
                            result['os'].get('PRETTY_NAME', '未知')),
        'CPU 架构：' + result['machine'] + '；软件包架构：' + (result['dpkg_arch'] or '未知'),
        '架构说明：' + result['arch_description'],
        'Python 版本：' + result['python'],
        '采集条件：' + ('全部满足' if result['core_ready'] else '尚未满足'),
        '依赖检测：',
    ]
    for p in result['packages']:
        lines.append(
            '  ' + ('[正常] ' if p['ok'] else '[待处理] ') +
            p['required'] + '；当前：' +
            (p['installed'] or '未能导入') +
            (('；' + p['error']) if p['error'] else '')
        )
    lines.append('其他：文件打开方式 ' + result.get(
        'file_open', file_open_description(xdg_open=result.get('xdg_open'))))
    lines.append('其他：names.txt / names.xlsx ' +
                 ('已找到' if result['has_names'] else '未找到'))
    if result['suggestions']:
        lines.append('建议：')
        lines.extend('  - ' + line for line in result['suggestions'])
    lines.append('诊断结束：本命令没有安装软件，也没有修改用户文件。')
    return '\n'.join(lines)


def main(argv=None):
    parser = argparse.ArgumentParser(description='Windows / 银河麒麟环境只读诊断')
    parser.add_argument('--root', type=Path, default=ROOT, help='项目目录')
    parser.add_argument('--json', action='store_true', help='输出机器可读的诊断 JSON')
    args = parser.parse_args(argv)
    result = diagnose(args.root)
    print(json.dumps(result, ensure_ascii=False, indent=2)
          if args.json else format_diagnostics(result))
    return 0


if __name__ == '__main__':
    sys.exit(main())
