"""Galaxy Kylin diagnosis regression tests: Python standard library only."""
import io
import json
from pathlib import Path
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch
import kylin_check as doctor


class TestKylinDiagnostics(unittest.TestCase):
    def test_windows_diagnostics_do_not_claim_kylin_or_need_xdg_open(self):
        with tempfile.TemporaryDirectory() as folder:
            fake={'name':'requests','required':'requests>=2.32,<3',
                  'installed':'2.34.2','ok':True,'error':''}
            with patch.object(doctor.sys,'platform','win32'), \
                 patch.object(doctor.platform,'machine',return_value='AMD64'), \
                 patch.object(doctor.platform,'platform',
                              return_value='Windows-11-10.0.26100'), \
                 patch.object(doctor,'check_package',return_value=fake), \
                 patch.object(doctor,'os_release',
                              side_effect=AssertionError('not linux')), \
                 patch.object(doctor.shutil,'which',return_value=None):
                result=doctor.diagnose(Path(folder))
                text=doctor.format_diagnostics(result)
            self.assertIn('Windows 环境诊断',text)
            self.assertIn('系统：Windows-11-10.0.26100',text)
            self.assertIn('x86-64 / AMD64',text)
            self.assertIn('Windows 文件关联',text)
            self.assertNotIn('银河麒麟 / 龙芯 3A4000 环境诊断',text)
            self.assertNotIn('文件打开命令 xdg-open',text)
            self.assertTrue(result['core_ready'])

    def test_windows_dependency_upgrade_command_uses_powershell_syntax(self):
        with patch.object(doctor.sys,'platform','win32'):
            upgrade=doctor.package_upgrade_command(
                'openpyxl','openpyxl>=3.1,<4')
            self.assertEqual(upgrade,
                r'.\.venv\Scripts\python.exe -m pip install --upgrade "openpyxl>=3.1,<4"')
            self.assertEqual(doctor.system_description(), 'Windows')
        with patch.object(doctor.sys,'platform','linux'):
            upgrade=doctor.package_upgrade_command(
                'openpyxl','openpyxl>=3.1,<4')
            self.assertIn('./.venv/bin/python',upgrade)
            self.assertEqual(doctor.system_description(), '银河麒麟 / Linux')

    def test_version_check(self):
        self.assertEqual(doctor.parse_version('3.1.5'), (3, 1, 5))
        self.assertEqual(doctor.parse_version('2.32.3.post1'), (2, 32, 3))
        self.assertEqual(doctor.parse_version('unknown'), (0,))

    def test_dependency_range_and_exact_single_package_command(self):
        from types import SimpleNamespace
        with patch('importlib.metadata.version', return_value='3.0.10'), \
             patch('kylin_check.importlib.import_module',
                   return_value=SimpleNamespace(__version__='3.0.10')):
            status=doctor.check_package('openpyxl','openpyxl',(3,1))
        self.assertFalse(status['ok'])
        self.assertEqual(status['error'], '版本过低')
        self.assertEqual(status['required'], 'openpyxl>=3.1,<4')
        if doctor.sys.platform == 'win32':
            self.assertIn(r'.\.venv\Scripts\python.exe', status['install_command'])
            self.assertIn('"openpyxl>=3.1,<4"', status['upgrade_command'])
        else:
            self.assertEqual(status['install_command'],
                             'sh ./市委大事记.sh install openpyxl')
            self.assertIn("'openpyxl>=3.1,<4'", status['upgrade_command'])
        with patch('importlib.metadata.version', return_value='4.0.0'), \
             patch('kylin_check.importlib.import_module',
                   return_value=SimpleNamespace(__version__='4.0.0')):
            newer=doctor.check_package('openpyxl','openpyxl',(3,1))
        self.assertFalse(newer['ok'])
        self.assertIn('版本过高', newer['error'])

    def test_diagnostics_prints_only_actionable_targeted_upgrades(self):
        with tempfile.TemporaryDirectory() as folder:
            packages=[
                {'name':'requests','required':'requests>=2.32,<3',
                 'installed':'2.31.0','ok':False,'error':'版本过低',
                 'install_command':'sh ./市委大事记.sh install requests',
                 'upgrade_command':"./.venv/bin/python -m pip install --upgrade 'requests>=2.32,<3'"},
                {'name':'openpyxl','required':'openpyxl>=3.1,<4',
                 'installed':'3.1.5','ok':True,'error':'',
                 'install_command':'sh ./市委大事记.sh install openpyxl',
                 'upgrade_command':"./.venv/bin/python -m pip install --upgrade 'openpyxl>=3.1,<4'"},
            ]
            with patch.object(doctor,'REQUIRED',(
                ('requests','requests',(2,32)),
                ('openpyxl','openpyxl',(3,1)) )), \
                 patch.object(doctor,'check_package',
                              side_effect=packages), \
                 patch.object(doctor.platform,'machine',
                              return_value='mips64el'), \
                 patch.object(doctor.shutil,'which',return_value=None):
                result=doctor.diagnose(Path(folder))
            text=doctor.format_diagnostics(result)
            self.assertIn('install requests',text)
            self.assertNotIn('install openpyxl',text)
            self.assertIn('requests>=2.32,<3',text)
            if doctor.sys.platform == 'win32':
                self.assertIn('pip install -r requirements.txt',text)
            else:
                self.assertIn('install all',text)

    def test_os_release_only_reads_public_fields(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'os-release'
            path.write_text(
                'PRETTY_NAME="银河麒麟桌面操作系统 V10"\n'
                'VERSION_ID="10"\n'
                'PRIVATE_TOKEN="never-expose-this"\n', encoding='utf-8')
            values=doctor.os_release(path)
            self.assertEqual(values['PRETTY_NAME'],'银河麒麟桌面操作系统 V10')
            self.assertEqual(values['VERSION_ID'],'10')
            self.assertNotIn('PRIVATE_TOKEN',values)

    def test_no_code_execution_or_installer_side_effects(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            (root/'names.txt').write_text('陈云\n',encoding='utf-8')
            before={p.name for p in root.iterdir()}
            with patch.object(doctor.platform, 'machine', return_value='mips64el'), \
                 patch.object(doctor, 'check_package',
                              side_effect=lambda d, m, v: {
                                  'name':d, 'installed':'', 'ok':False,
                                  'required':'example>=1,<2','error':'缺少软件包',
                                  'install_command':'sh ./市委大事记.sh install example',
                                  'upgrade_command':"./.venv/bin/python -m pip install --upgrade 'example>=1,<2'"}), \
                 patch.object(doctor.shutil, 'which', return_value=None):
                result=doctor.diagnose(root)
            self.assertIn('MIPS64',result['arch_description'])
            self.assertFalse(result['core_ready'])
            self.assertEqual(len(result['packages']), 5)
            self.assertEqual({p.name for p in root.iterdir()}, before)
            report=doctor.format_diagnostics(result)
            self.assertIn('采集条件：尚未满足',report)
            self.assertIn('缺少软件包',report)

    def test_json_cli_with_no_local_writes(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            text=io.StringIO()
            with redirect_stdout(text):
                self.assertEqual(doctor.main(['--root',str(root),'--json']),0)
            payload=json.loads(text.getvalue())
            self.assertIn('python',payload)
            self.assertIn('machine',payload)
            self.assertIn('packages',payload)
            self.assertEqual(list(root.iterdir()), [])

    def test_core_ready_false_on_old_python_even_with_packages(self):
        good={'name':'example','required':'1','installed':'2','error':'','ok':True}
        with tempfile.TemporaryDirectory() as folder:
            class Python39(tuple):
                def __ge__(self, other):
                    return False
            with patch.object(doctor.sys, 'version_info', Python39((3,9))), \
                 patch.object(doctor, 'check_package', return_value=good):
                result=doctor.diagnose(Path(folder))
            self.assertFalse(result['python_ok'])
            self.assertFalse(result['core_ready'])
            self.assertTrue(any('3.10' in x for x in result['suggestions']))


if __name__=='__main__':
    unittest.main()
