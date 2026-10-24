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
    def test_version_check(self):
        self.assertEqual(doctor.parse_version('3.1.5'), (3, 1, 5))
        self.assertEqual(doctor.parse_version('2.32.3.post1'), (2, 32, 3))
        self.assertEqual(doctor.parse_version('unknown'), (0,))

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
                                  'required':'1.0','error':'缺少软件包'}), \
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
