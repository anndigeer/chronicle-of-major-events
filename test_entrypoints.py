"""Guard the two platform launchers and the Chinese end-user manual."""
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parent


class TestRenamedStartupAndUserGuide(unittest.TestCase):
    def test_renamed_launchers_exist_without_legacy_files(self):
        for filename in ('Start.cmd', 'Start.sh'):
            with self.subTest(name=filename):
                self.assertTrue((ROOT / filename).is_file())
                self.assertGreater((ROOT / filename).stat().st_size, 200)
        for obsolete in ('市委大事记.cmd', '市委大事记.sh'):
            self.assertFalse((ROOT / obsolete).exists())
        self.assertIn('Start.cmd', (ROOT / 'Start.cmd').read_text(encoding='utf-8-sig'))
        self.assertIn('Start.sh', (ROOT / 'Start.sh').read_text(encoding='utf-8-sig'))

    def test_document_covers_both_platforms_and_manual_workflow(self):
        manual = (ROOT / '使用文档.md').read_text(encoding='utf-8')
        readme = (ROOT / 'README.md').read_text(encoding='utf-8')
        for phrase in ('Start.cmd', 'Start.sh', 'Windows', '银河麒麟',
                       'MIPS64', 'Python 3.10', 'install all', 'install openpyxl',
                       'jiujiang_news.xlsx', '--sync', '--export', '--verify',
                       'names.xlsx', 'names.txt', '常见问题', '事件日期',
                       '人工审核', '备份', '127.0.0.1'):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, manual)
        self.assertIn('[使用文档.md](使用文档.md)', readme)
        self.assertIn('sh ./Start.sh', readme)
        self.assertIn('.\\Start.cmd', readme)
        self.assertNotIn('市委大事记.cmd', readme)
        self.assertNotIn('市委大事记.sh', readme)

    def test_linux_desktop_entry_points_to_new_launcher(self):
        shell = (ROOT / 'Start.sh').read_text(encoding='utf-8')
        smoke = (ROOT / 'tests' / 'kylin_shell_smoke.sh').read_text(
            encoding='utf-8')
        windows = (ROOT / 'tests' / 'windows_cmd_smoke.py').read_text(
            encoding='utf-8')
        self.assertIn('$ROOT/Start.sh', shell)
        self.assertIn('Start.sh', smoke)
        self.assertIn('Start.cmd', windows)
        self.assertNotIn('市委大事记.sh', shell)
        self.assertNotIn('市委大事记.cmd', windows)


if __name__ == '__main__':
    unittest.main()
