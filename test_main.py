"""九江新闻网单数据源离线回归与端到端测试。"""
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from docx import Document
import main as j
from source_feeds import parse_jjxw_json


def article(title, body):
    return ('<html><h1>' + title + '</h1><div class="article-content">'
            '<p>' + body + '</p><p>' + body + '</p></div></html>')


class TestJjxw(unittest.TestCase):
    def test_only_news_source(self):
        self.assertEqual(list(j.SOURCES), ['jjxw'])
        self.assertNotIn('--gov-xml-file', j.args_parser().format_help())
        self.assertNotIn('--browser', j.args_parser().format_help())

    def test_news_json_api(self):
        obj = {'data': {'list': [
            {'fileID': 190001, 'publishTime': '2026-09-18 10:00:00',
             'title': '陈云调研'},
            {'fileID': 190002, 'publishTime': '2026-09-17 10:00:00',
             'title': '普通报道'}
        ]}}
        rows = parse_jjxw_json(obj)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0][2], '2026-09-18')
        self.assertEqual(rows[0][1],
                         'https://www.jjxw.cn/article/2026/09/18/190001.html')

    def test_organization_terms_not_personal_names(self):
        names = j.get_names(j.config_defaults())
        title = '市人大常委会党组理论学习中心组举行集体学习，廖奇志主持并讲话'
        self.assertEqual(j.match_names(title, '省委政协人大', names, True), [])
        self.assertEqual(j.match_names('陈云调研', '', names, False), ['陈云'])

    def test_duplicate_body_different_titles(self):
        body = '陈云部署推进重大项目建设质量效率。' * 9
        items = [
            dict(source='jjxw', title='报道甲', url='a', date='2026-09-18',
                 body=[body], status='成功'),
            dict(source='jjxw', title='报道乙', url='b', date='2026-09-18',
                 body=[body], status='成功')
        ]
        retained, failures = j.deduplicate_for_month(items, '2026-09')
        self.assertEqual((len(retained), len(failures)), (1, 0))

    def test_same_title_different_body_kept(self):
        a = '陈云安排产业建设发展合作工作。' * 10
        b = '黄强主持财政会议研究财政预算计划。' * 10
        rows = [
            dict(source='jjxw', title='市委召开会议', url='a', date='2026-09-18',
                 body=[a], status='成功'),
            dict(source='jjxw', title='市委召开会议', url='b', date='2026-09-18',
                 body=[b], status='成功')
        ]
        retained, failures = j.deduplicate_for_month(rows, '2026-09')
        self.assertEqual((len(retained), len(failures)), (2, 0))

    def test_migrate_old_excel_with_backup(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder) / 'old.xlsx'
            old = Workbook()
            govt = old.active
            govt.title = '九江市政府'
            govt.append(['新闻标题', '发布日期'])
            govt.append(['旧新闻', '2026-09-18'])
            news = old.create_sheet('九江新闻网')
            news.append(list(j.HEADERS))
            news.cell(2, 1, '陈云在九江调研')
            news.cell(2, 2, '2026-09-18')
            news.cell(2, 5, '收录')
            news.cell(2, 7,
                      'https://www.jjxw.cn/article/2026/09/18/190001.html')
            news.cell(2, 1).fill = PatternFill('solid', fgColor='FFFF00')
            old.save(p)
            j.backup_spreadsheet(p)
            converted = j.open_book(p)
            self.assertEqual(converted.sheetnames, ['九江新闻网'])
            self.assertEqual(converted.active['E2'].value, '收录')
            self.assertTrue(j.is_yellow(converted.active['A2']))
            converted.save(p)
            copies = list((Path(folder) / '备份').glob('*.xlsx'))
            self.assertEqual(len(copies), 1)
            self.assertIn('九江市政府', load_workbook(copies[0]).sheetnames)

    def test_old_government_cache_is_ignored(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder)
            with j.connect_db(p / 'cache.sqlite3') as con:
                con.execute("INSERT INTO news(url,source,title,pub_date) VALUES(?,?,?,?)",
                            ('https://www.jiujiang.gov.cn/old.html',
                             'gov', '陈云旧报道', '2026-09-18'))
                book = j.open_book(p / 'news.xlsx')
                cfg = dict(j.config_defaults(), delay_seconds=0)
                with patch.object(j, 'fetch_html',
                                  side_effect=AssertionError('should not fetch')) as fetch:
                    j.fetch_article_details(j.build_session('test'), con, cfg,
                                            2026, '2026-09', {'jjxw': []},
                                            book, 'month')
                    fetch.assert_not_called()

    def test_mock_month_export_and_document(self):
        with tempfile.TemporaryDirectory() as folder:
            p = Path(folder)
            cfg = dict(j.config_defaults(), names=['陈云'],
                       max_pages=3, delay_seconds=0, stop_after_known_pages=3)
            (p / 'config.json').write_text(
                json.dumps(cfg, ensure_ascii=False), encoding='utf-8')
            listing = {'data': {'list': [
                {'fileID': 190001, 'publishTime': '2026-09-18 10:00:00',
                 'title': '陈云调研甲'},
                {'fileID': 190002, 'publishTime': '2026-09-18 09:00:00',
                 'title': '陈云调研乙'},
                {'fileID': 190003, 'publishTime': '2026-09-17 09:00:00',
                 'title': '普通报道'}
            ]}}
            shared = '陈云部署产业发展与高质量建设项目相关工作。' * 7
            ordinary = '全市落实交通建设和公共服务改善各项任务。' * 7

            def fake_html(session, url, timeout):
                if url.endswith('190001.html'):
                    return article('陈云调研甲', shared)
                if url.endswith('190002.html'):
                    return article('陈云调研乙', shared)
                if url.endswith('190003.html'):
                    return article('普通报道', ordinary)
                raise AssertionError('Unexpected article URL: ' + url)

            class Response:
                def __init__(self, content):
                    self.content = content
                def raise_for_status(self):
                    return None
                def __enter__(self):
                    return self
                def __exit__(self, *args):
                    return False

            def fake_get(self, url, **kwargs):
                if 'getArticles?' not in url:
                    raise AssertionError('Unexpected feed URL: ' + url)
                page = listing if 'pageNumber=1' in url else {'data': {'list': []}}
                return Response(json.dumps(page, ensure_ascii=False).encode('utf-8'))

            with patch.object(j.requests.Session, 'get', fake_get), \
                 patch.object(j, 'fetch_html', fake_html):
                code = j.main([
                    '--config', str(p / 'config.json'),
                    '--month', '2026-09', '--year', '2026', '--only-month',
                    '--excel', str(p / 'news.xlsx'),
                    '--database', str(p / 'cache.sqlite3'),
                    '--output-dir', str(p),
                    '--debug-dir', str(p / 'debug')])
            self.assertEqual(code, 0)
            wb = load_workbook(p / 'news.xlsx')
            self.assertEqual(wb.sheetnames, ['九江新闻网'])
            self.assertEqual(wb.active.max_row, 4)
            self.assertEqual(sum(wb.active.cell(i, 3).value == '命中'
                                 for i in range(2, 5)), 2)
            doc = Document(p / '市委大事记（2026年9月）.docx')
            titles = [paragraph.text for paragraph in doc.paragraphs
                      if paragraph._p.xpath('./w:pPr/w:outlineLvl')]
            self.assertEqual(len(titles), 1)
            self.assertIn('陈云', titles[0])


if __name__ == '__main__':
    unittest.main()
