#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""九江新闻网单来源市委大事记采集器。"""
from __future__ import annotations

import argparse
import json
from contextlib import contextmanager
import logging
import os
import re
import shutil
import sqlite3
import sys
import time
from source_feeds import parse_jjxw_json
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin, urlsplit, urlunsplit, unquote

import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

LOG = logging.getLogger('jrjj')
SOURCES = {'jjxw': {'sheet': '九江新闻网', 'base': 'https://www.jjxw.cn/column/668/'}}
DEFAULT_NAMES = ['陈云', '邓永翔', '黄强', '陈永奇', '庄兆林', '黄喜忠']
HEADERS = ['新闻标题', '发布日期', '自动匹配', '命中姓名', '人工选择', '正文状态', '原文链接', '说明']
RE_DATE = re.compile(r'(20\d{2})[-年/.](1[0-2]|0?[1-9])[-月/.](3[01]|[12]\d|0?[1-9])(?:日)?(?!\d)')
RE_ISO = re.compile(r'20\d{2}-\d{2}-\d{2}')
INVALID_URL = re.compile(r'^(?:#|javascript:|mailto:|tel:)', re.I)
CONTENT_SELECTORS = [
    '#artibody > div.trs_editor_view', '#artibody', '#zoom',
    '.TRS_Editor', '.trs_editor_view', '.article-content', '.article_content',
    '.article-body', '.article_body', '.article-text', '.article_text',
    '.news-content', '.news_content', '.content-article', '.detail-content',
    '.detail_content', '.text-content', '.wzcon', '#content', '.content',
]
BLACKLIST_CLASS = re.compile(r'(related|recommend|copyright|share|comment|footer|nav|ad[-_]|pagination|source)', re.I)


@dataclass
class News:
    source: str
    title: str
    url: str
    pub_date: str


def norm_space(s: object) -> str:
    return re.sub(r'\s+', ' ', str(s or '').replace('\u200b', '')).strip()


def clean_title(s: object) -> str:
    return norm_space(s).strip(' -|_·\u3000')


def normalize_url(base: str, href: str) -> str:
    href = (href or '').strip()
    if not href or INVALID_URL.match(href):
        return ''
    # Some sites write a root-relative path without '/', e.g. column/668/article/...
    if re.match(r'^column/\d+/', href, re.I):
        parts = urlsplit(base)
        href = f'{parts.scheme}://{parts.netloc}/' + href
    url = urljoin(base, href)
    p = urlsplit(url)
    if p.scheme not in {'https', 'http'}:
        return ''
    path = p.path
    # The 668 column has published relative links that resolve to a 404
    # (/column/668/article/...), while articles live at /article/... .
    if 'jjxw.cn' in p.netloc.lower():
        path = re.sub(r'^/(?:column/\d+/)+article/', '/article/', path)
    return urlunsplit((p.scheme, p.netloc.lower(), path, p.query, ''))


def date_from_text(raw: str) -> str:
    m = RE_DATE.search(raw or '')
    if not m:
        return ''
    try:
        return date(*map(int, m.groups())).isoformat()
    except ValueError:
        return ''


def config_defaults() -> dict:
    return {
        'names': DEFAULT_NAMES,
        'include_body_matches': True,
        'max_pages': 170,
        'stop_after_known_pages': 3,
        'delay_seconds': 0.4,
        'timeout_seconds': 20,
        'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36',
    }


def load_config(path: Path) -> dict:
    if not path.exists():
        path.write_text(json.dumps(config_defaults(), ensure_ascii=False, indent=2), encoding='utf-8')
        LOG.info('已创建配置文件：%s', path)
    data = json.loads(path.read_text(encoding='utf-8'))
    cfg = {**config_defaults(), **data}
    cfg['names'] = sorted({norm_space(x) for x in cfg['names'] if norm_space(x) and len(norm_space(x)) >= 2})
    return cfg


def build_session(user_agent: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({'User-Agent': user_agent, 'Accept-Language': 'zh-CN,zh;q=0.9', 'Accept': 'text/html,application/xhtml+xml'})
    retry = Retry(total=3, connect=2, read=2, backoff_factor=0.8,
                  status_forcelist=(429, 500, 502, 503, 504), respect_retry_after_header=True)
    adapter = HTTPAdapter(max_retries=retry, pool_connections=8, pool_maxsize=8)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    return session


def fetch_html(session: requests.Session, url: str, timeout: int) -> str:
    r = session.get(url, timeout=timeout)
    r.raise_for_status()
    if not r.encoding or r.encoding.lower() in {'iso-8859-1', 'ascii'}:
        r.encoding = r.apparent_encoding or 'utf-8'
    return r.text


def fetch_list(session: requests.Session, page: int, cfg: dict,
               debug_dir: Path) -> tuple[list[News], bool]:
    """Only 九江新闻网 JSON listing, 100 articles per page."""
    url = (f'https://api.jjxw.cn/api/getArticles?cid=668&pageSize=100'
           f'&pageNumber={page + 1}&containSubCols=1')
    try:
        with session.get(url, timeout=cfg['timeout_seconds'],
                         headers={'Accept': 'application/json,text/plain,*/*',
                                  'Referer': 'https://www.jjxw.cn/'}) as response:
            response.raise_for_status()
            rows = parse_jjxw_json(response.content)
        if not rows:
            if page:
                return [], True
            raise ValueError('新闻网接口返回空列表或无法识别数据')
        LOG.info('新闻网 API 第%d页：%d条', page + 1, len(rows))
        return [News('jjxw', title, link, published)
                for title, link, published in rows], True
    except Exception as exc:
        debug_dir.mkdir(parents=True, exist_ok=True)
        (debug_dir / f'jjxw_feed_page_{page+1}.txt').write_text(
            f'URL: {url}\nERROR: {type(exc).__name__}: {exc}\n', encoding='utf-8')
        LOG.error('九江新闻网接口失败：%s — %s', url, exc)
        return [], False


def get_names(cfg: dict) -> list[str]:
    """Use only personal names explicitly supplied in the configuration."""
    return sorted(set(cfg['names']))


def clean_body(soup: BeautifulSoup) -> tuple[str, list[str]]:
    """Choose the best article container and omit nav/recommendation/ads before matching."""
    for tag in soup.select('script, style, noscript, iframe, form, svg, nav, footer, aside'):
        tag.decompose()
    for tag in list(soup.find_all(True)):
        if tag.attrs is None:  # A parent may have been decomposed already.
            continue
        attrs = f"{tag.get('id', '')} {' '.join(tag.get('class', []))}"
        if BLACKLIST_CLASS.search(attrs) and tag.name in {'div', 'section', 'ul'}:
            tag.decompose()
    candidates: list[tuple[float, list[str], str]] = []
    for priority, selector in enumerate(CONTENT_SELECTORS):
        for node in soup.select(selector)[:5]:
            paragraphs = [norm_space(p.get_text(' ', strip=True)) for p in node.select('p')]
            paragraphs = [p for p in paragraphs if len(p) >= 8]
            if not paragraphs:
                paragraphs = [norm_space(t) for t in node.get_text('\n', strip=True).split('\n')]
                paragraphs = [p for p in paragraphs if len(p) >= 12]
            text = '\n'.join(paragraphs)
            if len(text) < 65:
                continue
            all_a = sum(len(a.get_text(' ', strip=True)) for a in node.select('a'))
            ratio = all_a / max(len(text), 1)
            if ratio > .65:
                continue
            # More real text is useful; prefer article-specific containers for close ties.
            score = min(len(text), 20000) * (1 - ratio) + max(0, 20 - priority) * 18
            candidates.append((score, paragraphs, selector))
    if not candidates:
        # An unknown redesign should not force code changes solely to add a CSS selector.
        # Find the densest prose block with a few substantial paragraphs.
        for node in soup.select('article, main, section, div'):
            paragraphs = [norm_space(x.get_text(' ', strip=True)) for x in node.select('p')]
            paragraphs = [x for x in paragraphs if len(x) >= 15]
            length = sum(map(len, paragraphs))
            if length < 65 or len(paragraphs) < 2:
                continue
            link_chars = sum(len(a.get_text(' ', strip=True)) for a in node.select('a'))
            if link_chars / length > .22:
                continue
            # Penalize giant wrappers and prefer coherent, well-contained paragraphs.
            score = min(length, 12000) - link_chars * 4 - max(0, len(paragraphs) - 35) * 30
            candidates.append((score, paragraphs, '[自动发现的正文容器]'))
    if not candidates:
        return '', []
    score, paragraphs, selector = max(candidates, key=lambda x: x[0])
    cleaned = []
    for p in paragraphs:
        if p in cleaned[-1:]:
            continue
        if re.match(r'^(?:责任编辑|编辑：|来源：|分享到|打印本页|关闭窗口|上一篇|下一篇)', p) and len(p) < 90:
            continue
        cleaned.append(p)
    return selector, cleaned


def parse_article(html: str, fallback_title: str) -> tuple[str, list[str], str]:
    soup = BeautifulSoup(html, 'html.parser')
    title = ''
    for selector in ['#sybt .mainTitle', '#sybt', 'h1.article-title', '.article-title h1', '.article_title h1', 'h1.title', 'article h1', 'h1']:
        node = soup.select_one(selector)
        if node:
            candidate = clean_title(node.get_text(' ', strip=True))
            if len(candidate) >= 4 and len(candidate) <= 250:
                title = candidate
                break
    if not title:
        meta = soup.select_one('meta[property="og:title"], meta[name="title"]')
        title = clean_title(meta.get('content', '')) if meta else ''
    selector, paragraphs = clean_body(soup)
    return title or fallback_title, paragraphs, selector


@contextmanager
def connect_db(path: Path):
    """Manage a SQLite transaction AND close its file handle on all platforms.

    sqlite3.Connection.__exit__ commits/rolls back, but does not close the
    connection. Without an explicit close, Windows retains a lock on the
    database file and temporary-directory cleanup fails with WinError 32.
    """
    con = sqlite3.connect(path)
    try:
        con.row_factory = sqlite3.Row
        con.executescript('''
            CREATE TABLE IF NOT EXISTS news (
                url TEXT PRIMARY KEY,
                source TEXT NOT NULL,
                title TEXT NOT NULL,
                pub_date TEXT NOT NULL,
                article_title TEXT DEFAULT '',
                body_json TEXT DEFAULT '',
                status TEXT DEFAULT '未抓取',
                message TEXT DEFAULT '',
                fetched_at TEXT DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE INDEX IF NOT EXISTS idx_news_month ON news(pub_date, source);
            CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        ''')
        try:
            yield con
            con.commit()
        except BaseException:
            con.rollback()
            raise
    finally:
        con.close()


def upsert_news(con: sqlite3.Connection, record: News) -> bool:
    found = con.execute('SELECT url FROM news WHERE url=?', (record.url,)).fetchone()
    if found:
        con.execute('UPDATE news SET title=?,pub_date=?,source=? WHERE url=?',
                    (record.title, record.pub_date, record.source, record.url))
        return False
    con.execute('INSERT INTO news(url,source,title,pub_date) VALUES(?,?,?,?)',
                (record.url, record.source, record.title, record.pub_date))
    return True


def get_cell_url(cell, explicit='') -> str:
    if explicit and str(explicit).startswith(('https://', 'http://')):
        return str(explicit)
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    if cell.data_type == 'f' and isinstance(cell.value, str):
        m = re.search(r'^=HYPERLINK\(\s*"((?:[^"]|"")*)"', cell.value, re.I)
        return m.group(1).replace('""', '"') if m else ''
    if isinstance(cell.value, str) and cell.value.startswith('http'):
        return cell.value
    return ''


def get_cell_title(cell) -> str:
    v = cell.value
    if isinstance(v, str) and v.upper().startswith('=HYPERLINK('):
        m = re.search(r'^=HYPERLINK\(\s*"(?:[^"]|"")*"\s*[,;]\s*"((?:[^"]|"")*)"', v, re.I)
        return m.group(1).replace('""', '"') if m else ''
    return clean_title(v)


def is_yellow(cell) -> bool:
    if cell.fill is None or cell.fill.patternType != 'solid':
        return False
    color = cell.fill.fgColor
    if color.type == 'rgb' and color.rgb:
        return color.rgb.upper()[-6:] in {'FFFF00', 'FFEB3B', 'FFF200', 'FFEE00', 'FFF2CC'}
    return False


def selection(row) -> str:
    value = norm_space(row[4].value if len(row) > 4 else '').lower()
    if value in {'排除', '否', 'no', 'n', '0', '忽略'}:
        return 'exclude'
    if value in {'收录', '是', 'yes', 'y', '1', '人工收录'}:
        return 'include'
    if is_yellow(row[0]) or (len(row) > 1 and is_yellow(row[1])):
        return 'include'
    return 'auto'


def open_book(path: Path):
    """One worksheet only; main() backs up the original workbook first."""
    if path.exists():
        book = load_workbook(path)
    else:
        book = Workbook()
        book.active.title = '九江新闻网'
    name = SOURCES['jjxw']['sheet']
    if name not in book.sheetnames:
        book.create_sheet(name)
    for old in list(book.sheetnames):
        if old != name:
            LOG.warning('旧工作表“%s”不再处理，已保留在运行前Excel备份中', old)
            del book[old]
    ws = book[name]
    for i, header in enumerate(HEADERS, 1):
        cell = ws.cell(1, i)
        cell.value = header
        cell.fill = PatternFill('solid', fgColor='203A58')
        cell.font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for col, width in {'A': 73, 'B': 16, 'C': 14, 'D': 25,
                       'E': 14, 'F': 16, 'G': 68, 'H': 35}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = f'A1:H{max(ws.max_row, 2)}'
    ws.row_dimensions[1].height = 27
    if not ws.data_validations.dataValidation:
        validation = DataValidation(type='list', formula1='"收录,排除"', allow_blank=True)
        validation.error = '只能填写“收录”或“排除”'
        validation.errorTitle = '人工选择'
        validation.showErrorMessage = True
        ws.add_data_validation(validation)
        validation.add('E2:E1048576')
    return book


def existing_rows(book, con: sqlite3.Connection) -> dict[str, set[str]]:
    urls = {source: set() for source in SOURCES}
    for source, info in SOURCES.items():
        ws = book[info['sheet']]
        for idx in range(2, ws.max_row + 1):
            row = [ws.cell(idx, j) for j in range(1, 9)]
            url = get_cell_url(row[0], row[6].value)
            if not url:
                continue
            url = normalize_url(info['base'], url)
            if url in urls[source]:
                continue
            urls[source].add(url)
            d = date_from_text(str(row[1].value))
            if isinstance(row[1].value, (datetime, date)):
                d = row[1].value.strftime('%Y-%m-%d')
            title = get_cell_title(row[0])
            if row[0].data_type == 'f' and title:
                # Native hyperlinks survive editing and do not depend on WPS/Excel recalculation.
                # Changing only the value leaves the manual yellow fill intact.
                row[0].value = title
                row[0].hyperlink = url
            if d and title:
                upsert_news(con, News(source, title, url, d))
    con.commit()
    return urls


def write_new_rows(book, records: dict[str, list[News]]):
    for source, news in records.items():
        if not news:
            continue
        ws = book[SOURCES[source]['sheet']]
        new_items = sorted(news, key=lambda x: (x.pub_date, x.title), reverse=True)
        ws.insert_rows(2, len(new_items))
        for ix, item in enumerate(new_items, 2):
            a = ws.cell(ix, 1, item.title)
            a.hyperlink = item.url
            a.font = Font(name='楷体_GB2312', size=12, color='1565C0', underline='single')
            a.alignment = Alignment(vertical='center', wrap_text=True)
            b = ws.cell(ix, 2, item.pub_date)
            b.font = Font(name='微软雅黑', size=11)
            b.alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(ix, 7, item.url)
            ws.row_dimensions[ix].height = 28
        ws.auto_filter.ref = f'A1:H{ws.max_row}'
        LOG.info('%s：新增 %d 条', SOURCES[source]['sheet'], len(new_items))


def crawl(session: requests.Session, con: sqlite3.Connection,
          known: dict[str, set[str]], cfg: dict, year: int, debug: Path,
          target_month: str | None = None) -> dict[str, list[News]]:
    source = 'jjxw'
    new = {source: []}
    known_pages = 0
    for page in range(cfg['max_pages']):
        news, ok = fetch_list(session, page, cfg, debug)
        if not ok or not news:
            break
        target_prefix = target_month if target_month else f'{year}-'
        current = [n for n in news if n.pub_date.startswith(target_prefix)]
        old = bool(current) and all(n.url in known[source] for n in current)
        cutoff = f'{target_month}-01' if target_month else f'{year}-01-01'
        older = all(n.pub_date < cutoff for n in news)
        count = 0
        for n in current:
            if n.url not in known[source]:
                known[source].add(n.url)
                new[source].append(n)
                count += 1
            upsert_news(con, n)
        con.commit()
        LOG.info('第%d页：共%d条，目标范围%d条，新%d条',
                 page + 1, len(news), len(current), count)
        if older:
            break
        known_pages = known_pages + 1 if old else 0
        if known_pages >= cfg['stop_after_known_pages']:
            break
        if cfg['delay_seconds']:
            time.sleep(cfg['delay_seconds'])
    else:
        LOG.warning('达到 max_pages=%d，可能遗漏更早新闻', cfg['max_pages'])
    return new


def fetch_article_details(session: requests.Session, con: sqlite3.Connection,
                          cfg: dict, year: int, month: str,
                          new: dict[str, list[News]], book, scope: str):
    source = 'jjxw'
    fresh_urls = {x.url for x in new[source]}
    manual_urls = set()
    for row in book[SOURCES[source]['sheet']].iter_rows(min_row=2, max_col=8):
        url = get_cell_url(row[0], row[6].value)
        if selection(row) == 'include' and url:
            manual_urls.add(normalize_url(SOURCES[source]['base'], url))
    # Old databases may contain government rows; never fetch those.
    rows = con.execute(
        "SELECT * FROM news WHERE source=? AND pub_date LIKE ? ORDER BY pub_date DESC",
        (source, f'{year}-%')).fetchall()
    to_fetch = [
        n for n in rows if not (n['status'] == '成功' and n['body_json'])
        and ((scope == 'all' and n['url'] in fresh_urls)
             or n['pub_date'].startswith(month)
             or n['url'] in manual_urls)
    ]
    LOG.info('需要抓取正文%d篇；成功缓存直接复用', len(to_fetch))
    for idx, n in enumerate(to_fetch, 1):
        try:
            html = fetch_html(session, n['url'], cfg['timeout_seconds'])
            title, paragraphs, selector = parse_article(html, n['title'])
            if sum(map(len, paragraphs)) < 65:
                raise ValueError('正文少于65字，可能是访问限制或页面改版')
            con.execute(
                "UPDATE news SET article_title=?,body_json=?,status='成功',message=?,fetched_at=? WHERE url=?",
                (title, json.dumps(paragraphs, ensure_ascii=False), selector,
                 datetime.now().isoformat(timespec='seconds'), n['url']))
            LOG.info('[%d/%d] 正文成功：%s', idx, len(to_fetch), n['title'][:35])
        except Exception as exc:
            con.execute("UPDATE news SET status='失败',message=?,fetched_at=? WHERE url=?",
                        (str(exc)[:300], datetime.now().isoformat(timespec='seconds'), n['url']))
            LOG.warning('[%d/%d] 正文失败：%s — %s',
                        idx, len(to_fetch), n['title'][:35], exc)
        con.commit()
        if cfg['delay_seconds']:
            time.sleep(cfg['delay_seconds'])


def match_names(title: str, body: str, names: list[str], include_body: bool) -> list[str]:
    target = title + ('\n' + body if include_body else '')
    return [name for name in names if name in target]


def refresh_book(book, con: sqlite3.Connection, names: list[str], cfg: dict) -> list[dict]:
    chosen = []
    for source in SOURCES:
        ws = book[SOURCES[source]['sheet']]
        for i in range(2, ws.max_row + 1):
            row = [ws.cell(i, k) for k in range(1, 9)]
            url = get_cell_url(row[0], row[6].value)
            if not url:
                continue
            url = normalize_url(SOURCES[source]['base'], url)
            n = con.execute('SELECT * FROM news WHERE url=? AND source=?', (url, source)).fetchone()
            if n is None:
                continue
            paras = json.loads(n['body_json'] or '[]')
            matches = match_names(n['article_title'] or n['title'], '\n'.join(paras), names, cfg['include_body_matches'])
            choose = selection(row)
            selected = choose == 'include' or (choose == 'auto' and bool(matches))
            row[2].value = '命中' if matches else ('待正文' if n['status'] != '成功' else '未命中')
            row[3].value = '、'.join(matches)
            row[5].value = n['status']
            row[6].value = url
            row[7].value = n['message'] if n['status'] != '成功' else ('人工收录' if choose == 'include' else '人工排除' if choose == 'exclude' else '')
            row[2].alignment = Alignment(horizontal='center')
            row[4].alignment = Alignment(horizontal='center')
            row[5].alignment = Alignment(horizontal='center')
            if selected:
                chosen.append({'source': source, 'url': url, 'date': n['pub_date'],
                               'title': n['article_title'] or n['title'], 'body': paras,
                               'status': n['status'], 'matches': matches})
    return chosen


def canonical_title(title: str) -> str:
    return re.sub(r'[\W_]+', '', (title or '').lower())


def canonical_body(body: list[str]) -> str:
    """Ignore layout whitespace and punctuation when comparing full articles."""
    return re.sub(r'[\W_]+', '', '\n'.join(body).lower())


def _similar_article_body(a: str, b: str) -> bool:
    """Require substantial content overlap to deduplicate equal headlines.

    This prevents two different meetings with the same generic title from
    being collapsed merely because they were published close together.
    """
    if not a or not b:
        return False
    if a == b and len(a) >= 65:
        return True
    if min(len(a), len(b)) < 100:
        return False
    from collections import Counter
    a_pairs = Counter(a[i:i + 2] for i in range(len(a) - 1))
    b_pairs = Counter(b[i:i + 2] for i in range(len(b) - 1))
    if not a_pairs or not b_pairs:
        return False
    # Multiset counts keep long repeated phrases weighted correctly.
    shared = sum((a_pairs & b_pairs).values())
    total = sum((a_pairs | b_pairs).values())
    return shared / total >= 0.90


def _same_news_story(a: dict, b: dict) -> bool:
    a_body, b_body = canonical_body(a['body']), canonical_body(b['body'])
    # Reposting identical substantial body text under a different headline is
    # still one story, even if publication dates or titles differ.
    if a_body == b_body and len(a_body) >= 65:
        return True
    if canonical_title(a['title']) != canonical_title(b['title']):
        return False
    if abs((date.fromisoformat(a['date']) - date.fromisoformat(b['date'])).days) > 2:
        return False
    return _similar_article_body(a_body, b_body)


def deduplicate_for_month(selected: list[dict], month: str) -> tuple[list[dict], list[dict]]:
    """Return distinct articles for Word, leaving all source rows in Excel."""
    articles: list[dict] = []
    failed: list[dict] = []
    duplicates = 0
    eligible = sorted((x for x in selected if x['date'].startswith(month)),
                      key=lambda x: (x['date'], x['title']))
    for item in eligible:
        if item['status'] != '成功' or not item['body']:
            failed.append(item)
            continue
        match_idx = next((i for i, seen in enumerate(articles)
                          if _same_news_story(seen, item)), None)
        if match_idx is None:
            articles.append(item)
            continue
        duplicates += 1
        seen = articles[match_idx]
        # Keep the longer article among same-source duplicates.
        if len(canonical_body(item['body'])) > len(canonical_body(seen['body'])):
            articles[match_idx] = item
    articles.sort(key=lambda x: (x['date'], x['title']))
    LOG.info('Word 正文去重：候选正文 %d 篇，合并重复 %d 篇，最终保留 %d 篇',
             len(eligible) - len(failed), duplicates, len(articles))
    return articles, failed


def set_page_layout(doc):
    for section in doc.sections:
        section.page_width, section.page_height = Cm(21), Cm(29.7)
        section.top_margin, section.bottom_margin = Cm(2.54), Cm(2.54)
        section.left_margin, section.right_margin = Cm(2.7), Cm(2.7)
    normal = doc.styles['Normal']
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    normal.paragraph_format.line_spacing = Pt(28)
    normal.paragraph_format.space_before = normal.paragraph_format.space_after = Pt(0)


def add_run(p, text: str, font: str):
    r = p.add_run(text)
    r.font.name = font
    r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font)
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return r


def build_word(month: str, articles: list[dict], out_dir: Path) -> Path | None:
    if not articles:
        return None
    y, m = map(int, month.split('-'))
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f'市委大事记（{y}年{m}月）.docx'
    if path.exists():
        existing = Document(path)
        if 'jrjj-v2-generated' not in (existing.core_properties.keywords or ''):
            path = out_dir / f'市委大事记（{y}年{m}月）_自动整理.docx'
        if path.exists():
            backups = out_dir / '备份'
            backups.mkdir(exist_ok=True)
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            shutil.copy2(path, backups / f'{path.stem}_{stamp}.docx')
    doc = Document()
    set_page_layout(doc)
    doc.core_properties.keywords = 'jrjj-v2-generated'
    doc.core_properties.title = f'市委大事记（{y}年{m}月）'
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hrun = header.add_run('市 委 大 事 记')
    hrun.font.name = '方正小标宋简体'
    hrun._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), '方正小标宋简体')
    hrun.font.size = Pt(22)
    header.add_run('\n')
    date_run = header.add_run(f'（{y}年{m}月）')
    date_run.font.name = '楷体_GB2312'
    date_run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), '楷体_GB2312')
    date_run.font.size = Pt(18)
    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pnum = footer.add_run()
    pnum.font.name = '仿宋_GB2312'
    pnum.font.size = Pt(12)
    field_start = OxmlElement('w:fldChar')
    field_start.set(qn('w:fldCharType'), 'begin')
    field_instr = OxmlElement('w:instrText')
    field_instr.set(qn('xml:space'), 'preserve')
    field_instr.text = ' PAGE '
    field_end = OxmlElement('w:fldChar')
    field_end.set(qn('w:fldCharType'), 'end')
    pnum._r.append(field_start)
    pnum._r.append(field_instr)
    pnum._r.append(field_end)
    for article in articles:
        doc.add_paragraph('')
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = Pt(28)
        p.paragraph_format.keep_with_next = True
        add_run(p, article['title'], '黑体')
        ppr = p._p.get_or_add_pPr()
        lvl = OxmlElement('w:outlineLvl')
        lvl.set(qn('w:val'), '2')
        ppr.append(lvl)
        for para in article['body']:
            if not para:
                continue
            body = doc.add_paragraph()
            body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            body.paragraph_format.first_line_indent = Cm(1.12)
            body.paragraph_format.line_spacing = Pt(28)
            add_run(body, para, '仿宋_GB2312')
    temp = path.with_suffix('.tmp.docx')
    doc.save(temp)
    temp.replace(path)
    return path


def backup_spreadsheet(path: Path) -> None:
    if not path.exists():
        return
    backups = path.parent / '备份'
    backups.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    backup = backups / f'{path.stem}_运行前_{stamp}{path.suffix}'
    shutil.copy2(path, backup)
    LOG.info('已备份现有Excel：%s', backup)


def previous_month(today: date | None = None) -> str:
    today = today or date.today()
    return (today.replace(day=1) - timedelta(days=1)).strftime('%Y-%m')


def args_parser():
    parser = argparse.ArgumentParser(description='九江新闻网：增量抓取、姓名筛选、按月生成大事记')
    parser.add_argument('--month', help='输出年月，格式 YYYY-MM；默认上个月')
    parser.add_argument('--year', type=int, help='抓取年份，默认输出月份所在年份')
    parser.add_argument('--mode', choices=['all', 'sync', 'export'], default='all', help='all一键完整流程；sync只更新Excel和正文；export跳过列表抓取，只补正文并导出Word')
    parser.add_argument('--content-scope', choices=['all', 'month'], default='all', help='抓取新文章正文的范围：all或仅指定月份（旧Excel历史文章按需抓取）')
    parser.add_argument('--config', default='config.json', help='配置文件路径')
    parser.add_argument('--excel', default='jiujiang_news.xlsx', help='新闻列表Excel文件')
    parser.add_argument('--database', default='jiujiang_cache.sqlite3', help='新闻正文缓存数据库')
    parser.add_argument('--output-dir', default='.', help='Word输出目录')
    parser.add_argument('--debug-dir', default='debug_html', help='页面识别失败时保存HTML的目录')
    parser.add_argument('--max-pages', type=int, help='覆盖配置中的最多分页数')
    parser.add_argument('--full-scan', action='store_true', help='即使连续遇到旧记录，也扫完当年全部分页')
    parser.add_argument('--only-month', action='store_true', help='仅采集所选月份的新闻列表，遇到更早日期即停止；不会把其他月份写入Excel')
    return parser


def main(argv=None) -> int:
    args = args_parser().parse_args(argv)
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', datefmt='%H:%M:%S')
    month = args.month or previous_month()
    if not re.fullmatch(r'20\d{2}-(0[1-9]|1[0-2])', month):
        raise SystemExit('月份格式错误，请使用 --month 2026-09')
    year = args.year or int(month[:4])
    cfg = load_config(Path(args.config))
    if args.max_pages:
        cfg['max_pages'] = args.max_pages
    if args.full_scan:
        cfg['stop_after_known_pages'] = cfg['max_pages'] + 1
    LOG.info('目标月份：%s；抓取年份：%d；执行模式：%s', month, year, args.mode)
    backup_spreadsheet(Path(args.excel))
    book = open_book(Path(args.excel))
    with connect_db(Path(args.database)) as con:
        existing = existing_rows(book, con)
        session = build_session(cfg['user_agent'])
        new = {source: [] for source in SOURCES}
        if args.mode != 'export':
            new = crawl(session, con, existing, cfg, year, Path(args.debug_dir),
                        month if args.only_month else None)
            write_new_rows(book, new)
            # Save the new listing first; failed article requests must not lose the work.
            book.save(args.excel)
            fetch_article_details(session, con, cfg, year, month, new, book, args.content_scope)
        if args.mode == 'export':
            # Manual review: skip listing refresh; only fetch needed article bodies.
            fetch_article_details(session, con, cfg, year, month, new, book, 'month')
        # Always recompute matches from the current configuration; ignore old metadata.
        names = get_names(cfg)
        selected = refresh_book(book, con, names, cfg)
        book.save(args.excel)
        LOG.info('Excel已更新：%s；自动检测名单 %d 人', args.excel, len(names))
        if args.mode != 'sync':
            to_word, failed = deduplicate_for_month(selected, month)
            output = build_word(month, to_word, Path(args.output_dir))
            if output:
                LOG.info('Word已生成：%s，%d 篇（按日期由早到晚、正文去重）', output, len(to_word))
            else:
                LOG.warning('目标月份没有成功提取且符合收录条件的正文，未创建空Word')
            if failed:
                LOG.warning('有 %d 篇已选文章正文未抓取成功，本次未导出；下次会重试', len(failed))
        LOG.info('完成。请在Excel中使用E列“收录/排除”，或保持旧版黄色标记进行人工修订。')
    return 0


def configure_cli_encoding() -> None:
    """Make Chinese terminal output safe in Windows CI and redirected consoles.

    Some Windows environments default to cp1252. Setting the encoding before
    argparse handles --help prevents UnicodeEncodeError without requiring the
    caller to manually change the system code page.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            try:
                stream.reconfigure(encoding='utf-8', errors='backslashreplace')
            except (OSError, ValueError):
                pass


if __name__ == '__main__':
    configure_cli_encoding()
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        LOG.error('用户中断；已落盘的缓存、Excel仍被保留')
        sys.exit(130)
    except PermissionError as exc:
        LOG.error('文件被占用，请先关闭 Excel/Word 再运行：%s', exc)
        sys.exit(2)