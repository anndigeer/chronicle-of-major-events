"""Machine-readable news feeds for 九江市政府 and 九江新闻网.

These pure parsers have no network dependencies, and preserve the feed's
publication date rather than guessing from crawl time.
"""
from __future__ import annotations

import html
import json
import re
import xml.etree.ElementTree as ET
from datetime import date
from urllib.parse import urljoin, urlsplit, urlunsplit

DATE = re.compile(r'(20\d{2})[-年/.](1[0-2]|0?[1-9])[-月/.](3[01]|[12]\d|0?[1-9])')
URL_DATE = re.compile(r't(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])')
GOV_BASE = 'https://www.jiujiang.gov.cn/zwzx/jrjj/'


def day_from_text(text: object) -> str:
    match = DATE.search(str(text or ''))
    if match:
        try:
            return date(*map(int, match.groups())).isoformat()
        except ValueError:
            pass
    return ''


def parse_gov_xml(payload: str | bytes) -> list[tuple[str, str, str]]:
    """Return (title, URL, day) from PUBURL entries in the annual XML."""
    if isinstance(payload, str):
        payload = payload.lstrip('\ufeff')
    root = ET.fromstring(payload)
    parents = {c: p for p in root.iter() for c in p}

    def tag(node) -> str:
        return str(node.tag).rsplit('}', 1)[-1].upper()

    results = []
    seen = set()
    for link in root.iter():
        if tag(link) != 'PUBURL':
            continue
        entry = parents.get(link)
        if entry is None:
            continue
        fields = {tag(child): ''.join(child.itertext()).strip() for child in entry}
        href = fields.get('PUBURL', '').strip()
        url = urljoin(GOV_BASE, href)
        u = urlsplit(url)
        if (u.hostname != 'www.jiujiang.gov.cn' or
                not re.search(r'/20\d{4}/t20\d{6}_\d+\.s?html?$', u.path)):
            continue
        url = urlunsplit((u.scheme, u.netloc, u.path, '', ''))
        title = next((fields[k] for k in
                      ('TITLE', 'DOCTITLE', 'PUBTITLE', 'ARTTITLE', 'ARTICLE_TITLE', 'NAME')
                      if fields.get(k)), '')
        title = re.sub(r'<[^>]+>', '', html.unescape(title)).strip()
        dt = next((fields[k] for k in
                   ('RELTIME', 'PUBDATE', 'PUBTIME', 'PUBLISHDATE', 'PUBLISHTIME',
                    'DOCRELTIME', 'RELEASEDATE', 'DATE', 'CREATETIME')
                   if fields.get(k)), '')
        day = day_from_text(dt)
        if not day:
            m = URL_DATE.search(u.path)
            if m:
                day = day_from_text('-'.join(m.groups()))
        if title and day and url not in seen:
            results.append((title, url, day))
            seen.add(url)
    return results


def parse_jjxw_json(payload: str | bytes | dict) -> list[tuple[str, str, str]]:
    """Return (title, URL, day) from column 668's paginated API."""
    data = json.loads(payload) if isinstance(payload, (str, bytes)) else payload
    arrays = []

    def walk(x, depth=0):
        if depth > 8:
            return
        if isinstance(x, list):
            if any(isinstance(v, dict) and ('fileID' in v or 'fileId' in v) for v in x):
                arrays.append(x)
            else:
                for v in x[:5]:
                    walk(v, depth + 1)
        elif isinstance(x, dict):
            for v in x.values():
                if isinstance(v, (dict, list)):
                    walk(v, depth + 1)

    walk(data)
    found = []
    seen = set()
    for arr in arrays:
        for obj in arr:
            if not isinstance(obj, dict):
                continue
            ident = str(obj.get('fileID') or obj.get('fileId') or '').strip()
            day = day_from_text(obj.get('publishTime') or obj.get('publishedTime'))
            title = html.unescape(str(obj.get('title') or obj.get('articleTitle') or
                                      obj.get('fileTitle') or obj.get('name') or '')).strip()
            if not re.fullmatch(r'\d{1,15}', ident) or not day or len(title) < 2:
                continue
            yy, mm, dd = day.split('-')
            url = f'https://www.jjxw.cn/article/{yy}/{mm}/{dd}/{ident}.html'
            if url not in seen:
                found.append((title, url, day))
                seen.add(url)
    return found
