"""九江新闻网 JSON 接口的纯 Python 解析器。"""
from __future__ import annotations

import html
import json
import re
from datetime import date

DATE = re.compile(r'(20\d{2})[-年/.](1[0-2]|0?[1-9])[-月/.](3[01]|[12]\d|0?[1-9])')


def day_from_text(text: object) -> str:
    m = DATE.search(str(text or ''))
    if not m:
        return ''
    try:
        return date(*map(int, m.groups())).isoformat()
    except ValueError:
        return ''


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
