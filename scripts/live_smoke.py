#!/usr/bin/env python3
"""Manual Actions smoke test against two public news websites."""
from __future__ import annotations
import logging
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import jrjj_v2 as jrjj

def main() -> int:
    logging.basicConfig(level=logging.INFO, format='%(levelname)s %(message)s')
    cfg = jrjj.config_defaults()
    session = jrjj.build_session(cfg['user_agent'])
    failed = []
    for source, info in jrjj.SOURCES.items():
        logging.info('检查新闻列表：%s', info['base'])
        try:
            items, ok, _ = jrjj.fetch_list(session, source, 0, cfg, False, Path('debug_html'))
            if not ok or not items:
                raise AssertionError('首页无法提取新闻列表')
            logging.info('%s 首页识别到 %s 条新闻', info['sheet'], len(items))
            article = items[0]
            html = jrjj.fetch_html(session, article.url, cfg['timeout_seconds'])
            title, paragraphs, selector = jrjj.parse_article(html, article.title)
            if len(''.join(paragraphs)) < 65:
                raise AssertionError('正文未成功提取（少于 65 字）')
            logging.info('%s 样本正文解析成功：%s；段落 %s；选择器 %s',info['sheet'], title,len(paragraphs),selector)
        except Exception as exc:
            failed.append(info['sheet'])
            logging.exception('%s 联网冒烟测试失败：%s', info['sheet'], exc)
    if failed:
        logging.error('失败来源：%s', '、'.join(failed))
        return 1
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
