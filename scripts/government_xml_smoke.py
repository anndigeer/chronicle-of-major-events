#!/usr/bin/env python3
"""One bounded live test: annual government XML, September articles, and configured names."""
from __future__ import annotations

import argparse
import json
import logging
import socket
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import jrjj_v2 as app
from source_feeds import parse_gov_xml


def check(month: str, max_articles: int, max_matched: int, output: Path) -> int:
    output.mkdir(parents=True, exist_ok=True)
    cfg = app.load_config(Path("jrjj_config.json"))
    names = app.get_names(cfg)
    url = f"https://www.jiujiang.gov.cn/zwzx/jrjj/data_{month[:4]}.xml"
    result = {
        "month": month, "xml_url": url, "names": names,
        "started_at_utc": datetime.now(timezone.utc).isoformat(),
        "dns_ipv4": [], "dns_ipv6": [], "xml_bytes": 0,
        "annual_articles": 0, "month_articles": 0,
        "articles_attempted": 0, "articles_parsed": 0,
        "matched_articles": [], "article_errors": [], "stage": "dns",
        "success": False, "limited_sample": True,
    }
    report_path = output / "government_smoke.json"
    session = app.build_session(cfg["user_agent"])
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        try:
            answers = socket.getaddrinfo("www.jiujiang.gov.cn", 443, type=socket.SOCK_STREAM)
            result["dns_ipv4"] = sorted({x[4][0] for x in answers if x[0] == socket.AF_INET})
            result["dns_ipv6"] = sorted({x[4][0] for x in answers if x[0] == socket.AF_INET6})
        except OSError as exc:
            result["dns_error"] = str(exc)
        print("DNS IPv4:", result["dns_ipv4"], flush=True)
        print("DNS IPv6:", result["dns_ipv6"], flush=True)

        result["stage"] = "government_xml_http"
        raw = app.fetch_government_xml(session, url, min(cfg["timeout_seconds"], 8))
        result["xml_bytes"] = len(raw)
        (output / f"data_{month[:4]}.xml").write_bytes(raw)
        print("Downloaded XML bytes:", len(raw), flush=True)

        result["stage"] = "government_xml_parse"
        all_rows = parse_gov_xml(raw)
        rows = [r for r in all_rows if r[2].startswith(month)]
        result["annual_articles"] = len(all_rows)
        result["month_articles"] = len(rows)
        print(f"Parsed {len(all_rows)} annual records, {len(rows)} for {month}", flush=True)
        if not rows:
            raise ValueError(f"年度 XML 未找到 {month} 的新闻记录")

        # Try news with matching titles first, then body-only possibilities.
        rows.sort(key=lambda x: (not bool(app.match_names(x[0], "", names, False)), x[2], x[0]))
        result["stage"] = "article_body_http"
        for title, article_url, pub_date in rows[:max_articles]:
            result["articles_attempted"] += 1
            try:
                content = app.fetch_html(session, article_url, min(cfg["timeout_seconds"], 8))
                real_title, paragraphs, selector = app.parse_article(content, title)
                body = "\n".join(paragraphs)
                if len(body) < 65:
                    raise ValueError(f"正文过短（{len(body)} 字）")
                result["articles_parsed"] += 1
                matched = app.match_names(real_title, body, names, cfg["include_body_matches"])
                if matched:
                    result["matched_articles"].append({
                        "date": pub_date, "title": real_title, "url": article_url,
                        "matched_names": matched, "body_characters": len(body),
                        "paragraphs": len(paragraphs), "content_selector": selector,
                        "body_preview": body[:220],
                    })
                    print("MATCH", pub_date, matched, real_title, article_url, flush=True)
                else:
                    print("NO MATCH", pub_date, real_title, flush=True)
                if len(result["matched_articles"]) >= max_matched:
                    break
            except Exception as exc:
                result["article_errors"].append({
                    "title": title, "url": article_url,
                    "error": type(exc).__name__ + ": " + str(exc)[:400],
                })
                print("ARTICLE ERROR", type(exc).__name__, title, str(exc)[:160], flush=True)
            time.sleep(0.2)

        result["stage"] = "complete"
        result["success"] = bool(result["matched_articles"]) and result["articles_parsed"] > 0
        if not result["success"]:
            result["reason"] = "样本范围内没有成功提取且匹配配置姓名的正文"
    except Exception as exc:
        result["reason"] = type(exc).__name__ + ": " + str(exc)[:1500]
        print("FAILED at", result["stage"], "-", result["reason"], flush=True)
    finally:
        result["completed_at_utc"] = datetime.now(timezone.utc).isoformat()
        report_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        summary = [
            "### 九江市政府 XML / 关键姓名正文测试",
            f"- XML URL: {url}",
            f"- XML bytes: {result['xml_bytes']}",
            f"- 年度记录: {result['annual_articles']}; {month} 记录: {result['month_articles']}",
            f"- 抽样正文: {result['articles_parsed']}/{result['articles_attempted']} 成功",
            f"- 命中配置姓名: {len(result['matched_articles'])} 篇",
            f"- 结论: {'通过' if result['success'] else '失败'}",
            f"- 错误: {result.get('reason', '无')}",
        ]
        print("\n".join(summary), flush=True)
        step = __import__("os").environ.get("GITHUB_STEP_SUMMARY")
        if step:
            with open(step, "a", encoding="utf-8") as f:
                f.write("\n".join(summary) + "\n")
    return 0 if result["success"] else 1


def main() -> int:
    cli = argparse.ArgumentParser()
    cli.add_argument("--month", default="2026-09")
    cli.add_argument("--max-articles", type=int, default=25)
    cli.add_argument("--max-matched", type=int, default=3)
    cli.add_argument("--output-dir", type=Path, default=Path("results"))
    args = cli.parse_args()
    if not __import__("re").fullmatch(r"20\d{2}-(0[1-9]|1[0-2])", args.month):
        cli.error("month must be YYYY-MM")
    return check(args.month, max(1, min(args.max_articles, 40)),
                 max(1, min(args.max_matched, 5)), args.output_dir)


if __name__ == "__main__":
    raise SystemExit(main())
