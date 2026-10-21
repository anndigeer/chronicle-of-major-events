#!/usr/bin/env python3
"""Bounded live test of 2026 government XML and configured leader article bodies."""
from __future__ import annotations

import json
import logging
import socket
import sys
import time
from pathlib import Path
from urllib.parse import urlsplit

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import jrjj_v2 as j
from source_feeds import parse_gov_xml

OUT = Path("diagnostics")
OUT.mkdir(exist_ok=True)
URL = "https://www.jiujiang.gov.cn/zwzx/jrjj/data_2026.xml"
MONTH = "2026-09"

def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    cfg = j.load_config(Path("jrjj_config.json"))
    names = j.get_names(cfg)
    report = {"xml_url": URL, "month": MONTH, "approved_names": names,
              "xml_downloaded": False, "xml_articles": 0, "month_articles": 0,
              "title_matches": 0, "body_attempts": 0, "body_success": 0,
              "body_matches": [], "errors": []}
    started = time.monotonic()
    try:
        try:
            report["dns_v4"] = sorted({record[4][0] for record in socket.getaddrinfo(
                "www.jiujiang.gov.cn", 443, family=socket.AF_INET, type=socket.SOCK_STREAM)})
            print("IPv4 DNS:", report["dns_v4"], flush=True)
        except OSError as exc:
            report["dns_error"] = str(exc)
        session = j.build_session(cfg["user_agent"])
        payload = j.fetch_government_xml(session, URL, 7)
        report["xml_downloaded"] = True
        report["xml_bytes"] = len(payload)
        (OUT / "government_2026.xml").write_bytes(payload)
        rows = parse_gov_xml(payload)
        report["xml_articles"] = len(rows)
        if not rows:
            raise ValueError("XML downloaded but parsed 0 articles")
        monthly = [(title, url, date) for title, url, date in rows if date.startswith(MONTH)]
        report["month_articles"] = len(monthly)
        if not monthly:
            raise ValueError("No news records published in September 2026")
        print("XML records:", len(rows), "September records:", len(monthly), flush=True)
        title_hits = [r for r in monthly if j.match_names(r[0], "", names, False)]
        report["title_matches"] = len(title_hits)
        # Prefer titles containing configured personal names; if there are none,
        # inspect a bounded sample of bodies without flooding the website.
        sample = title_hits[:6] if title_hits else monthly[:6]
        report["sampled"] = len(sample)
        for title, url, published in sample:
            if urlsplit(url).hostname != "www.jiujiang.gov.cn":
                report["errors"].append("Unexpected article host: " + url)
                continue
            report["body_attempts"] += 1
            try:
                html = j.fetch_html(session, url, 8)
                article_title, paras, selector = j.parse_article(html, title)
                length = len("".join(paras))
                if length < 65:
                    raise ValueError("Body too short: " + str(length))
                report["body_success"] += 1
                found = j.match_names(article_title, "\n".join(paras), names, True)
                if found:
                    report["body_matches"].append({
                        "date": published, "title": article_title,
                        "url": url, "matched_names": found,
                        "characters": length, "selector": selector
                    })
                print("ARTICLE:", published, title[:50], "body_chars:", length,
                      "matched_names:", found, flush=True)
            except Exception as exc:
                report["errors"].append("Article " + url + ": " + str(exc))
                print("ARTICLE FAILED:", url, repr(exc), flush=True)
        if not report["body_matches"]:
            raise ValueError("No September article with configured leader names and valid body")
    except Exception as exc:
        report["errors"].append(type(exc).__name__ + ": " + str(exc))
        print("FAILED:", repr(exc), flush=True)
    finally:
        report["elapsed_seconds"] = round(time.monotonic() - started, 2)
        (OUT / "government_xml_report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print("SUMMARY:", json.dumps({k: v for k, v in report.items()
            if k not in ("body_matches", "approved_names")}, ensure_ascii=False), flush=True)
    return 0 if report["xml_downloaded"] and report["body_matches"] else 1

if __name__ == "__main__":
    raise SystemExit(main())
