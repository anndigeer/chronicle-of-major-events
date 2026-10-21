#!/usr/bin/env python3
"""Verify and summarize a real monthly extraction run, preserving output artifacts."""
from __future__ import annotations
import argparse
import json
import os
from collections import Counter
from pathlib import Path
from openpyxl import load_workbook
from docx import Document

def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument('--month', default='2026-09')
    p.add_argument('--folder', default='output')
    args = p.parse_args()
    folder = Path(args.folder)
    folder.mkdir(parents=True, exist_ok=True)
    result = {'month': args.month, 'sheets': {}, 'docx': [], 'xlsx': [], 'warnings': [], 'success': False}
    xlsx_files = sorted(folder.glob('*.xlsx'))
    docx_files = sorted(folder.glob('*.docx'))
    result['xlsx'] = [x.name for x in xlsx_files]
    result['docx'] = [x.name for x in docx_files]
    errors = []
    if not xlsx_files:
        errors.append('未生成xlsx')
    for xlsx in xlsx_files:
        try:
            wb = load_workbook(xlsx, read_only=True)
            for sheet in ('九江市政府', '九江新闻网'):
                if sheet not in wb.sheetnames:
                    errors.append(f'{xlsx.name} 缺少工作表 {sheet}')
                    continue
                ws = wb[sheet]
                dates, statuses = [], Counter()
                matched = 0
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or not row[0]:
                        continue
                    dates.append(str(row[1] or ''))
                    statuses[str(row[5] if len(row) > 5 else '')] += 1
                    if len(row) > 2 and row[2] == '命中':
                        matched += 1
                other_months = [d for d in dates if not d.startswith(args.month)]
                info = {'news': len(dates), 'matched': matched, 'body_status': dict(statuses), 'other_months': len(other_months)}
                result['sheets'][sheet] = info
                print(f'{sheet}: 共 {len(dates)} 条；匹配 {matched} 条；正文状态 {dict(statuses)}；非目标月份 {len(other_months)} 条')
                if not dates:
                    errors.append(f'{sheet} 未抓取到任何新闻')
                if other_months:
                    errors.append(f'{sheet} 含非 {args.month} 记录')
            wb.close()
        except Exception as e:
            errors.append(f'{xlsx.name} 打开失败: {e}')
    if not docx_files:
        errors.append('未生成docx（可能是正文全部失败或没有关键词命中）')
    for docx in docx_files:
        try:
            doc = Document(docx)
            title_count = sum(1 for p in doc.paragraphs if p._p.xpath('./w:pPr/w:outlineLvl'))
            char_count = sum(len(p.text) for p in doc.paragraphs)
            info = {'file': docx.name, 'title_count': title_count, 'characters': char_count}
            result.setdefault('word_details', []).append(info)
            print(f'{docx.name}: {title_count} 篇；文本 {char_count} 字')
            if char_count < 100:
                errors.append(f'{docx.name} 内容不足')
        except Exception as e:
            errors.append(f'{docx.name} 无法读取: {e}')
    result['warnings'] = errors
    result['success'] = not errors
    summary = folder / '抓取检查报告.json'
    summary.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    md = [f'## {args.month} 新闻抓取结果', '', '| 来源 | 新闻数 | 姓名命中数 |', '|---|---:|---:|']
    for name, info in result['sheets'].items():
        md.append(f"| {name} | {info['news']} | {info['matched']} |")
    md += ['', f"生成的 Excel: {', '.join(result['xlsx']) or '无'}", f"生成的 Word: {', '.join(result['docx']) or '无'}"]
    if errors:
        md += ['', '**检查问题：**', *[f'- {e}' for e in errors]]
    if os.environ.get('GITHUB_STEP_SUMMARY'):
        with open(os.environ['GITHUB_STEP_SUMMARY'], 'a', encoding='utf-8') as out:
            out.write('\n'.join(md) + '\n')
    print('\n'.join(md))
    return 0 if not errors else 1

if __name__ == '__main__':
    raise SystemExit(main())
