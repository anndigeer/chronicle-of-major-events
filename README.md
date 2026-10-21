# 九江市委大事记自动整理工具

从九江市政府“今日九江”和九江新闻网“九江时政”增量采集公开新闻，缓存正文，根据可编辑人名表自动筛选，导出双工作表 Excel 和按月排序的 Word 大事记。人工在 Excel 的 E 列选择“收录/排除”或者沿用黄色标记覆盖自动筛选。

**完整用法**见 [README_使用说明.md](README_使用说明.md)。

## 快速开始

建议 Python 3.10 以上。

```bash
python -m pip install -r requirements_v2.txt
python jrjj_v2.py                         # 一键同步 + 导出上个月
python jrjj_v2.py --mode sync             # 只采集和更新 Excel
python jrjj_v2.py --mode export           # 人工复核后重新生成 Word
python jrjj_v2.py --month 2026-09         # 指定月份
```

把旧 `jiujiang_news.xlsx` 放到程序所在目录就能迁移历史记录。名单等参数可在 [jrjj_config.json](jrjj_config.json) 调整。

## 测试 / GitHub Actions

- 离线测试：`python -m unittest discover -v -p 'test_*.py'`，涵盖 XML/JSON、正文、姓名过滤、去重、旧版 Excel 兼容等。
- [Python CI](.github/workflows/python-ci.yml)：正常提交只执行离线测试，不触发新闻抓取。
- [News collection](.github/workflows/news-collection.yml)：**仅手动**选择 `government_smoke` 或 `month_export`；前者针对政府 XML、配置姓名和正文做有限量诊断，后者按月份生成 Excel 和 Word。
- 政府 XML 的公网连接可能对 GitHub 托管 Runner 不可达；诊断失败报告会区分网络连接、HTTP 响应和解析问题。可将这个工作流配置到网络可达的自托管 Runner。

## 数据安全

仓库只提交程序、配置示例、说明和测试代码；`.gitignore` 排除个人 Excel、Word、数据库、备份及调试 HTML。不要将 GitHub PAT 放入代码、配置或 Actions YAML；CI 只使用 GitHub 提供的最小读取权限。

## 持续集成验证

在 GitHub Actions 的 **Python CI** 中检查 Ubuntu Python 3.10 / 3.12 和 Windows Python 3.12 的离线回归测试。网站连通性测试默认不运行，可在手动启动工作流时勾选 `live_smoke`。初次上传各文件时触发的早期失败记录不代表最新版本的状态，应查看最新提交对应的运行。

## 政府 XML 一次性连通性检查

提交信息包含 `[gov-xml-test]` 时，现有 Python CI 将额外执行一次政府年度 XML 和 2026 年 9 月目标姓名新闻正文的联网诊断。诊断报告及日志作为 Artifact 保存，若 TCP 连接失败会如实失败。平时的提交不会执行此联网任务。月度文件导出统一使用 `.github/workflows/september-2026-export.yml` 手动触发；不再保留重复的九月采集工作流和配置。
