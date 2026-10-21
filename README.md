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

- `python -m unittest discover -v -p 'test_*.py'`：完全离线的 6 项回归测试，包括模拟的完整运行、改版列表解析、正文提取、关键词过滤、重复文章合并、旧 Excel 黄色标记兼容和 Word 生成。
- `python jrjj_v2.py --help`：命令行可用性检查。
- [GitHub Actions · Python CI](.github/workflows/python-ci.yml)：提交和 PR 自动在 Ubuntu（Python 3.10、3.12）和 Windows（Python 3.12）运行离线测试。
- 需要实际检查网站时，进入 **Actions → Python CI → Run workflow**，勾选 `live_smoke`。联网测试仅下载两个栏目首页及每站一篇样本正文，网页访问受限或 HTML 改版时会失败并显示原因，不会把访问失败误报为通过。

> GitHub Actions 默认**不会自动批量爬取新闻、修改工作簿或提交生成的文档**；本项目不会把用户实际的工作簿、大事记文件、正文缓存或调试 HTML 推送到公共仓库。

## 数据安全

仓库只提交程序、配置示例、说明和测试代码；`.gitignore` 排除个人 Excel、Word、数据库、备份及调试 HTML。不要将 GitHub PAT 放入代码、配置或 Actions YAML；CI 只使用 GitHub 提供的最小读取权限。

## 持续集成验证

在 GitHub Actions 的 **Python CI** 中检查 Ubuntu Python 3.10 / 3.12 和 Windows Python 3.12 的离线回归测试。网站连通性测试默认不运行，可在手动启动工作流时勾选 `live_smoke`。初次上传各文件时触发的早期失败记录不代表最新版本的状态，应查看最新提交对应的运行。
