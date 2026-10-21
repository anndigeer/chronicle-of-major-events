# 九江新闻网 · 市委大事记自动整理

**唯一数据源：九江新闻网「九江时政」栏目**。通过公开 JSON 接口增量采集新闻；
不连接九江市政府，不获取政府 XML，也不会自动抓取外部政要名单。

功能：SQLite 缓存正文、config.json 的姓名自动筛选、Excel 人工收录/排除、
正文去重、按月份导出 Word。Excel 仅有「九江新闻网」工作表。

## 银河麒麟 MIPS64 Linux 安装

先检查服务器架构、Python 和包管理器：

~~~bash
uname -m
python3 --version
command -v apt || command -v dnf || command -v yum
~~~

建议 Python 3.10+。包依赖包括 python-docx，而 python-docx 依赖带 C 扩展的
**lxml**。在 MIPS64 平台，PyPI 不一定有适配 wheel；优先从麒麟的软件仓库安装
python3-lxml。若系统使用 APT，参考：

~~~bash
sudo apt update
sudo apt install -y python3-pip python3-venv python3-lxml
~~~

如果使用 DNF / YUM，则从对应软件仓库查询并安装 python3-lxml。
如需自行编译 lxml，还要安装 Python 开发头文件、C 编译器及
libxml2/libxslt 开发包；不能假设所有麒麟 MIPS64 版本都能直接从 PyPI 安装。

~~~bash
git clone https://github.com/anndigeer/chronicle-of-major-events.git
cd chronicle-of-major-events
python3 -m venv --system-site-packages .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -c "import lxml, requests, bs4, openpyxl, docx; print('依赖可用')"
python -m unittest -v test_main
~~~

虚拟环境使用 --system-site-packages 是为了复用系统提供的 MIPS64 lxml。
若仍触发 lxml 编译或缺少依赖，需检查系统 Python 与软件包版本是否匹配。

## 采集与导出

~~~bash
source .venv/bin/activate
python main.py --month 2026-09 --only-month --content-scope month
python main.py --mode sync --month 2026-09 --only-month
python main.py --mode export --month 2026-09
python main.py --help
~~~

默认月为上个月。生成：jiujiang_news.xlsx、jiujiang_cache.sqlite3、
市委大事记（YYYY年M月）.docx。主程序不会自动在后台执行；
如需定期运行请在服务器自行配置 cron。

Excel 的 E 列可填「收录」或「排除」，旧黄色标记仍表示人工收录。
配置文件 config.json 中的 names 是唯一自动匹配名单，include_body_matches
决定是否搜索正文。更改名单后使用 export 模式重新筛选。

Word 去重：正文一致但标题不同会合并；标题相同、日期相近时，还需正文高度相似。
如果标题相同但内容不同，则保留多篇。Excel 保留原始新闻记录。

## 旧数据与安全

**运行前会自动备份已有 Excel 到「备份」目录**。旧双工作表 Excel 中的
「九江新闻网」表格（包括手工选择和黄色标记）会保留；其他旧表不再进入新版 Excel，
但在原始备份中保留。旧 SQLite 中的政府来源记录不会被读取、请求或写入 Word。
首次迁移前仍建议自行另存原始 Excel/SQLite 文件。

仅通过手动 GitHub Actions 运行联网采集；普通代码提交的 Python CI 只运行离线测试。
CI 使用 x86_64 Linux，并不等于已在银河麒麟 MIPS64 原生设备验证。
请在实际服务器执行上述测试与一次月度采集。

项目源码：main.py、source_feeds.py；配置：config.json、requirements.txt；
测试：test_main.py；可选月度验证：scripts/verify_month_output.py。
