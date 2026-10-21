# 九江新闻网 · 市委大事记整理工具

只采集 **九江新闻网「九江时政」**。运行时从公开 JSON 接口获取新闻列表及正文，根据本地 `names.txt` 姓名名单（或旧版 `config.json` 名单）筛选，生成 Excel、Word，使用 SQLite 缓存并对 Word 正文去重。**不采集九江市政府网站**，也不会自动添加外部政要姓名。

## Windows：在本地电脑运行和测试

需要安装 Git 和 Python **3.10 或以上**（建议 3.12）。以下命令在 **PowerShell** 中执行。可以把 `D:\PythonProjects` 换成自己准备使用的文件夹（需事先创建）。

~~~powershell
cd "D:\PythonProjects"
git clone https://github.com/anndigeer/chronicle-of-major-events.git
cd .\chronicle-of-major-events

# 创建虚拟环境、安装依赖
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

# 第一次使用：复制一份个人名单并在记事本打开
Copy-Item .\names.example.txt .\names.txt
notepad .\names.txt

# 保存名单后进行离线自检
.\.venv\Scripts\python.exe main.py --test

# 实际采集2026年9月，并导出Excel、Word
.\.venv\Scripts\python.exe main.py 2026-09
~~~

**不需要执行 `Activate.ps1`**，直接使用虚拟环境中的 `python.exe`，避免 PowerShell 执行策略拦截。如果输入 `py` 显示未找到命令，可尝试 `python -m venv .venv`（需要已安装 Python 并添加到 PATH）。

以后进入项目目录，可以直接运行：

~~~powershell
.\.venv\Scripts\python.exe main.py
~~~

不指定月份时，自动处理**上一个自然月**。结果都在项目的 `results` 文件夹：

- `jiujiang_news.xlsx`：只有“九江新闻网”工作表；包含新闻标题、日期、命中姓名、人工选择、原文链接与正文状态。
- `jiujiang_cache.sqlite3`：已提取正文的缓存，后续重复运行无需重新下载成功的正文。
- `市委大事记（YYYY年M月）.docx`：按日期排序并对重复报道去重。

首次运行要联网访问九江新闻网；`--test` 只执行模拟数据的离线测试，**不代表已验证你电脑到新闻网站的实际连接**。

## 跨月迟发报道（默认延后 10 天）

制作 **2026 年 9 月**大事记时，程序实际扫描的新闻**发布日期**为：

- **2026-09-01 至 2026-10-10（含）**；不采集 10 月 11 日以后的新闻。
- 9 月发布且找不到明确事件日期的新闻，仍暂按发布日期归入 9 月。
- 10 月 1—10 日发布的新闻会写入 Excel，但**仅当标题或正文开头出现明确的 9 月事件日期**（如“9月30日，陈云……调研”“昨日，邓永翔主持……”），才自动列入 9 月 Word。
- 真正发生在 10 月的会议、只提到“9 月”但没有明确事件日期的新闻，不会被自动归入 9 月。日期无法确认的仍留在 Excel 供人工核对。

Excel 新增 **I 列「事件日期」**，A—H 列的人工操作位置不变。对于需要补录的迟发报道，在 I 列填实际日期（例如 `2026-09-30`），E 列填“收录”，保存关闭 Excel，再执行：

~~~powershell
.\.venv\Scripts\python.exe main.py 2026-09 --export
~~~

**B 列始终是新闻原始发布日期，I 列才是事件发生日期**，Word 按事件日期排序。自动识别是基于日期与活动动词的保守规则，不等于人工核实；对于“近日”“日前”等模糊表述，应手动确认。Word 仍按正文去重。

宽限天数可在 `config.json` 中设置 `late_publication_days`（默认 10；设为 0 则不扫描次月）。例如 2026 年 12 月会顺延扫描到 **2027 年 1 月 10 日**。

**时效提醒：** 如果在 10 月 5 日就运行 9 月大事记，10 月 6—10 日的新闻尚未发布，10 月 10 日后需再运行一次才能补全；不会预测未来新闻。

### 完整性与重复运行

- **Excel 安全保存**：程序先在 `results` 内生成完整的临时 XLSX，再替换原文件。若写入途中失败，原 Excel 不会被半成品覆盖；运行前自动备份仍会保留。
- **避免旧数据倒灌**：已有新闻的标题和发布日期以 SQLite 缓存、新闻接口为准；Excel 只用于补录缓存缺失的新闻，不会在下次启动时拿旧字段覆盖新数据。B 列同步真实发布日期，I 列人工核对的事件日期与 E 列“收录/排除”保持不变。


- 如果接口连续两页返回**完全一样的新闻链接列表**，会立即停止并报告“分页接口可能未正确翻页”，避免反复请求同一页。
- `config.json` 的 `max_pages`、`timeout_seconds`、`delay_seconds`、`names` 等配置会在请求前检查，错误配置不会拖到采集途中才报错。
- 目标月份没有符合条件且正文完整的新闻时，程序返回失败状态；已选文章有正文下载失败时，即使生成了可用的部分 Word，也会返回失败状态并在终端提示“可能缺少文章”，避免把部分成果误认为完整月报。


- 自动推断的 I 列事件日期被清空后，会标为待人工复核。以后在 I 列重新填写有效日期时，复核批注会自动清除。手动输入不可能存在的日期（例如 `2026-09-31`）时，该行暂不进入 Word，Excel 的 H 列会指出错误。
- 九江新闻网接口即使返回 HTTP 200，只要新闻 JSON 结构无法识别，也会判为采集失败，而不是误认为已经翻完最后一页。这样可避免站点接口改版后生成缺失内容的 Word。

- 程序会继续翻页直到确实越过目标日期范围，**不会因为连续几页已经缓存就提前停止**。这有助于补齐此前遗漏的新闻。
- 如果任何分页网络请求失败，或到达 `config.json` 的 `max_pages` 仍未扫完目标月份，终端将明确报告采集不完整，退出代码为 **1**；已经拿到的新闻列表会保存到 Excel/SQLite，但**不生成新的 Word**。检查网络后重复运行即可继续补齐。
- 自动推断的 I 列事件日期带有“采集程序”的单元格批注。重复运行时，未经过人工修改的推断日期可随正文更新；**人工改写的事件日期会保留**。如果不认可自动日期，可以清空该单元格进入人工复核状态，或直接填写正确的事件日期。
- 这一套规则在 Windows、Linux 都使用相同的 `python main.py YYYY-MM` 命令。

## 简化后的命令

| 命令 | 用途 |
|---|---|
| `python main.py` | 默认采集上个月并导出 |
| `python main.py 2026-09` | 采集指定月份，生成 Excel 和 Word |
| `python main.py --test` | 离线运行全部单元与模拟采集测试 |
| `python main.py 2026-09 --sync` | 同步新闻和正文缓存，不生成 Word |
| `python main.py 2026-09 --export` | 不更新新闻列表；重新导出 Word（必要时补抓正文） |
| `python main.py 2026-09 --verify` | 离线检查该月已生成的 Excel 与 Word |
| `python main.py --help` | 查看所有参数 |

> 表中的 `python` 可替换为 Windows 的 `.\.venv\Scripts\python.exe`。没有额外的年份、网站、数据库路径或工作表参数；翻页数、延迟、超时和姓名等在 `config.json` 中配置。

## 批量录入 30 多个人名：只需粘贴，不用改 JSON

**推荐使用 `names.txt`**，放在 `main.py` 同一个目录；它是你电脑上的私人名单，已加入 `.gitignore`，以后更新代码不会修改它。首次在 PowerShell 执行：

~~~powershell
Copy-Item .\names.example.txt .\names.txt
notepad .\names.txt
~~~

在 Excel/WPS 中把所有姓名放在**一列**（不需要序号和职务），选中整列中的人名，按 `Ctrl+C`，回到记事本 `Ctrl+V` 粘贴并 `Ctrl+S` 保存。每行一个：

~~~text
张三
李四
欧阳修
阿布都·拉西提
~~~

程序也接受一行输入：`张三、李四，欧阳修;阿布都·拉西提`，支持换行、制表符、中文/英文逗号、顿号、分号。空行、以 `#` 开头的注释、Excel 表头“姓名”会被跳过；**重复姓名自动去重**，保留粘贴顺序。带职务、序号、异常字符的内容会提示行号，不会静默加入名单。注意姓名中间的 `·` 不会被当作分隔符。

**优先级：** 只要 `names.txt` 存在，就以其全部姓名为准，不与 `config.json` 里的 `names` 混合；如删掉 `names.txt`，仍兼容旧版 JSON 名单。文件存在但为空会明确报错，避免意外继续使用旧名单。请使用 UTF-8 文本，Windows 记事本的 UTF-8（包括 BOM）均支持。

更换名单后无需重新下载已缓存的新闻，可以直接运行：

~~~powershell
.\.venv\Scripts\python.exe main.py 2026-09 --export
~~~

程序会按最新名单重新判断 Excel 的 C/D 列命中情况，并重新生成 Word；**E 列人工收录/排除、I 列事件日期不会因此被覆盖**。若首次尚未采集，则直接运行 `python main.py 2026-09`。

银河麒麟/Linux 同样可以执行 `cp names.example.txt names.txt`，然后用文本编辑器填写，一行一人。

## 修改姓名、人工筛选与历史记录

姓名优先从 `names.txt` 读取；若没有该文件则兼容 `config.json` 的旧名单。`include_body_matches` 决定是否在正文查找姓名，绝不会抓取外部政要名单。

Excel 的 E 列可填“收录”或“排除”；旧的黄色单元格标记仍表示人工收录。人工修改后，先保存并关闭 Excel，再执行 `python main.py 2026-09 --export`。Word 以正文为依据判断重复：即使标题不同，正文实质相同也合并；标题相同但报道内容不同则分别保留。

旧版本的 `jiujiang_news.xlsx` 和 `jiujiang_cache.sqlite3` 如需继续使用，请在第一次运行前**分别备份并复制到 `results`**。程序会再为已有 Excel 自动创建 `results/备份/` 备份；旧政府工作表不会进入新 Excel，政府来源的 SQLite 数据也不会被读取或抓取。

## 银河麒麟 MIPS64 Linux

项目仍可在兼容的 Python 3.10+ 环境中运行。MIPS64 平台上 `lxml` 可能缺少预编译 wheel，优先从麒麟软件仓库安装 `python3-lxml`，如需源码编译则须提供 libxml2/libxslt 开发包和 C 编译器。

~~~bash
git clone https://github.com/anndigeer/chronicle-of-major-events.git
cd chronicle-of-major-events
python3 -m venv --system-site-packages .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py --test
python main.py 2026-09
~~~

`--system-site-packages` 用于复用系统可能已经提供的 MIPS64 lxml。**GitHub Actions 的 Linux/Windows x86_64 测试不等于 MIPS64 原生验证**。

项目仅保留两个 Python 文件：`main.py`（采集、导出与验证）和 `test_main.py`（离线测试）；GitHub Actions 仅在手动运行“News collection”时才进行真实新闻采集。
