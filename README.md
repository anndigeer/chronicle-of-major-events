# 九江新闻网 · 市委大事记整理工具

**新闻文章仅采集九江新闻网「九江时政」**；领导人名单另可从大江网江西领导页面自动更新。运行时从公开 JSON 接口获取新闻列表及正文，优先根据本地 `names.xlsx`、其次 `names.txt` 或旧版 `config.json` 中的姓名名单筛选，生成 Excel、Word，使用 SQLite 缓存并对 Word 正文去重。**不采集九江市政府网站**，也不会自动添加外部政要姓名。

## Windows：在本地电脑运行和测试

需要安装 Git 和 Python **3.10 或以上**（建议 3.12）。以下命令在 **PowerShell** 中执行。可以把 `D:\PythonProjects` 换成自己准备使用的文件夹（需事先创建）。

~~~powershell
cd "D:\PythonProjects"
git clone https://github.com/anndigeer/chronicle-of-major-events.git
cd .\chronicle-of-major-events

# 创建虚拟环境、安装依赖
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

# 推荐将原始Excel的姓名表保存为项目目录中的names.xlsx
# 检查实际生效的人名及数量
.\.venv\Scripts\python.exe main.py --check-names

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

**B 列始终是新闻原始发布日期，I 列才是事件发生日期**，Word 按事件日期排序。自动识别是基于日期与活动动词的保守规则，不等于人工核实；现已进一步区分新闻报道日期与随后出现的会议活动日期，并兼容“前日”“前天”“本月某日”“上月某日”等写法。对于“近日”“日前”等模糊表述，或多个会议日期同时出现的复杂报道，应手动确认。Word 仍按正文去重。

宽限天数可在 `config.json` 中设置 `late_publication_days`（默认 10；设为 0 则不扫描次月）。例如 2026 年 12 月会顺延扫描到 **2027 年 1 月 10 日**。

**时效提醒：** 如果在 10 月 5 日就运行 9 月大事记，10 月 6—10 日的新闻尚未发布，10 月 10 日后需再运行一次才能补全；不会预测未来新闻。

## Excel 长期存储与链接、备份策略

**使用一个工作簿中的一张“九江新闻网”工作表，不按月份拆分 Sheet。**
每次运行 `python main.py 2026-09`、`python main.py 2026-10` 等，
新文章依据其**发布日期 B 列**增量追加到同一个 `results/jiujiang_news.xlsx`：
以原文链接去重，已经存在的行不移动、不再次插入。遇到 9 月事件在 10 月发布，
只存一条 10 月发布日期记录，**I 列事件日期**决定它是否进入 9 月 Word。
若需查看单个月份，在 Excel 的 B 列使用自动筛选或排序；
程序不会强制重排你已经人工修改的行，因此补抓更早的月份时新行仍会追加在底部。

每条新闻 **A 列标题的实际超链接目标必须与 G 列“原文链接”完全相同**。
再次同步或导出时会自动修正历史行：优先使用数据库已记录的有效新闻链接，
如果两个不同链接都有效，会先根据 A 列可见标题与 SQLite 缓存中的新闻标题进行**唯一精确匹配**。有唯一依据才自动修复 A/G；没有唯一依据时，**不强行选择、不覆盖原链接**，将该行列入 `results/链接冲突待复核.csv`。CSV 始终使用固定文件名，不会逐月增生文件。程序会提示异常并以退出代码 1 结束，避免把可能漏稿的 Word 当作完整成果；已成功处理的其他记录仍会保存。
待复核 CSV 列出 Excel 行号、新闻标题、A 列链接、G 列链接和原因。核实后请在 Excel 手动将 A/G 指向同一篇正确新闻，再关闭 Excel 重新执行同步或导出，报告会自动更新；所有冲突解决后报告将删除。
仅接受九江新闻网域名的新闻 HTML 链接；不会把其他网站误写进去。
A 列的颜色/黄色标记以及 E 列人工选择、I 列人工日期均不会因此被清除。

程序运行前的 Excel 备份现在是**每天最多一份，滚动只保留最近 7 份自动备份**。
老版本已积累的同类“运行前”备份也会按照该规则清理；
仅识别形如“jiujiang_news_运行前_日期_时间_微秒.xlsx”的自动文件，
**不删除自行命名的手动备份，也不清理 Word 备份**。
需要永久保留的历史版本请另存为一个独立文件名。
如果同一天人工做了重要修改，建议手动另存一份；当天不会重复创建自动快照。

### 累计 Excel 质量检查与每月统计

使用**同一个** `results/jiujiang_news.xlsx`、**同一个**“九江新闻网”工作表，
新闻按发布日期增量追加。B 列筛选可查看任意月份的报道；不再每月创建新表，也不移动旧行，
从而保留 E 列人工收录/排除以及 I 列事件日期。

Windows 本地可执行以下**离线、只读**检查：

~~~powershell
.\.venv\Scripts\python.exe main.py --check-excel
~~~

程序统计每个月已累计的新闻条数，并检查 A 列超链接与 G 列原文链接是否一致、
重复原文链接、缺少有效链接、无效发布日期和事件日期等。没有异常返回代码 0；
有异常会打印**具体行号**并返回代码 1；**不会修改文件或创建数据库**。
若发现 A/G 不一致，可先关闭 Excel，执行一次正常同步或
`python main.py YYYY-MM --export`，程序会尝试按已缓存的可信链接修复；
再用 `--check-excel` 复核。存在无法判断的冲突时，请人工核对原文链接。

自动 Excel 运行前备份每天最多一份、只保留最近 7 份。
现在 Word 的自动备份也采用**每天最多一份、最近 7 份**，
不会清理不符合自动备份命名规则的手动存档。
自动备份是防误操作辅助，不代替你自己的异地备份。

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
| `python main.py --check-names` | 离线检查生效名单及人数，不联网、不修改文件 |
| `python main.py --check-excel` | 离线统计各月新闻数，检查 A/G 链接、重复新闻和日期 |
| `python main.py --update-leaders` | 联网更新四个板块领导名单至 names.auto.txt |
| `python main.py --help` | 查看所有参数 |

> 表中的 `python` 可替换为 Windows 的 `.\.venv\Scripts\python.exe`。没有额外的年份、网站、数据库路径或工作表参数；翻页数、延迟、超时和姓名等在 `config.json` 中配置。

## 批量录入 30 多个人名：直接保存 Excel（推荐）

在 Windows 的 Excel 或 WPS 中打开原始通讯录，保留**第一个工作表**。
把需要匹配的人名所在列的表头设为**「姓名」**，将文件另存为
**names.xlsx**，放到与 main.py 同一级的项目目录。无需逐一输入、复制粘贴或修改 JSON。

程序扫描第一张工作表前 10 行，自动识别「姓名」列，
所以姓名在 A、B、C 等任意列都可以。部门、职务等其他列不会用于匹配；
没有“姓名”表头时，默认读取 A 列。空行跳过、重复姓名去重且保留顺序。
若混入序号、职务、多个姓名或无效字符，会提示对应 Excel 单元格（例如 B17）。
请先保存为 .xlsx，旧 .xls 需在 Excel/WPS 中另存为 .xlsx。

| 序号 | 姓名 | 部门 |
|---|---|---|
| 1 | 陈云 | 部门甲 |
| 2 | 邓永翔 | 部门乙 |
| 3 | 黄强 | 部门丙 |

**先执行离线名单检查命令：**

~~~powershell
.\.venv\Scripts\python.exe main.py --check-names
~~~

会显示“姓名来源：names.xlsx”“生效姓名：N 人”和全部人名。
该命令**不联网、不创建数据库、不修改 Excel**，建议每次替换 30 多人名单后先运行。

### 仍可用 names.txt 批量粘贴

从 Excel 复制一列人名到记事本，保存为 UTF-8 的 names.txt。
首次可使用以下命令创建空白示例，再修改内容：

~~~powershell
if (-not (Test-Path .\names.txt)) { Copy-Item .\names.example.txt .\names.txt }
notepad .\names.txt
~~~

支持一行一人或“陈云、邓永翔，黄强;陈永奇”这种分隔形式，
自动跳过空行、注释、姓名表头和重复名字。

**姓名匹配精度：** 如果名单同时包含“陈云”和“陈云飞”，而文章只出现“陈云飞”，程序不会把长姓名内部的“陈云”片段另算成独立命中；如果正文其他位置又单独出现“陈云”，仍会命中两人。运行 `--check-names` 时会提示存在包含关系的姓名，方便人工检查。这个规则只能解决**名单内部**的包含冲突；若名单只有“陈云”，文章出现未录入名单的“陈云飞”时，程序仍可能命中“陈云”，请通过完整名单避免这种情况。

**人工来源优先级：names.xlsx → names.txt → config.json 原有名单。**
人工来源三者**不合并**；但会额外加入来自公开页面的 `names.auto.txt`，并去重。两个个人文件同时存在时以 names.xlsx 为准。
如果优先文件有错误，程序直接提示，不会默默使用旧名单。
本地 names.xlsx、names.txt 都由 .gitignore 排除，不会随 git pull 覆盖。
银河麒麟/Linux 同样支持。

更换名单后，缓存中的正文无需重新下载：

~~~powershell
.\.venv\Scripts\python.exe main.py 2026-09 --export
~~~

程序按新名单刷新 Excel 的 C/D 列和 Word，保留 E 列人工选择与 I 列事件日期。
还没有采集时正常运行 python main.py 2026-09。

## 从大江网自动更新“省委、人大、政府、政协”领导名单

程序另支持从 [江西领导页面](https://jiangxi.jxnews.com.cn/leader.shtml)
指定的 `table[width="98%"][border="0"][align="center"][cellpadding="0"][cellspacing="5"]`
（样式为 `line-height:30px`）提取四个板块中的人名。此页面**只用于更新姓名关键词**，
新闻报道数据源仍仅为九江新闻网；不会提取“人事任免”报道标题。
会规范化姓名中的全角空格，并在匹配时对不同板块的重复姓名去重。

手动联网更新：

~~~powershell
.\.venv\Scripts\python.exe main.py --update-leaders
~~~

成功后在项目目录生成或更新 **`names.auto.txt`**，按四个板块分组保存。
`names.auto.txt` 属于**可追踪的公开来源名单**，通过 GitHub 版本历史可查看变化。
它与个人名单 `names.xlsx` / `names.txt` **分开保存**：
定时更新只修改自动名单，不覆盖个人文件。

实际匹配名单规则是：

1. 人工部分优先读取 `names.xlsx`，否则读取 `names.txt`，否则使用旧 JSON 名单。
2. 自动部分来自 `names.auto.txt`；**人工名单与自动名单合并、去重**。
3. 若有自动名单但没有人工名单文件，自动名单直接生效，默认示例人名不会额外混入。
4. 无论更新成功与否，你都可以运行 `python main.py --check-names` 查看最终生效姓名。

更新前会验证四个板块均存在姓名，识别失败、网络异常或网页结构变化时
命令以失败状态退出，**不会清空或覆盖先前成功的 names.auto.txt**。
正常情况下如果名单未变化也不会重复写文件。
网页上公布的领导名单不一定实时反映所有职务变化，需要以相关主管机关发布的信息为准。

### 首次名单快照与实际联网限制

仓库已附带一个**初始网页缓存快照** `names.auto.txt`，
按省委、人大、政府、政协分组记录可见名单（四组共 35 个席位，人名跨板块去重后 32 人）。
**它不是脚本刚刚联网抓取的结果，也不能代替实时核验**；
运行 `--check-names` 会提示这一点。后续任何一次 `--update-leaders`
成功后，程序会根据目标网页重新生成该文件并清除快照标记。

首次 GitHub Actions 实际测试已发现托管 Runner 访问该站点时出现
TLS 握手失败，HTTP 访问也不能取得网页，因此当时的自动更新**没有成功**。
项目不会关闭证书验证，也不会在抓取失败时把缓存快照当成最新数据。
你可以从可访问网站的 Windows 电脑或银河麒麟服务器执行
`python main.py --update-leaders`，随后运行 `--check-names` 核实。

### 在 Windows 上设置本地每周更新

在项目根目录 PowerShell 中执行（Windows 计划任务运行时，电脑需要开机且联网）：

~~~powershell
$project = (Get-Location).Path
$python = (Resolve-Path .\.venv\Scripts\python.exe).Path
$script = (Resolve-Path .\main.py).Path
$action = New-ScheduledTaskAction -Execute $python -Argument ('"' + $script + '" --update-leaders') -WorkingDirectory $project
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At (Get-Date -Hour 10 -Minute 15 -Second 0)
Register-ScheduledTask -TaskName "JiangxiLeadersWeekly" -Action $action -Trigger $trigger -Description "每周更新江西省领导人名名单"
~~~

如果 Windows 提示没有创建任务的权限，可以在“任务计划程序”中
手动创建同名任务：程序设为虚拟环境的 python.exe，参数设为
main.py 的完整路径加 `--update-leaders`，起始目录设为项目根目录。

### 在银河麒麟上用 cron 每周更新

先确保已安装并启用系统 cron 服务，执行 `crontab -e`，
以下路径替换为项目在服务器上的实际绝对路径（服务器时区决定执行时间）：

~~~cron
15 10 * * 1 /home/user/chronicle-of-major-events/.venv/bin/python /home/user/chronicle-of-major-events/main.py --update-leaders >> /tmp/jiangxi-leaders.log 2>&1
~~~

本地计划任务与 GitHub 每周任务相互独立；如果 GitHub Runner 连接不上网站，
本地任务能否成功取决于本机网络，**不能保证一定成功**。

### 定期自动更新：GitHub Actions

新增 `.github/workflows/leader-names.yml`，默认**每周一北京时间上午 10:15**
（UTC 02:15）运行；也可在 GitHub **Actions → Update Jiangxi leaders → Run workflow**
随时手动触发。

工作流用 GitHub 自带权限获取页面、校验四个板块、更新 `names.auto.txt`，
并仅在内容变化时向 `main` 提交一次；此提交会触发离线 CI。
首次新增工作流的提交也会触发一次自动更新。
**GitHub 定时任务运行于 GitHub Runner，而不是你的 Windows 或银河麒麟服务器**；
使用 `git pull origin main` 即可把更新后的自动名单拉到本地。
网站连不上、存储库权限限制或定时工作流被 GitHub 停用时，需要检查 Actions 结果，
不能视为已经成功更新。定时任务可能有延迟。

本地不必定时联网抓姓名；若希望在不连接 GitHub 的服务器上保持更新，
可以定期执行 `python main.py --update-leaders`，不会改变本地的个人名单文件。

## 修改姓名、人工筛选与历史记录

人工名单以 `names.xlsx` 为优先，其次是 `names.txt`，最后是 `config.json` 旧名单；如果 `names.auto.txt` 已存在，则与人工姓名去重合并。`include_body_matches` 决定是否在正文查找姓名，绝不会抓取外部政要名单。

Excel 的 E 列可填“收录”或“排除”；旧的黄色单元格标记仍表示人工收录。人工修改后，先保存并关闭 Excel，再执行 `python main.py 2026-09 --export`。Word 以正文为依据判断重复：即使标题不同，正文实质相同也合并；标题相同但报道内容不同则分别保留。

旧版本的 `jiujiang_news.xlsx` 和 `jiujiang_cache.sqlite3` 如需继续使用，请在第一次运行前**分别备份并复制到 `results`**。程序会为已有 Excel 创建每日最多一份的 `results/备份/` 自动备份（保留最近 7 份）；旧政府工作表不会进入新 Excel，政府来源的 SQLite 数据也不会被读取或抓取。

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
