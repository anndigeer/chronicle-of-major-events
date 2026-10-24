# 九江新闻网 · 市委大事记整理工具

仓库已附带可直接使用的 `names.txt`（无须从 `names.example.txt` 复制）。

**新闻文章仅采集九江新闻网「九江时政」**；领导人名单另可从大江网江西领导页面自动更新。运行时从公开 JSON 接口获取新闻列表及正文，优先根据本地 `names.xlsx`、其次 `names.txt` 或旧版 `config.json` 中的姓名名单筛选，生成 Excel、Word，使用 SQLite 缓存并对 Word 正文去重。**不采集九江市政府网站**，也不会自动添加外部政要姓名。

## 银河麒麟桌面版 · 龙芯 3A4000（MIPS64）网页图形界面

针对 **银河麒麟桌面系统 + 龙芯 3A4000**，项目增加了 `webui.py`：只依赖 Python **标准库**提供本地网页操作界面，工作台调用已有 `main.py` 完成新闻抓取、缓存、Excel 和 Word 导出，避免复制两套业务规则。服务**仅监听本机 `127.0.0.1`**，自动在默认浏览器打开页面，不对外提供网络服务。

工作台按钮包括：选择月份、增量采集并生成 Word、仅同步 Excel、按 E 列收录/排除重建 Word、导出结果验证、检查名单、检查 Excel、更新领导名单、打开结果目录和办公文件，以及实时查看执行日志。**任务不会同时执行两个**，避免多个采集进程争写同一份 Excel。

### 一个 Shell 脚本完成检查、安装与启动

银河麒麟现在统一使用 **`市委大事记.sh`**，不再维护三个独立启动/安装/快捷方式脚本。克隆项目后进入项目目录，先运行：

~~~bash
# 只读环境诊断，显示 Python、CPU 架构与每个依赖的最低/最高版本
sh ./市委大事记.sh check

# 满足采集要求后，打开本机浏览器图形界面
sh ./市委大事记.sh
~~~

检查会比较项目 `requirements.txt` 对应的版本范围，并对不满足条件的每个依赖打印**当前版本、需要的版本、单包更新命令**。例如 openpyxl 为 3.0.10，会提示：

~~~text
[待处理] openpyxl>=3.1,<4；当前：3.0.10；版本过低
单独安装/更新：sh ./市委大事记.sh install openpyxl
等价的虚拟环境命令：./.venv/bin/python -m pip install --upgrade 'openpyxl>=3.1,<4'
~~~

**依赖升级只在用户主动运行 install 后才进行：**

~~~bash
# 创建/复用项目的 .venv，并安装或更新所有不满足版本范围的依赖
sh ./市委大事记.sh install all

# 仅安装或升级一个指定依赖（其他包不主动升级）
sh ./市委大事记.sh install openpyxl

# 其他可指定的名字：requests、beautifulsoup4、python-docx、urllib3
# 安装完成后重新检测
sh ./市委大事记.sh check
~~~

安装前会显示目标软件包并要求输入 `y` 确认；默认拒绝、无终端时拒绝。**不使用 sudo、不修改系统 Python，只对本项目的 `.venv` 安装依赖**；必要时使用可用 Python 3.10+ 创建虚拟环境，失败时会提示处理对应的 venv/pip 组件。MIPS64 的某些原生依赖可能需要编译或使用系统匹配的软件包，因此不保证每台麒麟机器均能联网直接安装完成。不要使用 ARM/x86/LoongArch 安装包替代 MIPS64。

如果当前系统只有 Python 3.8/3.9，可以先运行 `sh ./市委大事记.sh` 打开诊断界面，但**采集和导出必须有 Python 3.10+**。脚本不会在版本不合格时强行调用新闻采集核心。升级系统 Python 应当从兼容的银河麒麟 MIPS64 官方软件源获得，统一脚本不会擅自替换系统解释器。

以后创建桌面快捷方式，只需运行一次：

~~~bash
sh ./市委大事记.sh desktop
~~~

它会添加「市委大事记」到用户应用菜单，并在存在桌面文件夹时创建快捷方式。之后点击图标即可启动；部分麒麟桌面需要右键设为“允许启动”。完整参数运行 `sh ./市委大事记.sh help`；安装包升级后请关闭并重新打开网页工作台，让界面重新检测运行环境。

### 在网页中完成常用工作

1. **选择目标月份**，点「增量采集 + 更新 Word」：抓取新增新闻并重新生成 Word。
2. 点「打开 Excel」，在 E 列设置「收录/排除」，在 H 列核实事件日期，保存并**关闭 Excel/Word**。
3. 点「按人工收录/排除重建 Word」，自动按最新人工选择覆盖旧版自动 Word（按既有规则备份）。
4. 点「打开本月 Word」或「打开结果目录」查看文件，遇到失败查看页面底部的日志。

**没有安装 Python 或第三方依赖的电脑能否使用？**分两种情况：同一台麒麟电脑要运行**本地采集服务**，仍需要兼容 MIPS64 的 Python 及业务依赖；仅作为**远程网页客户端**时，不需要安装 Python，但必须把采集后端部署到另一台电脑/服务器，并增加登录、访问控制、文件授权及 HTTPS。**本版仅提供本机网页服务，并未实现可直接公开部署的远程多用户系统。**

不要下载普通 Windows EXE 或 x86 Linux 程序在 3A4000 上运行。本项目当前**没有**经龙芯 3A4000 原生构建和验证的免 Python `.deb`/独立二进制发行包；如后续需要完全免装依赖的本地程序，应在同型号或兼容的 MIPS64 麒麟环境中完成打包和实机测试。

### GitHub Actions：自动测试图形界面与采集全流程

每次推送 `main`、提交 Pull Request，或在 GitHub 网页手动点击 **Actions → GUI and full-flow CI → Run workflow**，会运行以下自动检查：

| 测试任务 | 主要检查 |
|---|---|
| Windows/Linux 离线测试 | Python 3.10、3.12；脚本语法、新闻采集/导出、Excel 数据、人工收录/排除、本机网页服务接口安全 |
| 银河麒麟启动自检 | POSIX shell 脚本语法；包含中文和空格的项目目录启动；用户菜单/桌面快捷方式生成；只读架构与依赖诊断 |
| Chromium 真实浏览器端到端测试 | 用无头 Chromium 打开中文网页工作台，依次点击「增量采集 + 更新 Word」「按人工收录/排除重建 Word」「仅增量同步 Excel」「核验成果」「检查 Excel」「检查名单」 |
| E2E 数据断言 | 核对实际生成的 XLSX/DOCX、新闻正文、跨月事件日期；确保手动排除消失、手动收录出现、增量新闻不重复、`--sync` 不改 Word |
| 异常与恢复 | 全部排除时不会保留过期 Word；重新收录后可重新导出；核对生成的 Word 自动备份 |

Chromium E2E 使用**固定模拟新闻网列表和正文**，但是真实调用 `main.py` 的业务函数和 `webui.py` 的本机 HTTP 服务，并真实生成、读取 XLSX/DOCX。它不依赖外部新闻网站的短时可用性，失败时会将截图和运行日志上传到 GitHub Actions 的 **Artifacts → gui-e2e-results**，通常能直接定位是哪一步失败。

测试结果请以 **Actions 页面的任务状态** 为准：绿色成功表示这一套自动测试通过；红色失败时进入对应 Job 展开失败步骤下载日志。单纯把 workflow 文件提交到仓库并**不代表已经通过了 GitHub Actions 测试**。

**架构限制：** GitHub 的新增麒麟启动自检是对 POSIX shell 和桌面文件的**模拟检查**，在 x86_64 的 Ubuntu runner 上运行，并非 MIPS64 指令集仿真。GitHub 提供的 `ubuntu-latest`、`windows-latest` 测试环境并非龙芯 3A4000 的 MIPS64 银河麒麟机器。这些测试能够验证浏览器交互、Python 核心逻辑以及生成文件，但**无法替代 3A4000 原机测试**。最终仍需在麒麟机器执行一次 `sh ./市委大事记.sh`，重点确认该系统的 Python 版本、MIPS64 依赖包、办公软件关联和桌面启动器。

### 图形界面离线测试

~~~bash
python3 -m unittest -v test_webui
# 在依赖完整且 Python>=3.10 的环境下测试整个项目：
python3 main.py --test
~~~

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

- `jiujiang_news.xlsx`：只有“九江新闻网”工作表；包含新闻标题、日期、命中姓名、人工选择、正文状态及真实事件日期；A 列标题本身是唯一原文超链接。
- `jiujiang_cache.sqlite3`：已提取正文的缓存，后续重复运行无需重新下载成功的正文。
- `市委大事记（YYYY年M月）.docx`：按日期排序并对重复报道去重。

首次运行要联网访问九江新闻网；`--test` 只执行模拟数据的离线测试，**不代表已验证你电脑到新闻网站的实际连接**。

## 跨月迟发报道（默认延后 10 天）

制作 **2026 年 9 月**大事记时，程序实际扫描的新闻**发布日期**为：

- **2026-09-01 至 2026-10-10（含）**；不采集 10 月 11 日以后的新闻。
- 9 月发布且找不到明确事件日期的新闻，仍暂按发布日期归入 9 月。
- 10 月 1—10 日发布的新闻会写入 Excel，但**仅当标题或正文开头出现明确的 9 月事件日期**（如“9月30日，陈云……调研”“昨日，邓永翔主持……”），才自动列入 9 月 Word。
- 真正发生在 10 月的会议、只提到“9 月”但没有明确事件日期的新闻，不会被自动归入 9 月。日期无法确认的仍留在 Excel 供人工核对。

Excel 使用 **H 列「事件日期」**，原 G 列独立「原文链接」已经删除，原 H 列「说明」移至 G 列。对于需要补录的迟发报道，在 H 列填实际日期（例如 `2026-09-30`），E 列填“收录”，保存关闭 Excel，再执行：

~~~powershell
.\.venv\Scripts\python.exe main.py 2026-09 --export
~~~

**B 列始终是新闻原始发布日期，H 列才是事件发生日期**；记者获悉日期不单独显示或填入 H 列，Word 按事件日期排序。自动识别是基于日期与活动动词的保守规则，不等于人工核实；现已进一步区分新闻报道日期与随后出现的会议活动日期，并兼容“前日”“前天”“本月某日”“上月某日”等写法。对于“近日”“日前”等模糊表述，或多个会议日期同时出现的复杂报道，应手动确认。Word 仍按正文去重。

宽限天数可在 `config.json` 中设置 `late_publication_days`（默认 10；设为 0 则不扫描次月）。例如 2026 年 12 月会顺延扫描到 **2027 年 1 月 10 日**。

**时效提醒：** 如果在 10 月 5 日就运行 9 月大事记，10 月 6—10 日的新闻尚未发布，10 月 10 日后需再运行一次才能补全；不会预测未来新闻。

## Excel 长期存储与链接、备份策略

**使用一个工作簿中的一张“九江新闻网”工作表，不按月份拆分 Sheet。**
每次运行 `python main.py 2026-09`、`python main.py 2026-10` 等，
新文章依据其**发布日期 B 列**增量追加到同一个 `results/jiujiang_news.xlsx`：
以原文链接去重，已经存在的行不移动、不再次插入。遇到 9 月事件在 10 月发布，
只存一条 10 月发布日期记录，**H 列事件日期**决定它是否进入 9 月 Word。
若需查看单个月份，在 Excel 的 B 列使用自动筛选或排序；
程序不会强制重排你已经人工修改的行，因此补抓更早的月份时新行仍会追加在底部。

新表严格保留 **A:H 八列**：A 新闻标题（唯一可点击原文链接）、B 发布日期、C 自动匹配、D 命中姓名、E 人工选择、F 正文状态、G 说明、H 事件日期。**没有独立的 URL 列**；所有增量去重、正文获取、质量核对以 A 列实际链接为准。

第一次运行会将旧版 A:I 工作表迁移为 A:H：如果旧 A 列缺少有效超链接，但旧 G 列有合法网址，会自动补齐 A 的超链接；原 H 列说明与原 I 列事件日期自动前移，同时保留 E 列人工选择、黄色标记和事件日期批注。迁移前先创建自动 Excel 备份。**原来 A/G 两个地址不同的行不会丢失候选链接**：旧 G 地址会暂存于新版 G 列的「原文链接待核」备注，并在 `results/链接冲突待复核.csv` 中记录行号和两个候选地址。复核时在 Excel 中确认 A 列正确的超链接；若采用原 G 的链接，则把 A 的链接改为该地址，程序下次自动清除标记；若确认原 A 正确，则手动删除 G 列的「【原文链接待核】...」标记（不要删掉其他说明）。在标记解决前该行不参与 Word 自动收录，程序返回异常状态，防止误收录。

**E 列「人工选择」操作规则：**

| E 列 | 作用 |
|---|---|
| 留空 | 自动识别姓名并按规则决定是否候选收录 |
| 收录 | 强制将当前文章列入候选（即使没有命中姓名） |
| 排除 | 强制不进入 Word（即使命中姓名或黄色标记） |

优先级：**人工排除 ＞ 人工收录 ＞ 旧版黄色标记 ＞ 自动关键词命中**。设置「收录」也仍必须满足正确的事件月份、正文可用、文章未被去重等基本条件；两个选项都**不会删除 Excel 的新闻记录或 SQLite 正文缓存**。完成编辑后保存并关闭 Excel，运行 `python main.py YYYY-MM --export` 即可重新筛选生成 Word。

Word 中主标题「市 委 大 事 记」为大纲 **1 级**，副标题「（YYYY年M月）」单独成段，为大纲 **2 级**。新闻标题保留大纲 **3 级**。

程序运行前的 Excel 备份现在是**每天最多一份，滚动只保留最近 7 份自动备份**。
老版本已积累的同类“运行前”备份也会按照该规则清理；
仅识别形如“jiujiang_news_运行前_日期_时间_微秒.xlsx”的自动文件，
**不删除自行命名的手动备份，也不清理 Word 备份**。
需要永久保留的历史版本请另存为一个独立文件名。
如果同一天人工做了重要修改，建议手动另存一份；当天不会重复创建自动快照。

### 累计 Excel 质量检查与每月统计

使用**同一个** `results/jiujiang_news.xlsx`、**同一个**“九江新闻网”工作表，
新闻按发布日期增量追加。B 列筛选可查看任意月份的报道；不再每月创建新表，也不移动旧行，
从而保留 E 列人工收录/排除以及 H 列事件日期。

Windows 本地可执行以下**离线、只读**检查：

~~~powershell
.\.venv\Scripts\python.exe main.py --check-excel
~~~

程序统计每个月已累计的新闻条数，检查 A 列标题是否有合法原文超链接、重复链接、B 列发布日期、H 列事件日期以及 G 列历史冲突复核标记。没有异常返回代码 0；有异常会打印具体行号并返回代码 1；**不会修改文件或创建数据库**。

首次更新旧表前可执行此命令预检；随后关闭 Excel，运行一次同步或导出，完成迁移，再重新检查。

自动 Excel 运行前备份每天最多一份、只保留最近 7 份。
现在 Word 的自动备份也采用**每天最多一份、最近 7 份**，
不会清理不符合自动备份命名规则的手动存档。
自动备份是防误操作辅助，不代替你自己的异地备份。

### 完整性与重复运行

- **Excel 安全保存**：程序先在 `results` 内生成完整的临时 XLSX，再替换原文件。若写入途中失败，原 Excel 不会被半成品覆盖；运行前自动备份仍会保留。
- **避免旧数据倒灌**：已有新闻的标题和发布日期以 SQLite 缓存、新闻接口为准；Excel 只用于补录缓存缺失的新闻，不会在下次启动时拿旧字段覆盖新数据。B 列同步真实发布日期，H 列人工核对的事件日期与 E 列“收录/排除”保持不变。


- 如果接口连续两页返回**完全一样的新闻链接列表**，会立即停止并报告“分页接口可能未正确翻页”，避免反复请求同一页。
- `config.json` 的 `max_pages`、`timeout_seconds`、`delay_seconds`、`names` 等配置会在请求前检查，错误配置不会拖到采集途中才报错。
- 目标月份没有符合条件且正文完整的新闻时，程序返回失败状态；已选文章有正文下载失败时，即使生成了可用的部分 Word，也会返回失败状态并在终端提示“可能缺少文章”，避免把部分成果误认为完整月报。


- 自动推断的 H 列事件日期被清空后，会标为待人工复核。以后在 H 列重新填写有效日期时，复核批注会自动清除。手动输入不可能存在的日期（例如 `2026-09-31`）时，该行暂不进入 Word，Excel 的 G 列会指出错误。
- 九江新闻网接口即使返回 HTTP 200，只要新闻 JSON 结构无法识别，也会判为采集失败，而不是误认为已经翻完最后一页。这样可避免站点接口改版后生成缺失内容的 Word。

- 程序会继续翻页直到确实越过目标日期范围，**不会因为连续几页已经缓存就提前停止**。这有助于补齐此前遗漏的新闻。
- 如果任何分页网络请求失败，或到达 `config.json` 的 `max_pages` 仍未扫完目标月份，终端将明确报告采集不完整，退出代码为 **1**；已经拿到的新闻列表会保存到 Excel/SQLite，但**不生成新的 Word**。检查网络后重复运行即可继续补齐。
- 自动推断的 H 列事件日期带有“采集程序”的单元格批注。重复运行时，未经过人工修改的推断日期可随正文更新；**人工改写的事件日期会保留**。如果不认可自动日期，可以清空该单元格进入人工复核状态，或直接填写正确的事件日期。
- 这一套规则在 Windows、Linux 都使用相同的 `python main.py YYYY-MM` 命令。

## 同一个月份反复同步、人工收录与排除

假设 9 月 Word 已经生成，想把不重要的新闻删出 Word、把漏选新闻补进来：在 `results/jiujiang_news.xlsx` 的 **E 列**设置“收录”或“排除”，需要调整事件发生日时填写 **H 列**，保存并**关闭 Excel/Word**。

- **需要增量补抓网站新发布的报道，同时同步人工选择到 Word**：运行 `python main.py 2026-09`（推荐）。它会重新检查 **2026-09-01 至 2026-10-10** 的发布列表，仅把新的原文链接追加到同一个 Excel；成功下载的正文从 SQLite 复用，不重复下载。随后根据现有 E 列重新生成该月 Word；被人工“排除”的报道不再出现在新 Word 中。
- **只修改了 E 列，不需要联网刷新发布列表**：运行 `python main.py 2026-09 --export`。它不查询新闻列表，但此前正文缺失的报道仍可能联网补抓；若正文缓存齐全，直接用缓存重建 Word。
- **只想获取新新闻而暂不改 Word**：运行 `python main.py 2026-09 --sync`。此命令**不会**应用排除操作到旧 Word，之后还需要 `--export`，或者直接执行不带参数的完整月度命令。

“排除”指**从生成的 Word 中剔除**，不会删除 Excel 中的新闻行或 SQLite 缓存，这样日后随时可以重新设为“收录”。仅需要剔除部分文章时，已有自动生成的 Word 会先备份，然后被新版本替换。如果该月所有候选都被排除（或最终没有可导出正文），旧版**自动生成**的 Word 也会先备份、再清除以防误用，脚本以退出代码 1 提示当月没有可导出正文；**人工自行创建的 Word 不删除**。所有人工修改都只会在**下一次成功的 Word 导出**中体现，故不能只运行 `--sync`。

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
| `python main.py --check-excel` | 离线统计各月新闻数，检查 A 列原文链接、重复新闻和日期 |
| `python main.py --update-leaders` | 联网更新四个板块领导名单至 names.auto.txt |
| `python main.py --help` 或 `python main.py -h` | 显示每项参数的具体用途、联网和文件影响、操作举例及常见流程 |

### 命令行详细帮助（直接在终端查看）

运行 `python main.py --help` 或 `python main.py -h`，会在 PowerShell / Linux 终端显示完整的中文使用说明，不需要另开文档。**无需联网，也不会创建或修改采集文件。**

帮助页为每个选项提供“作用 + 示例”，包括 `YYYY-MM`、`--test`、`--check-names`、`--check-excel`、`--update-leaders`、`--sync`、`--export`、`--verify` 和 `--help`，并额外说明以下三种容易混淆的任务：

- **同步新新闻并更新 Word**：`python main.py 2026-09`。
- **只在 E 列人工收录/排除后重建 Word**：保存并关闭 Excel，执行 `python main.py 2026-09 --export`。
- **只增量更新新闻与缓存、不改 Word**：`python main.py 2026-09 --sync`。

所有操作模式互斥；`--verify` 虽不联网，但会写入 `results/抓取检查报告.json`；`--export` 虽不刷新列表，正文缺失时仍可能联网补抓。运行帮助命令本身没有副作用。

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

### 直接修改自带的 names.txt（无需复制或重命名）

仓库已经直接包含 **`names.txt`**，克隆项目后即可读取，`names.example.txt` 已取消。需要批量修改姓名时，在项目目录运行：

~~~powershell
notepad .\names.txt
.\.venv\Scripts\python.exe main.py --check-names
~~~

从 Excel 复制一列人名后，直接粘贴到 `names.txt`，按 UTF-8 保存。
支持一行一人或“陈云、邓永翔，黄强;陈永奇”这种分隔形式，
自动跳过空行、注释、姓名表头和重复名字。

**姓名匹配精度：** 如果名单同时包含“陈云”和“陈云飞”，而文章只出现“陈云飞”，程序不会把长姓名内部的“陈云”片段另算成独立命中；如果正文其他位置又单独出现“陈云”，仍会命中两人。运行 `--check-names` 时会提示存在包含关系的姓名，方便人工检查。这个规则只能解决**名单内部**的包含冲突；若名单只有“陈云”，文章出现未录入名单的“陈云飞”时，程序仍可能命中“陈云”，请通过完整名单避免这种情况。

**人工来源优先级：names.xlsx → names.txt → config.json 原有名单。**
人工来源三者**不合并**；但会额外加入来自公开页面的 `names.auto.txt`，并去重。两个个人文件同时存在时以 names.xlsx 为准。
如果优先文件有错误，程序直接提示，不会默默使用旧名单。
**注意版本管理：** 默认的 `names.txt` 现已纳入 Git 管理；直接修改它会形成 Git 本地修改，后续 `git pull` 若有同名文件更新，可能需要先保存或合并更改。**个人长期维护的 30 多人名单推荐使用 `names.xlsx`**：它优先于 `names.txt` 读取，并继续被 `.gitignore` 忽略，不会随 `git push` 提交。曾经自行创建过 `names.txt` 的旧用户，在首次更新前请先另行备份，以免新仓库中的同名文件造成 `git pull` 冲突。详细处理方式见下文。

银河麒麟/Linux 同样支持。

更换名单后，缓存中的正文无需重新下载：

~~~powershell
.\.venv\Scripts\python.exe main.py 2026-09 --export
~~~

程序按新名单刷新 Excel 的 C/D 列和 Word，保留 E 列人工选择与 H 列事件日期。
还没有采集时正常运行 python main.py 2026-09。

### 旧用户首次更新：防止本地 names.txt 被覆盖

如果以前根据示例自己创建过 **未被 Git 跟踪** 的 `names.txt`，新版仓库中也有同名文件，直接执行 `git pull` 可能提示 `untracked working tree files would be overwritten`。在 PowerShell 项目目录，先把旧文件备份到仓库目录之外，再移走旧文件、更新仓库，最后按需要合并名单：

~~~powershell
# 仅旧用户需要；备份到项目上一级目录，避免丢失手工名单
Copy-Item .\names.txt ..\names-personal-backup.txt
# 确认备份文件存在后，再继续
Test-Path ..\names-personal-backup.txt

# 首次处理同名冲突时可先改名，而非删除旧文件
Rename-Item .\names.txt names.previous.txt
git pull origin main

# 新的 names.txt 已直接存在。参考备份内容进行人工合并即可。
notepad .\names.txt
~~~

建议把私人/长期维护的人名保存为 `names.xlsx`，不会覆盖仓库的默认 `names.txt`。如果已经存在本地修改并且 Git 报冲突，不要强制 `git reset --hard`，先保留人工名单再解决。

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

核心采集脚本仍为 `main.py`；新增 `webui.py`（本地网页界面）和 `test_webui.py`（网页安全/流程测试），原有 `test_main.py` 保留。GitHub Actions 的真实新闻采集仍由手动触发的“News collection”工作流负责。
