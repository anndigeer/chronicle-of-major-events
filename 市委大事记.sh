#!/bin/sh
# 银河麒麟 / 龙芯3A4000 · 市委大事记统一操作入口
set -eu
cd "$(dirname "$0")"
ROOT=$(pwd -P)

help() {
cat <<'EOF'
市委大事记 · 银河麒麟统一入口
  sh ./市委大事记.sh                 检查环境、打开本机网页（默认）
  sh ./市委大事记.sh check           只读检查架构、Python 和依赖版本
  sh ./市委大事记.sh start           检查环境、启动网页（可附 --no-browser）
  sh ./市委大事记.sh install all     创建/使用 .venv，升级所有不合格依赖
  sh ./市委大事记.sh install openpyxl 仅安装/升级指定包，需人工确认
  sh ./市委大事记.sh desktop         创建应用菜单和桌面快捷方式
  sh ./市委大事记.sh help            显示本帮助
  允许的单包：requests、beautifulsoup4、openpyxl、python-docx、urllib3
示例：openpyxl 版本太旧（需 >=3.1,<4）：
  sh ./市委大事记.sh install openpyxl
  或在已创建的项目虚拟环境中执行：
  ./.venv/bin/python -m pip install --upgrade 'openpyxl>=3.1,<4'
检查/启动不会修改系统依赖；install 只向项目 .venv 安装，不使用 sudo。
龙芯 MIPS64 依赖可能需要本机编译，不能直接安装 ARM/x86/LoongArch 包。
EOF
}

select_python() {
  PYTHON=""
  for p in .venv/bin/python python3.12 python3.11 python3.10 python3.9 python3.8 python3; do
    if command -v "$p" >/dev/null 2>&1 &&
       "$p" -c 'import sys;assert sys.version_info>=(3,10)' >/dev/null 2>&1; then
      PYTHON="$p"; break
    fi
  done
  if [ -z "$PYTHON" ]; then
    for p in .venv/bin/python python3.9 python3.8 python3; do
      if command -v "$p" >/dev/null 2>&1 &&
         "$p" -c 'import sys;assert sys.version_info>=(3,8)' >/dev/null 2>&1; then
        PYTHON="$p"; break
      fi
    done
  fi
  if [ -z "$PYTHON" ]; then
    echo '未找到 Python 3.8+。请通过银河麒麟 MIPS64 官方软件源安装。' >&2
    exit 1
  fi
}

check() {
  select_python
  printf '使用解释器：%s\n' "$PYTHON"
  "$PYTHON" kylin_check.py
}

spec_for() {
  case "$1" in
    requests) SPEC='requests>=2.32,<3' ;;
    beautifulsoup4) SPEC='beautifulsoup4>=4.13,<5' ;;
    openpyxl) SPEC='openpyxl>=3.1,<4' ;;
    python-docx) SPEC='python-docx>=1.2,<2' ;;
    urllib3) SPEC='urllib3>=1.26,<3' ;;
    *) echo "不支持的依赖：$1" >&2; exit 2 ;;
  esac
}

install() {
  if [ "$#" -ne 1 ]; then
    echo '用法：sh ./市委大事记.sh install all 或 install 包名' >&2
    exit 2
  fi
  target="$1"
  if [ "$target" != all ]; then spec_for "$target"; fi
  select_python
  if ! "$PYTHON" -c 'import sys;assert sys.version_info>=(3,10)' >/dev/null 2>&1; then
    echo '采集依赖要求 Python 3.10+，请先安装 MIPS64 兼容的 Python 版本。' >&2
    exit 1
  fi
  if [ "$target" = all ]; then
    names=$("$PYTHON" -c 'from kylin_check import diagnose;print(" ".join(p["name"] for p in diagnose()["packages"] if not p["ok"]))')
  else
    names=$("$PYTHON" -c 'from kylin_check import diagnose;import sys;print(" ".join(p["name"] for p in diagnose()["packages"] if p["name"]==sys.argv[1] and not p["ok"]))' "$target")
  fi
  if [ -z "$names" ]; then echo '选定依赖符合版本要求，无须升级。'; exit 0; fi
  echo '准备安装/升级以下依赖（仅限项目 .venv）：'
  for name in $names; do
    spec_for "$name"
    printf '  %s\n' "$SPEC"
  done
  if [ ! -t 0 ]; then
    echo '检测到非交互输入；为避免静默安装，操作已取消。' >&2
    exit 1
  fi
  printf '是否开始联网安装/升级？[y/N] '
  read -r answer || answer=N
  case "$answer" in y|Y|yes|YES) ;; *) echo '已取消。'; exit 1 ;; esac

  if [ -e .venv ] && [ ! -x .venv/bin/python ]; then
    echo '.venv 已存在但不是可用虚拟环境，请人工备份排查（不会自动删除）。' >&2
    exit 1
  fi
  if [ -x .venv/bin/python ] &&
     ! .venv/bin/python -c 'import sys;assert sys.version_info>=(3,10)' >/dev/null 2>&1; then
    echo '现有 .venv 版本低于 3.10，请人工备份并更换（不会自动覆盖）。' >&2
    exit 1
  fi
  if [ ! -x .venv/bin/python ]; then
    "$PYTHON" -m venv --system-site-packages .venv || {
      echo '无法创建 .venv，请从银河麒麟软件源安装对应 Python 的 venv/ensurepip 组件。' >&2
      exit 1
    }
  fi
  if ! .venv/bin/python -m pip --version >/dev/null 2>&1; then
    echo '虚拟环境缺少 pip，请先安装对应 Python 的 pip/ensurepip 组件。' >&2
    exit 1
  fi
  for name in $names; do
    spec_for "$name"
    printf '执行：.venv/bin/python -m pip install --upgrade %s\n' "$SPEC"
    .venv/bin/python -m pip install --upgrade "$SPEC" || {
      echo "升级 $name 失败，MIPS64 可能需要本机编译依赖，请检查 pip 错误。" >&2
      exit 1
    }
  done
  echo '重新诊断：'
  .venv/bin/python kylin_check.py
  if ! .venv/bin/python -c 'from kylin_check import diagnose;import sys;sys.exit(0 if diagnose()["core_ready"] else 1)'; then
    echo '尚有其他未满足的依赖，请执行 check 查看对应单包更新命令。' >&2
    exit 1
  fi
  echo '依赖就绪，运行 sh ./市委大事记.sh start'
}

desktop() {
  app_dir="$HOME/.local/share/applications"
  mkdir -p "$app_dir"
  app="$app_dir/chronicle-of-major-events.desktop"
  cat > "$app" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=市委大事记
Comment=市委大事记新闻采集与Excel人工审核
Exec=/bin/sh "$ROOT/市委大事记.sh" start
Path=$ROOT
Terminal=true
Icon=accessories-text-editor
Categories=Office;
EOF
  chmod +x "$app"
  user_desktop=""
  if command -v xdg-user-dir >/dev/null 2>&1; then
    user_desktop=$(xdg-user-dir DESKTOP 2>/dev/null || true)
  fi
  if [ -z "$user_desktop" ] || [ "$user_desktop" = "$HOME" ]; then
    if [ -d "$HOME/桌面" ]; then user_desktop="$HOME/桌面"; else user_desktop="$HOME/Desktop"; fi
  fi
  if [ -d "$user_desktop" ]; then
    cp "$app" "$user_desktop/市委大事记.desktop"
    chmod +x "$user_desktop/市委大事记.desktop"
    printf '桌面快捷方式：%s\n' "$user_desktop/市委大事记.desktop"
  fi
  printf '应用菜单入口：%s\n' "$app"
  echo '如系统提示不信任桌面启动器，请选择「允许启动」。'
}

if [ "$#" -eq 0 ]; then action=start; else action="$1"; shift; fi
case "$action" in
  help|--help|-h) help ;;
  check)
    if [ "$#" -ne 0 ]; then echo 'check 不接受额外参数' >&2; exit 2; fi
    check ;;
  start)
    check
    echo '正在打开本机网页工作台……'
    exec "$PYTHON" webui.py "$@" ;;
  install) install "$@" ;;
  desktop)
    if [ "$#" -ne 0 ]; then echo 'desktop 不接受额外参数' >&2; exit 2; fi
    desktop ;;
  *) echo "未知命令：$action" >&2; help; exit 2 ;;
esac
