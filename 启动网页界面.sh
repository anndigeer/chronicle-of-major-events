#!/bin/sh
# 银河麒麟 / 龙芯 3A4000：在本机打开网页工作台
set -eu
cd "$(dirname "$0")"
PYTHON=""
for candidate in .venv/bin/python python3.12 python3.11 python3.10 python3; do
    if command -v "$candidate" >/dev/null 2>&1; then
        if "$candidate" -c 'import sys; assert sys.version_info >= (3, 10)' >/dev/null 2>&1; then
            PYTHON="$candidate"
            break
        fi
    fi
done
if [ -z "$PYTHON" ]; then
    printf '%s\n' '未找到 Python 3.10 或更新版本。' \
        '请先检查银河麒麟软件商店/系统软件源是否提供 Python 3.10+，' \
        '也可以将本程序部署在另一台服务器，本机只用浏览器访问。'
    exit 1
fi
printf '使用运行环境：%s\n' "$PYTHON"
"$PYTHON" webui.py "$@"
