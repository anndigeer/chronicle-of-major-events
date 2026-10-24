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
    # 银河麒麟旧版本可能自带 Python 3.8/3.9：允许打开标准库
    # 网页工作台查看诊断，但核心采集仍要求 Python 3.10+。
    for candidate in .venv/bin/python python3.9 python3.8 python3; do
        if command -v "$candidate" >/dev/null 2>&1 && \
           "$candidate" -c 'import sys; assert sys.version_info >= (3, 8)' >/dev/null 2>&1; then
            PYTHON="$candidate"
            printf '%s\n' '提示：Python 3.8/3.9 只提供图形界面和环境诊断，不能运行采集核心。'
            break
        fi
    done
fi
if [ -z "$PYTHON" ]; then
    printf '%s\n' '未找到 Python 3.8 或更新版本；网页工作台暂不能在本机启动。' \
        '请在银河麒麟软件商店中安装对应 MIPS64 架构的 Python。'
    exit 1
fi
printf '当前网页工作台使用：%s\n' "$PYTHON"
"$PYTHON" webui.py "$@"
