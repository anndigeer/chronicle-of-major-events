#!/bin/sh
# 银河麒麟龙芯机器依赖安装指导。优先使用系统原生软件包。
set -eu
cd "$(dirname "$0")"
PYTHON=""
for candidate in .venv/bin/python python3.12 python3.11 python3.10 python3; do
    if command -v "$candidate" >/dev/null 2>&1 && \
       "$candidate" -c 'import sys; assert sys.version_info >= (3, 10)' >/dev/null 2>&1; then
        PYTHON="$candidate"
        break
    fi
done
if [ -z "$PYTHON" ]; then
    echo '未找到兼容的 Python 3.10 或更高版本。'
    echo '请先通过银河麒麟 MIPS64 软件源安装，不要安装 ARM/LoongArch/x86 程序包。'
    exit 1
fi
echo "检测到符合要求的 Python：$PYTHON"
if "$PYTHON" -c 'import requests, bs4, openpyxl, docx' >/dev/null 2>&1; then
    echo '系统已具备采集依赖，可执行：sh ./启动网页界面.sh'
    exit 0
fi
if command -v apt-get >/dev/null 2>&1; then
    printf '准备使用系统软件源安装原生 MIPS64 Python 依赖，可能需要管理员密码。继续？[y/N] '
    read -r reply
    case "$reply" in y|Y|yes|YES) ;; *) echo '已取消安装。'; exit 1;; esac
    sudo apt-get update
    if sudo apt-get install -y python3-requests python3-bs4 python3-openpyxl python3-docx; then
        if "$PYTHON" -c 'import requests, bs4, openpyxl, docx' >/dev/null 2>&1; then
            echo '安装完成。运行 sh ./启动网页界面.sh 即可打开图形界面。'
            exit 0
        fi
    fi
fi
echo '系统软件源未能提供全部依赖，或 Python 与原生软件包版本不匹配。'
echo '请在本机安装对应架构的软件包；也可在隔离虚拟环境中尝试：'
echo '使用兼容的 Python 3.10+ 创建虚拟环境，例如：python3.10 -m venv --system-site-packages .venv'
echo '.venv/bin/python -m pip install -r requirements.txt'
echo '注意：龙芯 MIPS64 的第三方依赖可能需要本机编译，无法保证一次安装成功。'
exit 1
