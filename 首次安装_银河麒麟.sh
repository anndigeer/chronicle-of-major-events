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
if "$PYTHON" -c 'from kylin_check import diagnose; import sys; sys.exit(0 if diagnose()["core_ready"] else 1)' >/dev/null 2>&1; then
    echo '当前 Python 版本和全部采集依赖已满足要求。'
    echo '运行 sh ./启动网页界面.sh 即可打开图形界面。'
    exit 0
fi
# 另装的 Python3.10+ 可能与系统 python3 版本不同；此时 apt 的
# python3-* 模块通常安装到系统 Python，不保证可用于当前解释器。
SYSTEM_VERSION=""
if command -v python3 >/dev/null 2>&1; then
    SYSTEM_VERSION=$(python3 -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || true)
fi
SELECTED_VERSION=$("$PYTHON" -c 'import sys; print("%d.%d" % sys.version_info[:2])')
if [ -n "$SYSTEM_VERSION" ] && [ "$SELECTED_VERSION" != "$SYSTEM_VERSION" ]; then
    echo "系统 python3 为 $SYSTEM_VERSION，当前采集解释器为 $SELECTED_VERSION。"
    echo '为防止把依赖安装进错误的 Python，暂不自动执行 apt；请用当前解释器的 venv/pip 或咨询系统管理员。'
elif command -v apt-get >/dev/null 2>&1; then
    printf '准备使用系统软件源安装原生 MIPS64 Python 依赖，可能需要管理员密码。继续？[y/N] '
    read -r reply
    case "$reply" in y|Y|yes|YES) ;; *) echo '已取消安装。'; exit 1;; esac
    sudo apt-get update
    if sudo apt-get install -y python3-requests python3-bs4 python3-openpyxl python3-docx; then
        if "$PYTHON" -c 'from kylin_check import diagnose; import sys; sys.exit(0 if diagnose()["core_ready"] else 1)' >/dev/null 2>&1; then
            echo '安装完成。运行 sh ./启动网页界面.sh 即可打开图形界面。'
            exit 0
        fi
    fi
fi
echo '系统软件源未能提供满足版本要求的全部依赖，或软件包与当前 Python 不兼容。'
"$PYTHON" kylin_check.py || true
echo '请在本机安装对应架构的软件包；也可在隔离虚拟环境中尝试：'
echo '使用兼容的 Python 3.10+ 创建虚拟环境，例如：python3.10 -m venv --system-site-packages .venv'
echo '.venv/bin/python -m pip install -r requirements.txt'
echo '注意：龙芯 MIPS64 的第三方依赖可能需要本机编译，无法保证一次安装成功。'
exit 1
