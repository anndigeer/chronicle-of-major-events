#!/bin/sh
# Kylin unified Shell launcher: no sudo, downloads, or real news requests.
set -eu
cd "$(dirname "$0")/.."
sh -n "./Start.sh"

tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT HUP INT TERM
project="$tmp/银河麒麟 项目"
mkdir -p "$project" "$tmp/home/桌面"
cp ./Start.sh "$project/Start.sh"
cat > "$project/kylin_check.py" <<'PY'
def diagnose():
    return {'core_ready': False, 'packages': [
        {'name':'openpyxl','ok':False},
        {'name':'requests','ok':True}]}
if __name__ == '__main__':
    print('KYLINDIAG_MOCK')
PY
cat > "$project/webui.py" <<'PY'
import sys
print('GUI_LAUNCH_SMOKE ' + ' '.join(sys.argv[1:]))
PY

# Default and explicit start must both forward args, even from unicode/spaces.
launch=$(sh "$project/Start.sh" start --no-browser --port 0)
case "$launch" in
  *"KYLINDIAG_MOCK"*"GUI_LAUNCH_SMOKE --no-browser --port 0"*) ;;
  *) echo "Launcher/start failed: $launch" >&2; exit 1 ;;
esac
default=$(sh "$project/Start.sh" --no-browser 2>&1 || true)
case "$default" in
  *"未知命令"*) ;; # Default only applies if NO args: documented contract.
  *) echo "Unknown options should fail clearly" >&2; exit 1 ;;
esac
default=$(sh "$project/Start.sh")
case "$default" in
  *"GUI_LAUNCH_SMOKE"*) ;;
  *) echo "Default startup failed" >&2; exit 1 ;;
esac
check=$(sh "$project/Start.sh" check)
case "$check" in
  *"KYLINDIAG_MOCK"*) ;;
  *) echo "Check failed" >&2; exit 1 ;;
esac
help=$(sh "$project/Start.sh" help)
case "$help" in
  *"install openpyxl"*"openpyxl>=3.1,<4"*) ;;
  *) echo "Help missing targeted installation guidance" >&2; exit 1 ;;
esac

# Invalid package and noninteractive install must not create a venv.
if sh "$project/Start.sh" install arbitrary-package >/dev/null 2>&1; then
  echo "Invalid package unexpectedly accepted" >&2; exit 1
fi
if sh "$project/Start.sh" install openpyxl </dev/null >/dev/null 2>&1; then
  echo "Noninteractive dependency installation unexpectedly succeeded" >&2; exit 1
fi
test ! -e "$project/.venv" || {
  echo "Noninteractive install modified the environment" >&2; exit 1;
}

HOME="$tmp/home" sh "$project/Start.sh" desktop
shortcut="$tmp/home/.local/share/applications/chronicle-of-major-events.desktop"
test -f "$shortcut" || { echo "Missing application shortcut" >&2; exit 1; }
grep -q '^Terminal=true$' "$shortcut"
grep -q "^Path=$project$" "$shortcut"
grep -q "^Exec=/bin/sh \"$project/Start.sh\" start$" "$shortcut"
grep -q '^Name=市委大事记$' "$shortcut"
echo 'PASS: unified Kylin shell start/check/install guard/desktop smoke'
