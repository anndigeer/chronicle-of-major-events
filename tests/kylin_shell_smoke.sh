#!/bin/sh
# Simulate an ordinary Kylin user launching the project from a path with spaces.
# No sudo, no package installation and no actual network news requests.
set -eu
cd "$(dirname "$0")/.."
for file in "启动网页界面.sh" "首次安装_银河麒麟.sh" "创建桌面快捷方式.sh"; do
    sh -n "./$file"
    printf '语法正常：%s\n' "$file"
done

tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT HUP INT TERM
project="$tmp/银河麒麟 项目"
mkdir -p "$project" "$tmp/home/桌面"
cp "./启动网页界面.sh" "$project/启动网页界面.sh"
cp "./创建桌面快捷方式.sh" "$project/创建桌面快捷方式.sh"
cat > "$project/webui.py" <<'PY'
import sys
print("GUI_LAUNCH_SMOKE " + " ".join(sys.argv[1:]))
PY

launch_output=$(sh "$project/启动网页界面.sh" --no-browser --port 0)
printf '%s\n' "$launch_output"
case "$launch_output" in
    *"GUI_LAUNCH_SMOKE --no-browser --port 0"*) ;;
    *) echo "ERROR: launcher did not forward GUI arguments" >&2; exit 1 ;;
esac

HOME="$tmp/home" sh "$project/创建桌面快捷方式.sh"
shortcut="$tmp/home/.local/share/applications/chronicle-of-major-events.desktop"
test -f "$shortcut" || { echo "ERROR: missing desktop menu launcher" >&2; exit 1; }
grep -q '^Terminal=true$' "$shortcut"
grep -q "^Path=$project$" "$shortcut"
grep -q '^Name=市委大事记$' "$shortcut"
printf '%s\n' "Kylin launcher and shortcut smoke: PASS"
