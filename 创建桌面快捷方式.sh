#!/bin/sh
# 银河麒麟：为当前用户创建应用菜单和桌面启动入口
set -eu
cd "$(dirname "$0")"
ROOT="$(pwd -P)"
APP_DIR="$HOME/.local/share/applications"
mkdir -p "$APP_DIR"
APP="$APP_DIR/chronicle-of-major-events.desktop"
cat > "$APP" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=市委大事记
Comment=新闻增量采集、Excel人工审核与Word重新导出
Exec=/bin/sh "$ROOT/启动网页界面.sh"
Path=$ROOT
Terminal=true
Icon=accessories-text-editor
Categories=Office;
EOF
chmod +x "$APP"
DESKTOP=""
if command -v xdg-user-dir >/dev/null 2>&1; then
    DESKTOP="$(xdg-user-dir DESKTOP 2>/dev/null || true)"
fi
if [ -z "$DESKTOP" ] || [ "$DESKTOP" = "$HOME" ]; then
    if [ -d "$HOME/桌面" ]; then
        DESKTOP="$HOME/桌面"
    else
        DESKTOP="$HOME/Desktop"
    fi
fi
if [ -d "$DESKTOP" ]; then
    cp "$APP" "$DESKTOP/市委大事记.desktop"
    chmod +x "$DESKTOP/市委大事记.desktop"
    echo "桌面快捷方式已生成：$DESKTOP/市委大事记.desktop"
else
    echo "未发现桌面文件夹；已加入用户应用菜单。"
fi
echo "程序入口：$APP"
echo "如果麒麟桌面提示不信任启动器，请右键该图标选择“允许启动”或对应信任选项。"
