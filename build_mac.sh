#!/bin/bash
# 🍎 Mac (ARM/Intel) 本地打包脚本

# 开启错误即退出
set -e

echo "==========================================="
echo "🚀 开始进行 Mac 版 PyWebView 一键打包..."
echo "==========================================="

# 跳转到脚本所在的绝对路径下，防止路径对不上
cd "$(dirname "$0")"

# 1. 检查并激活虚拟环境
if [ -d "venv" ]; then
    echo "📦 检测到本地 venv 环境，正在激活..."
    source venv/bin/activate
else
    echo "⚠️ 未找到 venv 环境，正在为您初始化并安装依赖..."
    python3 -m venv venv
    source venv/bin/activate
    pip install --upgrade pip
    pip install pyinstaller python-pptx requests pywebview
fi

# 2. 检查本环境是否安装了 PyInstaller
if ! command -v pyinstaller &> /dev/null; then
    echo "📥 未检测到 pyinstaller，正在安装..."
    pip install pyinstaller
fi

# 3. 开始打包
# --onefile: 压制为单一可执行文件
# --noconsole: 关闭黑色的终端弹窗
# --add-data "tpl:tpl": 将静态资源打包进去（Mac 下路径分隔符为 :）
# 同时必须把 0-6 脚本也打包进去才能用 importlib 动态加载
echo "🛠️ 正在编译中，请稍候..."
pyinstaller --onedir --noconsole --noconfirm \
    --add-data "tpl:tpl" \
    --add-data "0-fetch_mubu.py:." \
    --add-data "1-mubu_parser.py:." \
    --add-data "2-generate_master_json.py:." \
    --add-data "3-create_master.py:." \
    --add-data "4-generate_cover.py:." \
    --add-data "5-generate_toc.py:." \
    --add-data "6-generate_body.py:." \
    --icon "app_icon.icns" \
    --hidden-import "lxml" \
    --hidden-import "html.parser" \
    --hidden-import "html" \
    --name "大纲生成PPT" gui_app.py

echo "==========================================="

echo "🔤 正在修改 Info.plist 以支持 macOS 原生中文弹窗界面..."
APP_PLIST="dist/大纲生成PPT.app/Contents/Info.plist"
if [ -f "$APP_PLIST" ]; then
    plutil -replace CFBundleLocalizations -json '["zh_CN", "zh_TW", "en"]' "$APP_PLIST"
    plutil -replace CFBundleDevelopmentRegion -string "zh_CN" "$APP_PLIST"
    echo "✅ Info.plist 已完成多语言本地化配置"

    # 修改 Info.plist 会破坏 PyInstaller 生成的默认签名，导致 Mac 提示“文件已损坏”
    echo "🔏 正在修复重签名，防止 Mac 提示文件损坏..."
    codesign --force --deep --sign - "dist/大纲生成PPT.app"
    echo "✅ 重签名完成！"
fi


echo "🧹 正在清除隔离属性，防止对方收到后提示已损坏..."
xattr -cr "dist/大纲生成PPT.app"
echo "✅ 隔离属性已清除！"

echo "📦 正在压缩为 zip..."
cd dist
rm -f "大纲生成PPT.zip"
zip -ry "大纲生成PPT.zip" "大纲生成PPT.app"
cd ..
echo "✅ 压缩完成！"

echo "✅ 打包成功！"
echo "📂 可分发的压缩包: $(pwd)/dist/大纲生成PPT.zip"
echo "==========================================="
