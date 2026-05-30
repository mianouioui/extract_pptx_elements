#!/bin/bash
# ============================================================
#  PPTX 元素提取工具 - macOS 启动器
#  直接双击即可运行（无需手动打开终端）
# ============================================================

cd "$(dirname "$0")" || exit 1

# 如果没有传参数，显示交互式提示
if [ $# -eq 0 ]; then
    echo "╔══════════════════════════════════════════╗"
    echo "║     PPTX 元素提取工具 V1.0.0            ║"
    echo "║     extract_pptx_elements               ║"
    echo "╚══════════════════════════════════════════╝"
    echo ""
    echo "用法: 将 .pptx 文件拖拽到此窗口，然后按回车"
    echo "      或直接输入 .pptx 文件路径"
    echo ""
    echo "常用选项:"
    echo "  --with-text   同时提取幻灯片文本"
    echo "  --media-only  仅提取图片/视频/音频"
    echo "  --overwrite   覆盖已有文件"
    echo "  -o 目录名     指定输出目录"
    echo ""
    echo "──────────────────────────────────────────"
    read -r -p "请输入 .pptx 文件路径（可拖拽）: " input_args
    # 将输入解析为参数（去掉拖拽产生的首尾空格和引号）
    eval "set -- $input_args"
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PY_SOURCE="$SCRIPT_DIR/extract_pptx_elements.py"
BINARY="$SCRIPT_DIR/dist/extract_pptx_elements"

# 策略1：优先使用 Python 源码（兼容所有 Mac，包括 Intel）
if command -v python3 &>/dev/null; then
    echo "→ 使用 Python 源码运行 ..."
    python3 "$PY_SOURCE" "$@"
    EXIT_CODE=$?

# 策略2：使用预编译二进制（仅 Apple Silicon）
elif [ -x "$BINARY" ]; then
    ARCH=$(uname -m)
    if [ "$ARCH" = "arm64" ]; then
        echo "→ 使用预编译二进制运行 (ARM64) ..."
        "$BINARY" "$@"
        EXIT_CODE=$?
    else
        echo "✗ 预编译二进制仅支持 Apple Silicon (M1/M2/M3)，"
        echo "  您的 Mac 是 Intel ($ARCH) 芯片。"
        echo "  请安装 Python 3 后重试:"
        echo "    https://www.python.org/downloads/"
        echo "  或自行编译:"
        echo "    pip3 install pyinstaller && ./build_macos.sh"
        exit 1
    fi

# 策略3：都没找到
else
    echo "✗ 未找到 Python 3 或可执行文件。"
    echo "  请安装 Python 3: https://www.python.org/downloads/"
    echo "  安装后双击此文件即可运行。"
    exit 1
fi

echo ""
echo "完成！(退出码: $EXIT_CODE)"
echo "按任意键关闭此窗口 ..."
read -r -n 1 -s
exit $EXIT_CODE
