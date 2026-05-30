@echo off
set VERSION=V1.0.2
chcp 65001 >nul
REM ============================================================
REM  PPTX 元素提取工具 %VERSION% - Windows 启动器
REM  直接双击即可运行
REM ============================================================

cd /d "%~dp0"

if "%~1"=="" (
    echo ╔══════════════════════════════════════════╗
    echo ║     PPTX 元素提取工具 %VERSION%            ║
    echo ║     extract_pptx_elements               ║
    echo ╚══════════════════════════════════════════╝
    echo.
    echo 用法: 将 .pptx 文件拖拽到此窗口，然后按回车
    echo       或直接输入 .pptx 文件路径
    echo.
    echo 常用选项:
    echo   --with-text   同时提取幻灯片文本
    echo   --media-only  仅提取图片/视频/音频
    echo   --overwrite   覆盖已有文件
    echo   -o 目录名     指定输出目录
    echo.
    echo ──────────────────────────────────────────
    set /p input_args="请输入 .pptx 文件路径（可拖拽）: "
    REM 简单处理：去掉开头结尾的双引号
    set input_args=%input_args:"=%
    %0 %input_args%
    goto :end
)

set PY_SOURCE=%~dp0extract_pptx_elements.py
set BINARY=%~dp0dist\extract_pptx_elements.exe

REM 策略1：优先使用 Python 源码
python3 --version >/dev/null 2>&1
if %errorlevel% equ 0 (
    echo → 使用 Python 源码运行 ...
    python3 "%PY_SOURCE%" %*
    goto :end
)

REM 也尝试 python（不带3）
python --version >/dev/null 2>&1
if %errorlevel% equ 0 (
    echo → 使用 Python 源码运行 ...
    python "%PY_SOURCE%" %*
    goto :end
)

REM 策略2：使用预编译二进制
if exist "%BINARY%" (
    echo → 使用预编译二进制运行 ...
    "%BINARY%" %*
    goto :end
)

REM 策略3：都没找到
echo ✗ 未找到 Python 或可执行文件。
echo   请安装 Python 3: https://www.python.org/downloads/
echo   或在 Windows 上编译: build_windows.bat
echo.
echo 按任意键关闭此窗口 ...
pause >/dev/null
exit /b 1

:end
echo.
echo 完成！(退出码: %errorlevel%)
echo 按任意键关闭此窗口 ...
pause >/dev/null
exit /b %errorlevel%
