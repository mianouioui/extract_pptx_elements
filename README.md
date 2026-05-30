# extract_pptx_elements

## PPTX 内容提取器

从 PowerPoint `.pptx` 文件中按幻灯片提取图片、视频、音频、图表、图示、嵌入文件和文本，按类型放入中文子文件夹，输出文件以幻灯片序号命名，清晰明了。
> Extract slide-level resources (images, videos, audio, charts, diagrams, embedded files, and text) from PowerPoint `.pptx` files into type-based folders with clean, slide-number-based file naming.

## 功能 / Features

| 类型 | 支持格式 |
|------|---------|
| **图片 / Images** | JPG, PNG, GIF, SVG, BMP, EMF, WMF, TIFF, WebP, JFIF |
| **视频 / Videos** | MP4, AVI, MOV, MKV, WebM, WMV 等 |
| **音频 / Audio** | MP3, WAV, AAC, M4A, OGG, MIDI, WMA 等 |
| **图表 / Charts** | 图表 XML + 样式和颜色定义 |
| **图示 / Diagrams** | SmartArt 图示 XML |
| **嵌入对象 / Embedded** | PDF, DOCX, XLSX, ZIP 等 |
| **幻灯片文本 / Text** | 通过 `--with-text` 提取纯文本 |
| **中文分类目录 / Chinese Folders** | 图片、视频、音频、图表、图示、嵌入文件、文本 |
| **清单 / Manifest** | 自动生成 `manifest.csv`，记录每个输出文件的来源 |

输出文件命名规则：`图片/001_JPG.jpg`、`视频/002_MP4.mp4`、`图表/003_CHART.xml`——前3位数字对应幻灯片编号。
> Output files are named like `图片/001_JPG.jpg`, `视频/002_MP4.mp4`, `图表/003_CHART.xml` — the 3-digit prefix matches the slide number.

---

## 🚀 快速开始 / Quick Start

### macOS

[extract_pptx_elements.command](extract_pptx_elements.command) 是一个独立文件，内嵌完整 Python 源码，双击即可运行。**兼容 Intel 和 Apple Silicon Mac**，只需系统自带 Python 3。
> [extract_pptx_elements.command](extract_pptx_elements.command) is a single self-contained file. Double-click to run. **Works on Intel and Apple Silicon Macs** with system Python 3.

1. 双击 `extract_pptx_elements.command`
2. 若提示「无法验证开发者」→ 右键点击 → **打开** → 确认
3. 将 `.pptx` 文件拖入窗口，按回车
4. 结果保存在 `pptx_extracted_elements/` 目录

```bash
# 或终端运行 / Or run in terminal
./extract_pptx_elements.command presentation.pptx
```

### Windows

需将 `extract_pptx_elements.cmd` 与 `extract_pptx_elements.py` 放在同一目录，双击 `.cmd` 运行。或自行编译 `.exe`。
> Place `extract_pptx_elements.cmd` and `extract_pptx_elements.py` in the same folder, double-click `.cmd`. Or build `.exe`.

```cmd
extract_pptx_elements.cmd presentation.pptx
```

---

## 预编译二进制（可选）/ Standalone Binaries (Optional)

不需要，上面的 `.command` 文件已经够用了。以下仅作备选：
> The `.command` file above is all you need. These are optional alternatives:

| 平台 / Platform | 文件 / File |
|----------|------|
| **macOS** (Apple Silicon / M1-M3) | [`dist/extract_pptx_elements`](dist/extract_pptx_elements) |
| **macOS** (Intel) | 请双击 `.command` 启动器（自动用 Python 源码）或自行编译 |
| **Windows** (x64) | 在 Windows 上运行 `scripts\build_windows.bat` 编译 / Build via `scripts\build_windows.bat` on Windows |

### macOS

```bash
# 终端运行 / Terminal
./extract_pptx_elements.command presentation.pptx

# 或直接用二进制 / Or binary directly (Apple Silicon only)
./dist/extract_pptx_elements presentation.pptx
```

### Windows

```cmd
extract_pptx_elements.cmd presentation.pptx
```

---

## 环境要求（Python 版）/ Requirements (Python Version)

- Python 3.8+
- 无需第三方依赖（仅使用标准库：`zipfile`、`xml.etree`、`argparse`、`csv`）
- No third-party dependencies (stdlib only)

## 安装（Python 版）/ Installation (Python)

```bash
# 克隆仓库 / Clone the repository
git clone https://github.com/mianouioui/extract_pptx_elements.git
cd extract_pptx_elements

# 赋予执行权限（可选）/ Make the script executable (optional)
chmod +x extract_pptx_elements.py
```

---

## 使用方法 / Usage

```bash
# 提取单个 .pptx 文件中所有元素 / Extract everything from a single .pptx file
python3 extract_pptx_elements.py presentation.pptx

# 指定输出目录 / Extract to a custom output directory
python3 extract_pptx_elements.py presentation.pptx -o my_assets/

# 仅提取图片、视频、音频 / Only images, videos, and audio
python3 extract_pptx_elements.py presentation.pptx --media-only

# 同时提取幻灯片文本 / Also extract slide text as 001_TXT.txt, etc.
python3 extract_pptx_elements.py presentation.pptx --with-text

# 覆盖已有文件 / Overwrite existing files
python3 extract_pptx_elements.py presentation.pptx --overwrite

# 处理当前目录下所有 .pptx 文件 / Process all .pptx files in current directory
python3 extract_pptx_elements.py
```

---

## 输出目录结构 / Output Structure

```
pptx_extracted_elements/
├── 图片/
│   ├── 001_JPG.jpg      # 第1页幻灯片，第1张图片 / Slide 1, first image
│   ├── 001_JPG_02.jpg   # 第1页幻灯片，第2张图片 / Slide 1, second image
│   └── 002_PNG.png      # 第2页幻灯片，图片 / Slide 2, image
├── 视频/
│   └── 002_MP4.mp4      # 第2页幻灯片，视频 / Slide 2, video
├── 图表/
│   └── 003_CHART.xml    # 第3页幻灯片，图表 / Slide 3, chart
├── 文本/
│   └── 003_TXT.txt      # 第3页幻灯片，文本（需 --with-text）
└── manifest.csv         # 提取文件总清单

# 处理多个 .pptx 时，每个文件独立子目录：
# When processing multiple .pptx files, each gets its own subfolder:
pptx_extracted_elements/
├── presentation1/
│   ├── 图片/
│   │   └── 001_JPG.jpg
│   └── manifest.csv
└── presentation2/
    ├── 图片/
    │   └── 001_PNG.png
    └── manifest.csv
```

---

## Manifest CSV 字段说明 / Manifest CSV

| 列名 / Column | 说明 / Description |
|-------------------|------------------------------------------|
| slide             | 3位幻灯片编号 / 3-digit slide number |
| output_file       | 提取文件的相对路径，如 `图片/001_JPG.jpg` / Extracted relative path |
| kind              | 类型：image, video, audio, chart, diagram 等 |
| source_part       | .pptx 内部源路径 / Source path in zip |
| target_part       | .pptx 内部资源路径 / Target resource path |
| relationship_id   | XML 关系 ID |
| relationship_type | 完整的关系类型 URI |

---

## 工作原理 / How It Works

PowerPoint `.pptx` 文件本质上是一个包含 XML 和媒体文件的 ZIP 压缩包。本工具的工作流程：
> PowerPoint `.pptx` files are ZIP archives containing XML and media files. This tool:

1. 将 `.pptx` 作为 ZIP 打开 / Opens the `.pptx` as a ZIP archive
2. 读取 `ppt/presentation.xml` 确定幻灯片顺序 / Reads `ppt/presentation.xml` to determine slide order
3. 遍历每页幻灯片的关系树，查找图片、视频、音频、图表、图示和嵌入对象 / Walks each slide's relationship tree to find resources
4. 按类型放入中文子文件夹，并以幻灯片编号为前缀提取每个资源 / Extracts each resource into a type folder with a slide-prefixed filename
5. （可选）从幻灯片 XML 中提取可见文本 / Optionally extracts visible text from slide XML
6. 生成 `manifest.csv` 记录完整溯源信息 / Writes a `manifest.csv` with full provenance

---

## 构建独立可执行文件 / Build Standalone Executables

可将脚本编译为无需 Python 环境的独立二进制文件。启动器（`.command`/`.cmd`）已能覆盖绝大多数场景，以下为进阶用法。
> Build a self-contained binary that runs without Python. The `.command`/`.cmd` launchers cover most use cases already.

### macOS

```bash
pip3 install pyinstaller

# Apple Silicon (M1/M2/M3)
./scripts/build_macos.sh

# Intel Mac（在 Intel Mac 上运行）
./scripts/build_macos.sh

# 或在 Apple Silicon 上编译通用二进制（同时支持 Intel + ARM）
pyinstaller --onefile --name extract_pptx_elements --specpath build/specs --target-arch universal2 extract_pptx_elements.py
```

### Windows

```cmd
pip install pyinstaller
scripts\build_windows.bat
```

> **注意 / Note:** PyInstaller 只能在当前操作系统下编译，跨平台编译需分别执行。Intel Mac 用户直接用 `.command` 启动器即可，无需编译。

---

## 开源协议 / License

MIT
