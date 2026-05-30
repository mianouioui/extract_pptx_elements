# extract_pptx_elements

从 PowerPoint `.pptx` 文件中按幻灯片提取图片、视频、音频、图表、图示、嵌入文件和文本，输出文件以幻灯片序号命名，清晰明了。
> Extract slide-level resources (images, videos, audio, charts, diagrams, embedded files, and text) from PowerPoint `.pptx` files with clean, slide-number-based file naming.

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
| **清单 / Manifest** | 自动生成 `manifest.csv`，记录每个输出文件的来源 |

输出文件命名规则：`001_JPG.jpg`、`002_MP4.mp4`、`003_CHART.xml`——前3位数字对应幻灯片编号。
> Output files are named like `001_JPG.jpg`, `002_MP4.mp4`, `003_CHART.xml` — the 3-digit prefix matches the slide number.

---

## 无需 Python 环境即可使用 / Quick Start — No Python Required

`dist/` 目录下提供了预编译的独立可执行文件：
> Pre-built standalone executables are in the [`dist/`](dist/) folder:

| 平台 / Platform | 文件 / File |
|----------|------|
| **macOS** (Apple Silicon) | [`dist/extract_pptx_elements`](dist/extract_pptx_elements) |
| **Windows** (x64) | 在 Windows 上运行 `build_windows.bat` 编译 / Build via `build_windows.bat` on Windows |

### macOS

```bash
# 下载后直接运行 / Download and run directly
./dist/extract_pptx_elements presentation.pptx

# 或安装到系统路径 / Or install system-wide
cp dist/extract_pptx_elements /usr/local/bin/
extract_pptx_elements presentation.pptx
```

### Windows

```cmd
dist\extract_pptx_elements.exe presentation.pptx
```

---

## 环境要求（Python 版）/ Requirements (Python Version)

- Python 3.8+
- 无需第三方依赖（仅使用标准库：`zipfile`、`xml.etree`、`argparse`、`csv`）
- No third-party dependencies (stdlib only)

## 安装（Python 版）/ Installation (Python)

```bash
# 克隆仓库 / Clone the repository
git clone https://github.com/yourusername/extract_pptx_elements.git
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
├── 001_JPG.jpg          # 第1页幻灯片，第1张图片 / Slide 1, first image
├── 001_JPG_02.jpg       # 第1页幻灯片，第2张图片 / Slide 1, second image
├── 002_MP4.mp4          # 第2页幻灯片，视频 / Slide 2, video
├── 002_PNG.png          # 第2页幻灯片，图片 / Slide 2, image
├── 003_CHART.xml        # 第3页幻灯片，图表 / Slide 3, chart
├── 003_TXT.txt          # 第3页幻灯片，文本（需 --with-text）
├── manifest.csv         # 提取文件总清单

# 处理多个 .pptx 时，每个文件独立子目录：
# When processing multiple .pptx files, each gets its own subfolder:
pptx_extracted_elements/
├── presentation1/
│   ├── 001_JPG.jpg
│   └── manifest.csv
└── presentation2/
    ├── 001_PNG.png
    └── manifest.csv
```

---

## Manifest CSV 字段说明 / Manifest CSV

| 列名 / Column | 说明 / Description |
|-------------------|------------------------------------------|
| slide             | 3位幻灯片编号 / 3-digit slide number |
| output_file       | 提取出的文件名 / Extracted file name |
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
4. 以幻灯片编号为前缀提取每个资源 / Extracts each resource with a slide-prefixed filename
5. （可选）从幻灯片 XML 中提取可见文本 / Optionally extracts visible text from slide XML
6. 生成 `manifest.csv` 记录完整溯源信息 / Writes a `manifest.csv` with full provenance

---

## 构建独立可执行文件 / Build Standalone Executables

可将脚本编译为无需 Python 环境的独立二进制文件：
> Build a self-contained binary that runs without Python:

### macOS

```bash
pip3 install pyinstaller
./build_macos.sh
# 输出 / Binary at: dist/extract_pptx_elements
```

### Windows

```cmd
pip install pyinstaller
build_windows.bat
REM 输出 / Binary at: dist\extract_pptx_elements.exe
```

> **注意 / Note:** PyInstaller 只能在当前操作系统下编译对应平台的二进制文件，需分别在 macOS 和 Windows 上执行。
> PyInstaller can only build for the current OS. Build on each platform separately.

---

## 开源协议 / License

MIT
