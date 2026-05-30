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

## 运行方式 / Runtime

本工具的核心逻辑位于 [extract_pptx_elements.py](extract_pptx_elements.py)。macOS 和 Windows 启动器会调用同一套 Python 逻辑，便于在未配置项目环境的终端用户机器上直接运行。项目不依赖第三方 Python 包，不需要虚拟环境或 `pip install`。
> The core extraction logic lives in [extract_pptx_elements.py](extract_pptx_elements.py). The macOS and Windows launchers invoke the same Python logic, allowing direct execution on end-user machines without project setup. No third-party Python packages, virtual environment, or `pip install` step is required.

### macOS：`extract_pptx_elements.command`

[extract_pptx_elements.command](extract_pptx_elements.command) 是 macOS 启动入口，文件内嵌完整 Python 源码。运行时会将内嵌源码写入临时 Python 文件，并使用系统中的 `python3` 执行。
> [extract_pptx_elements.command](extract_pptx_elements.command) is the macOS entry point. It embeds the full Python source, writes it to a temporary Python file at runtime, and executes it with the system `python3`.

1. 双击 `extract_pptx_elements.command`
2. 若提示「无法验证开发者」→ 右键点击 → **打开** → 确认
3. 将 `.pptx` 文件拖入窗口，按回车
4. 结果保存在 PPTX 文件旁边的 `pptx_extracted_elements/` 目录

```bash
# 或终端运行 / Or run in terminal
./extract_pptx_elements.command presentation.pptx
```

### Windows：`extract_pptx_elements.cmd`

[extract_pptx_elements.cmd](extract_pptx_elements.cmd) 是 Windows 启动入口，文件内嵌完整 Python 源码。运行时会从 `.cmd` 中提取内嵌源码到临时 Python 文件，并依次尝试使用 `py -3`、`python` 或 `python3` 执行。
> [extract_pptx_elements.cmd](extract_pptx_elements.cmd) is the Windows entry point. It extracts the embedded Python source to a temporary Python file, then attempts to run it with `py -3`, `python`, or `python3`.

```cmd
extract_pptx_elements.cmd presentation.pptx
```

Windows 环境需要已安装 Python 3 解释器；除此之外不需要配置项目环境或安装第三方依赖。如果目标机器完全没有 Python，可安装 Python 3，或使用发布包中的 `.exe`。
> Windows requires a Python 3 interpreter, but no project setup or third-party dependencies. If the target machine has no Python installed, install Python 3 or use the packaged `.exe` from a release.

---

## 环境要求 / Requirements

- Python 3.8+
- 无需第三方依赖（仅使用标准库：`zipfile`、`xml.etree`、`argparse`、`csv`）
- No third-party dependencies (stdlib only)

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

默认情况下，输出目录会创建在 PPTX 文件旁边：`pptx_extracted_elements/`。如果一次处理多个文件，每个 PPTX 会有独立子目录，避免文件混在一起。
> By default, the output folder is created next to the PPTX file: `pptx_extracted_elements/`. Multiple input files get separate subfolders.

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

## 关于 scripts 目录 / About `scripts/`

`scripts/` 已移除。原目录只包含 PyInstaller 打包脚本，不属于软件运行路径，也不是终端用户执行本工具所需内容。当前推荐运行入口是：

- macOS: `extract_pptx_elements.command`
- Windows: `extract_pptx_elements.cmd`

如果维护者需要重新打包 `.exe` 或二进制文件，可在对应系统上直接运行 PyInstaller 命令；该流程属于发布构建流程，不影响启动器的日常运行。
> `scripts/` has been removed because it only contained optional PyInstaller build helpers. Runtime entry points are the `.command` and `.cmd` launchers; binary packaging remains a separate release workflow.

---

## 开源协议 / License

MIT
