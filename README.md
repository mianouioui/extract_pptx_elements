# PPTX 内容提取器 / PPTX Content Extractor

[English Documentation](#english-documentation)

## 中文说明

### PPTX 内容提取器

`extract_pptx_elements` 用于从 PowerPoint `.pptx` 文件中按幻灯片提取图片、视频、音频、图表、图示、嵌入文件和文本。提取结果会按类型放入中文子文件夹，输出文件使用幻灯片序号作为前缀，便于定位资源来源。

### 功能

| 类型 | 支持内容 |
|------|----------|
| 图片 | JPG, PNG, GIF, SVG, BMP, EMF, WMF, TIFF, WebP, JFIF |
| 视频 | MP4, AVI, MOV, MKV, WebM, WMV 等 |
| 音频 | MP3, WAV, AAC, M4A, OGG, MIDI, WMA 等 |
| 图表 | 图表 XML、样式和颜色定义 |
| 图示 | SmartArt 图示 XML |
| 嵌入对象 | PDF, DOCX, XLSX, ZIP 等 |
| 幻灯片文本 | 通过 `--with-text` 导出纯文本 |
| 清单 | 自动生成 `manifest.csv`，记录输出文件与 PPTX 内部资源路径的对应关系 |

输出文件示例：`图片/001_JPG.jpg`、`视频/002_MP4.mp4`、`图表/003_CHART.xml`。前三位数字对应幻灯片编号；同一页同类型资源重复时，会自动追加 `_02`、`_03` 等序号。

### 运行入口

本项目提供三个运行入口：

- Python 源码入口：`extract_pptx_elements.py`
- macOS 单文件启动器：`extract_pptx_elements.command`
- Windows PowerShell 单文件启动器：`extract_pptx_elements.cmd`

### 运行环境

| 入口 | 运行要求 |
|------|----------|
| `extract_pptx_elements.py` | Python 3.8+ |
| `extract_pptx_elements.command` | macOS + Python 3 |
| `extract_pptx_elements.cmd` | Windows PowerShell 5.1+，以及系统自带 .NET ZIP/XML 组件 |

本项目不需要安装第三方 Python 包，也不需要虚拟环境。Windows `.cmd` 入口不依赖 Python，也不依赖 `.exe` 文件。

### macOS 使用方式

`extract_pptx_elements.command` 内嵌完整 Python 源码。运行时会将源码写入临时 Python 文件，并使用系统中的 `python3` 执行。

1. 双击 `extract_pptx_elements.command`
2. 如果系统提示“无法验证开发者”，右键点击文件，选择“打开”，再确认运行
3. 将 `.pptx` 文件拖入窗口，按回车
4. 提取结果会保存在 PPTX 文件旁边的 `pptx_extracted_elements/` 目录

也可以在终端中运行：

```bash
./extract_pptx_elements.command presentation.pptx
```

### Windows 使用方式

`extract_pptx_elements.cmd` 内嵌完整 PowerShell 提取逻辑。运行时会调用 Windows 自带的 `powershell.exe`，读取 `.cmd` 文件中的 PowerShell 段并直接执行。

```cmd
extract_pptx_elements.cmd presentation.pptx
```

Windows 入口支持与 Python 源码一致的常用选项：

```cmd
extract_pptx_elements.cmd presentation.pptx --with-text
extract_pptx_elements.cmd presentation.pptx --media-only
extract_pptx_elements.cmd presentation.pptx --overwrite
extract_pptx_elements.cmd presentation.pptx -o my_assets
```

### Python 源码使用方式

```bash
# 提取单个 PPTX 文件中的所有支持资源
python3 extract_pptx_elements.py presentation.pptx

# 指定输出目录
python3 extract_pptx_elements.py presentation.pptx -o my_assets

# 仅提取图片、视频和音频
python3 extract_pptx_elements.py presentation.pptx --media-only

# 同时提取幻灯片文本
python3 extract_pptx_elements.py presentation.pptx --with-text

# 覆盖已有输出文件
python3 extract_pptx_elements.py presentation.pptx --overwrite

# 处理当前目录下所有非临时 .pptx 文件
python3 extract_pptx_elements.py
```

默认情况下，输出目录会创建在 PPTX 文件旁边，目录名为 `pptx_extracted_elements/`。如果一次处理多个 PPTX 文件，每个文件会拥有独立子目录。

### 输出目录结构

```text
pptx_extracted_elements/
├── 图片/
│   ├── 001_JPG.jpg
│   ├── 001_JPG_02.jpg
│   └── 002_PNG.png
├── 视频/
│   └── 002_MP4.mp4
├── 图表/
│   └── 003_CHART.xml
├── 文本/
│   └── 003_TXT.txt
└── manifest.csv
```

处理多个 PPTX 文件时：

```text
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

### 清单字段

| 字段 | 说明 |
|------|------|
| `slide` | 三位幻灯片编号 |
| `output_file` | 输出文件相对路径 |
| `kind` | 资源类型，例如 `image`、`video`、`audio`、`chart`、`diagram` |
| `source_part` | PPTX 内部发起关系的 XML 路径 |
| `target_part` | PPTX 内部目标资源路径 |
| `relationship_id` | XML 关系 ID |
| `relationship_type` | 完整 XML 关系类型 URI |

### 工作原理

PowerPoint `.pptx` 文件本质上是 ZIP 压缩包，内部包含 XML、关系文件和媒体资源。本工具的流程如下：

1. 将 `.pptx` 作为 ZIP 打开
2. 读取 `ppt/presentation.xml` 和对应 relationships，确定幻灯片顺序
3. 遍历每页幻灯片的 relationship 树，识别图片、视频、音频、图表、图示和嵌入对象
4. 按类型复制资源到中文子文件夹
5. 使用幻灯片编号生成输出文件名
6. 根据 `--with-text` 选项提取可见文本
7. 生成 `manifest.csv`，记录资源来源

### 关于 `scripts/` 目录

`scripts/` 已移除。原目录只包含可选的 PyInstaller 打包脚本，不属于软件运行路径，也不是终端用户执行本工具所需内容。

当前推荐运行入口：

- Python：`extract_pptx_elements.py`
- macOS：`extract_pptx_elements.command`
- Windows：`extract_pptx_elements.cmd`

如需重新打包 `.exe` 或二进制文件，可在对应系统上直接运行 PyInstaller 命令；该流程属于发布构建流程，不影响上述入口的日常运行。

### 开源协议

MIT

---

## English Documentation

### PPTX Content Extractor

`extract_pptx_elements` extracts slide-level resources from PowerPoint `.pptx` files, including images, videos, audio files, charts, diagrams, embedded files, and optional slide text. Extracted files are grouped into Chinese type folders and named with slide-number prefixes so each asset can be traced back to its source slide.

### Features

| Type | Supported Content |
|------|-------------------|
| Images | JPG, PNG, GIF, SVG, BMP, EMF, WMF, TIFF, WebP, JFIF |
| Videos | MP4, AVI, MOV, MKV, WebM, WMV, and more |
| Audio | MP3, WAV, AAC, M4A, OGG, MIDI, WMA, and more |
| Charts | Chart XML, style definitions, and color definitions |
| Diagrams | SmartArt diagram XML |
| Embedded objects | PDF, DOCX, XLSX, ZIP, and more |
| Slide text | Plain text export through `--with-text` |
| Manifest | Automatically writes `manifest.csv` with source mapping |

Example output names include `图片/001_JPG.jpg`, `视频/002_MP4.mp4`, and `图表/003_CHART.xml`. The three-digit prefix is the slide number. Repeated resources of the same type on the same slide receive suffixes such as `_02` and `_03`.

### Runtime Entry Points

This project provides three runtime entry points:

- Python source entry point: `extract_pptx_elements.py`
- macOS single-file launcher: `extract_pptx_elements.command`
- Windows PowerShell single-file launcher: `extract_pptx_elements.cmd`

### Requirements

| Entry Point | Requirement |
|-------------|-------------|
| `extract_pptx_elements.py` | Python 3.8+ |
| `extract_pptx_elements.command` | macOS + Python 3 |
| `extract_pptx_elements.cmd` | Windows PowerShell 5.1+ and built-in .NET ZIP/XML APIs |

No third-party Python packages or virtual environment are required. The Windows `.cmd` entry point does not require Python or an `.exe` file.

### macOS Usage

`extract_pptx_elements.command` embeds the full Python source. At runtime, it writes the embedded source to a temporary Python file and executes it with the system `python3`.

1. Double-click `extract_pptx_elements.command`
2. If macOS shows an unidentified-developer warning, right-click the file, choose Open, and confirm
3. Drag a `.pptx` file into the terminal window and press Enter
4. Extracted files are written next to the PPTX file under `pptx_extracted_elements/`

Terminal usage:

```bash
./extract_pptx_elements.command presentation.pptx
```

### Windows Usage

`extract_pptx_elements.cmd` embeds the full PowerShell extraction logic. At runtime, it calls the built-in `powershell.exe`, reads the PowerShell section from the `.cmd` file, and executes it directly.

```cmd
extract_pptx_elements.cmd presentation.pptx
```

The Windows entry point supports the same common options as the Python source entry point:

```cmd
extract_pptx_elements.cmd presentation.pptx --with-text
extract_pptx_elements.cmd presentation.pptx --media-only
extract_pptx_elements.cmd presentation.pptx --overwrite
extract_pptx_elements.cmd presentation.pptx -o my_assets
```

### Python Source Usage

```bash
# Extract all supported resources from one PPTX file
python3 extract_pptx_elements.py presentation.pptx

# Extract to a custom output directory
python3 extract_pptx_elements.py presentation.pptx -o my_assets

# Extract only images, videos, and audio
python3 extract_pptx_elements.py presentation.pptx --media-only

# Also export visible slide text
python3 extract_pptx_elements.py presentation.pptx --with-text

# Overwrite existing output files
python3 extract_pptx_elements.py presentation.pptx --overwrite

# Process all non-temporary .pptx files in the current directory
python3 extract_pptx_elements.py
```

By default, the output directory is created next to the PPTX file and named `pptx_extracted_elements/`. When multiple PPTX files are processed at once, each file receives its own subdirectory.

### Output Structure

```text
pptx_extracted_elements/
├── 图片/
│   ├── 001_JPG.jpg
│   ├── 001_JPG_02.jpg
│   └── 002_PNG.png
├── 视频/
│   └── 002_MP4.mp4
├── 图表/
│   └── 003_CHART.xml
├── 文本/
│   └── 003_TXT.txt
└── manifest.csv
```

When processing multiple PPTX files:

```text
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

### Manifest CSV Fields

| Field | Description |
|-------|-------------|
| `slide` | Three-digit slide number |
| `output_file` | Relative output file path |
| `kind` | Resource type, such as `image`, `video`, `audio`, `chart`, or `diagram` |
| `source_part` | Internal PPTX XML part that owns the relationship |
| `target_part` | Internal PPTX target resource path |
| `relationship_id` | XML relationship ID |
| `relationship_type` | Full relationship type URI |

### How It Works

A PowerPoint `.pptx` file is a ZIP package containing XML files, relationship files, and media resources. The extraction process is:

1. Open the `.pptx` as a ZIP package
2. Read `ppt/presentation.xml` and its relationships to determine slide order
3. Walk each slide relationship tree to identify images, videos, audio files, charts, diagrams, and embedded objects
4. Copy resources into Chinese type folders
5. Generate output file names using slide-number prefixes
6. Extract visible text when `--with-text` is enabled
7. Write `manifest.csv` with source mapping

### About `scripts/`

The `scripts/` directory has been removed. It only contained optional PyInstaller build helpers, which are not part of the runtime path and are not required for end-user execution.

Recommended runtime entry points:

- Python: `extract_pptx_elements.py`
- macOS: `extract_pptx_elements.command`
- Windows: `extract_pptx_elements.cmd`

If maintainers need to build an `.exe` or binary package, they can run PyInstaller directly on the target operating system. That process belongs to release packaging and does not affect everyday launcher usage.

### License

MIT
