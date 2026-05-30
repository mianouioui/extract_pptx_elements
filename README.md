# PPTX 内容提取器 / PPTX Content Extractor

[English Documentation](#english-documentation) | [Documentation française](#documentation-française)

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
| 幻灯片文本 | 默认导出纯文本，可通过 `--no-text` 关闭 |
| 清单 | 自动生成 `manifest.csv`，记录输出文件与 PPTX 内部资源路径的对应关系 |

输出文件示例：`图片/001_JPG.jpg`、`视频/002_MP4.mp4`、`图表/003_CHART.xml`。前三位数字对应幻灯片编号；同一页同类型资源重复时，会自动追加 `_02`、`_03` 等序号。

### 运行入口

本项目提供三个运行入口：

- Python 源码入口：`extract_pptx_elements.py`
- macOS 单文件启动器：`extract_pptx_elements.command`
- Windows 启动器：`extract_pptx_elements.cmd`

### 运行环境

| 入口 | 运行要求 |
|------|----------|
| `extract_pptx_elements.py` | Python 3.8+ |
| `extract_pptx_elements.command` | macOS + Python 3 |
| `extract_pptx_elements.cmd` | Windows；优先使用 Python 3，回退到 PowerShell 5.1+ 和系统自带 .NET ZIP/XML 组件 |

本项目不需要安装第三方 Python 包，也不需要虚拟环境。Windows `.cmd` 已内嵌 Python 主实现和 PowerShell 回退实现：如果系统有 Python 3，会优先运行内嵌 Python；如果没有 Python，则使用内嵌 PowerShell。因此 Windows 入口不依赖外部 `.py` 或 `.exe` 文件。单独的 `extract_pptx_elements.py` 保留为可读、可修改的参考源码。

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

`extract_pptx_elements.cmd` 是 Windows 单文件启动入口。运行时会依次尝试 `py -3`、`python` 和 `python3`；如果找到 Python，会从 `.cmd` 自身提取内嵌 Python 源码到临时 `.py` 文件并执行。如果 Python 不可用，则回退到内嵌 PowerShell 提取逻辑，将 PowerShell 段写入临时 `.ps1` 文件后执行。

双击 `.cmd` 且没有传入参数时，窗口会提示拖入或输入 `.pptx` 文件路径；也可以将 PPTX 文件直接拖到 `.cmd` 文件上运行。

```cmd
extract_pptx_elements.cmd presentation.pptx
```

Windows 入口支持与 Python 源码一致的常用选项：

```cmd
extract_pptx_elements.cmd presentation.pptx --no-text
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

# 不提取幻灯片文本
python3 extract_pptx_elements.py presentation.pptx --no-text

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
6. 默认提取可见文本；使用 `--no-text` 可关闭文本导出
7. 生成 `manifest.csv`，记录资源来源

### 关于 `scripts/` 目录

`scripts/` 已移除。原目录只包含可选的 PyInstaller 打包脚本，不属于软件运行路径，也不是终端用户执行本工具所需内容。

当前推荐运行入口：

- Python：`extract_pptx_elements.py`
- macOS：`extract_pptx_elements.command`
- Windows：`extract_pptx_elements.cmd`

如需重新打包 `.exe` 或二进制文件，可在对应系统上直接运行 PyInstaller 命令；该流程属于发布构建流程，不影响上述入口的日常运行。

### 发布记录

详见 [docs/](docs/) 目录下各版本发布记录。

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
| Slide text | Plain text export by default; disable it with `--no-text` |
| Manifest | Automatically writes `manifest.csv` with source mapping |

Example output names include `图片/001_JPG.jpg`, `视频/002_MP4.mp4`, and `图表/003_CHART.xml`. The three-digit prefix is the slide number. Repeated resources of the same type on the same slide receive suffixes such as `_02` and `_03`.

### Runtime Entry Points

This project provides three runtime entry points:

- Python source entry point: `extract_pptx_elements.py`
- macOS single-file launcher: `extract_pptx_elements.command`
- Windows launcher: `extract_pptx_elements.cmd`

### Requirements

| Entry Point | Requirement |
|-------------|-------------|
| `extract_pptx_elements.py` | Python 3.8+ |
| `extract_pptx_elements.command` | macOS + Python 3 |
| `extract_pptx_elements.cmd` | Windows; prefers Python 3, then falls back to PowerShell 5.1+ and built-in .NET ZIP/XML APIs |

No third-party Python packages or virtual environment are required. The Windows `.cmd` embeds both the primary Python implementation and a PowerShell fallback. If Python 3 is available, it runs the embedded Python source first; if Python is unavailable, it falls back to the embedded PowerShell implementation. The Windows entry point does not depend on an external `.py` or `.exe` file. The standalone `extract_pptx_elements.py` remains available as readable and editable reference source.

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

`extract_pptx_elements.cmd` is the Windows single-file entry point. At runtime, it tries `py -3`, `python`, and `python3`; if Python is available, it extracts the embedded Python source from the `.cmd` file into a temporary `.py` file and executes it. If Python is unavailable, it falls back to the embedded PowerShell extraction logic, writes the PowerShell section to a temporary `.ps1` file, and executes that file.

When the `.cmd` file is double-clicked without arguments, the window prompts for a `.pptx` path. A PPTX file can also be dragged directly onto the `.cmd` file.

```cmd
extract_pptx_elements.cmd presentation.pptx
```

The Windows entry point supports the same common options as the Python source entry point:

```cmd
extract_pptx_elements.cmd presentation.pptx --no-text
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

# Disable visible slide text export
python3 extract_pptx_elements.py presentation.pptx --no-text

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
6. Extract visible text by default; disable text export with `--no-text`
7. Write `manifest.csv` with source mapping

### About `scripts/`

The `scripts/` directory has been removed. It only contained optional PyInstaller build helpers, which are not part of the runtime path and are not required for end-user execution.

Recommended runtime entry points:

- Python: `extract_pptx_elements.py`
- macOS: `extract_pptx_elements.command`
- Windows: `extract_pptx_elements.cmd`

If maintainers need to build an `.exe` or binary package, they can run PyInstaller directly on the target operating system. That process belongs to release packaging and does not affect everyday launcher usage.

### Release Notes

See per-version release notes in the [docs/](docs/) directory.

### License

MIT

---

## Documentation française

### Extracteur de contenu PPTX

`extract_pptx_elements` extrait les ressources d’un fichier PowerPoint `.pptx` au niveau de chaque diapositive, notamment les images, les vidéos, les fichiers audio, les graphiques, les diagrammes, les fichiers intégrés et, en option, le texte visible des diapositives. Les fichiers extraits sont classés dans des dossiers de type en chinois et nommés avec un préfixe correspondant au numéro de la diapositive, afin de faciliter le suivi de leur origine.

### Fonctionnalités

| Type | Contenu pris en charge |
|------|------------------------|
| Images | JPG, PNG, GIF, SVG, BMP, EMF, WMF, TIFF, WebP, JFIF |
| Vidéos | MP4, AVI, MOV, MKV, WebM, WMV, etc. |
| Audio | MP3, WAV, AAC, M4A, OGG, MIDI, WMA, etc. |
| Graphiques | XML du graphique, définitions de style et de couleurs |
| Diagrammes | XML des diagrammes SmartArt |
| Objets intégrés | PDF, DOCX, XLSX, ZIP, etc. |
| Texte des diapositives | Export en texte brut par défaut ; désactivation avec `--no-text` |
| Manifeste | Génération automatique de `manifest.csv` avec la correspondance des sources |

Exemples de noms de sortie : `图片/001_JPG.jpg`, `视频/002_MP4.mp4` et `图表/003_CHART.xml`. Le préfixe à trois chiffres correspond au numéro de la diapositive. Lorsque plusieurs ressources du même type existent sur une même diapositive, des suffixes comme `_02` et `_03` sont ajoutés automatiquement.

### Points d’entrée

Le projet fournit trois points d’entrée :

- Script Python source : `extract_pptx_elements.py`
- Lanceur macOS en fichier unique : `extract_pptx_elements.command`
- Lanceur Windows : `extract_pptx_elements.cmd`

### Prérequis

| Point d’entrée | Prérequis |
|----------------|-----------|
| `extract_pptx_elements.py` | Python 3.8+ |
| `extract_pptx_elements.command` | macOS + Python 3 |
| `extract_pptx_elements.cmd` | Windows ; utilise Python 3 en priorité, puis bascule vers PowerShell 5.1+ et les API .NET ZIP/XML intégrées |

Aucun paquet Python tiers ni environnement virtuel n’est nécessaire. Le point d’entrée Windows `.cmd` intègre à la fois l’implémentation Python principale et une implémentation PowerShell de secours. Si Python 3 est disponible, il exécute d’abord le code Python intégré ; si Python n’est pas disponible, il bascule vers l’implémentation PowerShell intégrée. Le point d’entrée Windows ne dépend d’aucun fichier `.py` ou `.exe` externe. Le fichier séparé `extract_pptx_elements.py` reste disponible comme source lisible et modifiable.

### Utilisation sur macOS

`extract_pptx_elements.command` contient le code source Python complet. À l’exécution, il écrit ce code intégré dans un fichier Python temporaire, puis l’exécute avec le `python3` du système.

1. Double-cliquez sur `extract_pptx_elements.command`
2. Si macOS affiche un avertissement concernant un développeur non identifié, faites un clic droit sur le fichier, choisissez Ouvrir, puis confirmez
3. Glissez un fichier `.pptx` dans la fenêtre du terminal, puis appuyez sur Entrée
4. Les fichiers extraits sont écrits à côté du fichier PPTX, dans `pptx_extracted_elements/`

Utilisation dans le terminal :

```bash
./extract_pptx_elements.command presentation.pptx
```

### Utilisation sur Windows

`extract_pptx_elements.cmd` est le point d’entrée Windows en fichier unique. À l’exécution, il essaie `py -3`, `python` puis `python3` ; si Python est disponible, il extrait le code Python intégré depuis le fichier `.cmd` vers un fichier `.py` temporaire, puis l’exécute. Si Python n’est pas disponible, il bascule vers la logique PowerShell intégrée, écrit la section PowerShell dans un fichier `.ps1` temporaire, puis exécute ce fichier.

Lorsque le fichier `.cmd` est ouvert sans argument, la fenêtre demande le chemin d’un fichier `.pptx`. Il est également possible de faire glisser un fichier PPTX directement sur le fichier `.cmd`.

```cmd
extract_pptx_elements.cmd presentation.pptx
```

Le point d’entrée Windows prend en charge les mêmes options courantes que le script Python :

```cmd
extract_pptx_elements.cmd presentation.pptx --no-text
extract_pptx_elements.cmd presentation.pptx --media-only
extract_pptx_elements.cmd presentation.pptx --overwrite
extract_pptx_elements.cmd presentation.pptx -o my_assets
```

### Utilisation du script Python

```bash
# Extraire toutes les ressources prises en charge depuis un fichier PPTX
python3 extract_pptx_elements.py presentation.pptx

# Extraire vers un dossier de sortie personnalisé
python3 extract_pptx_elements.py presentation.pptx -o my_assets

# Extraire uniquement les images, les vidéos et l’audio
python3 extract_pptx_elements.py presentation.pptx --media-only

# Désactiver l’export du texte visible des diapositives
python3 extract_pptx_elements.py presentation.pptx --no-text

# Écraser les fichiers de sortie existants
python3 extract_pptx_elements.py presentation.pptx --overwrite

# Traiter tous les fichiers .pptx non temporaires du dossier courant
python3 extract_pptx_elements.py
```

Par défaut, le dossier de sortie est créé à côté du fichier PPTX et porte le nom `pptx_extracted_elements/`. Lorsque plusieurs fichiers PPTX sont traités en une seule fois, chaque fichier reçoit son propre sous-dossier.

### Structure de sortie

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

Lors du traitement de plusieurs fichiers PPTX :

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

### Champs du fichier Manifest CSV

| Champ | Description |
|-------|-------------|
| `slide` | Numéro de diapositive à trois chiffres |
| `output_file` | Chemin relatif du fichier de sortie |
| `kind` | Type de ressource, par exemple `image`, `video`, `audio`, `chart` ou `diagram` |
| `source_part` | Partie XML interne du PPTX qui contient la relation |
| `target_part` | Chemin interne de la ressource cible dans le PPTX |
| `relationship_id` | Identifiant XML de la relation |
| `relationship_type` | URI complète du type de relation |

### Fonctionnement

Un fichier PowerPoint `.pptx` est un paquet ZIP contenant des fichiers XML, des fichiers de relations et des ressources multimédias. Le processus d’extraction est le suivant :

1. Ouvrir le fichier `.pptx` comme paquet ZIP
2. Lire `ppt/presentation.xml` et ses relations pour déterminer l’ordre des diapositives
3. Parcourir l’arbre des relations de chaque diapositive afin d’identifier les images, vidéos, fichiers audio, graphiques, diagrammes et objets intégrés
4. Copier les ressources dans des dossiers de type en chinois
5. Générer les noms de fichiers avec un préfixe basé sur le numéro de diapositive
6. Extraire le texte visible par défaut ; désactiver l’export avec `--no-text`
7. Écrire `manifest.csv` avec la correspondance des sources

### À propos de `scripts/`

Le dossier `scripts/` a été supprimé. Il ne contenait que des scripts optionnels de construction PyInstaller, qui ne font pas partie du chemin d’exécution et ne sont pas nécessaires pour l’utilisation finale.

Points d’entrée recommandés :

- Python : `extract_pptx_elements.py`
- macOS : `extract_pptx_elements.command`
- Windows : `extract_pptx_elements.cmd`

Si les mainteneurs doivent générer un fichier `.exe` ou un paquet binaire, ils peuvent exécuter PyInstaller directement sur le système d’exploitation cible. Ce processus relève de la publication et n’affecte pas l’utilisation quotidienne des lanceurs.

### Notes de version

Voir les notes de version par édition dans le dossier [docs/](docs/).

### Licence

MIT
