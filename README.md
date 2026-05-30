# extract_pptx_elements

Extract slide-level resources (images, videos, audio, charts, diagrams, embedded files, and text) from PowerPoint `.pptx` files with clean, slide-number-based file naming.

## Features

- **Images** — JPG, PNG, GIF, SVG, BMP, EMF, WMF, TIFF, WebP, JFIF
- **Videos** — MP4, AVI, MOV, MKV, WebM, WMV, etc.
- **Audio** — MP3, WAV, AAC, M4A, OGG, MIDI, WMA, etc.
- **Charts** — chart XML + associated style and color definitions
- **Diagrams** — SmartArt diagram XML
- **Embedded / OLE objects** — PDF, DOCX, XLSX, ZIP, etc.
- **Slide text** — optional plain-text extraction via `--with-text`
- **Manifest** — auto-generated `manifest.csv` maps every output file back to its slide and source

Output files are named like `001_JPG.jpg`, `002_MP4.mp4`, `003_CHART.xml` — the 3-digit prefix matches the slide number.

## Requirements

- Python 3.8+
- No third-party dependencies (stdlib only: `zipfile`, `xml.etree`, `argparse`, `csv`)

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/extract_pptx_elements.git
cd extract_pptx_elements

# Make the script executable (optional)
chmod +x extract_pptx_elements.py
```

## Usage

```bash
# Extract everything from a single .pptx file
python3 extract_pptx_elements.py presentation.pptx

# Extract to a custom output directory
python3 extract_pptx_elements.py presentation.pptx -o my_assets/

# Only images, videos, and audio
python3 extract_pptx_elements.py presentation.pptx --media-only

# Also extract slide text as 001_TXT.txt, 002_TXT.txt, etc.
python3 extract_pptx_elements.py presentation.pptx --with-text

# Overwrite existing files
python3 extract_pptx_elements.py presentation.pptx --overwrite

# Process all .pptx files in the current directory
python3 extract_pptx_elements.py
```

## Output Structure

```
pptx_extracted_elements/
├── 001_JPG.jpg          # Slide 1, first image
├── 001_JPG_02.jpg       # Slide 1, second image
├── 002_MP4.mp4          # Slide 2, video
├── 002_PNG.png          # Slide 2, image
├── 003_CHART.xml        # Slide 3, chart
├── 003_TXT.txt          # Slide 3, text (with --with-text)
├── manifest.csv         # Full inventory of extracted files
│
# When processing multiple .pptx files, each gets its own subfolder:
pptx_extracted_elements/
├── presentation1/
│   ├── 001_JPG.jpg
│   └── manifest.csv
└── presentation2/
    ├── 001_PNG.png
    └── manifest.csv
```

## Manifest CSV

The `manifest.csv` file contains:

| Column            | Description                              |
|-------------------|------------------------------------------|
| slide             | 3-digit slide number                     |
| output_file       | Extracted file name                      |
| kind              | image, video, audio, chart, diagram, etc.|
| source_part       | Internal path in the .pptx zip           |
| target_part       | Internal path of the resource            |
| relationship_id   | XML relationship ID                      |
| relationship_type | Full relationship type URI               |

## How It Works

PowerPoint `.pptx` files are ZIP archives containing XML and media files. This tool:

1. Opens the `.pptx` as a ZIP archive
2. Reads `ppt/presentation.xml` to determine slide order
3. Walks each slide's relationship tree to find images, videos, audio, charts, diagrams, and embedded objects
4. Extracts each resource with a slide-prefixed filename
5. Optionally extracts visible text from slide XML
6. Writes a `manifest.csv` with full provenance information

## License

MIT
