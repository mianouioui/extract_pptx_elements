# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-05-30

### Added

- Extract images (JPG, PNG, GIF, SVG, BMP, EMF, WMF, TIFF, WebP, JFIF) from .pptx files
- Extract videos (MP4, AVI, MOV, MKV, WebM, WMV, etc.) from .pptx files
- Extract audio (MP3, WAV, AAC, M4A, OGG, MIDI, WMA, etc.) from .pptx files
- Extract charts (chart XML + style and color definitions) from .pptx files
- Extract diagrams (SmartArt XML) from .pptx files
- Extract embedded / OLE objects (PDF, DOCX, XLSX, ZIP, txt, etc.) from .pptx files
- Optional slide text extraction via `--with-text` flag
- Slide-number-based output naming (`001_JPG.jpg`, `002_MP4.mp4`, etc.)
- Automatic duplicate file handling (`001_JPG_02.jpg`, `001_JPG_03.jpg`, etc.)
- `manifest.csv` generation for full extraction provenance
- `--media-only` flag to extract only images, videos, and audio
- `--overwrite` flag to overwrite existing output files
- Custom output directory via `-o` / `--output`
- Multi-file processing with per-file output subdirectories
- Automatic discovery of all .pptx files in the working directory
- Skip temporary PowerPoint lock files (`~$` prefix)
- Zero external dependencies (Python stdlib only)
- Support for Python 3.8+

[1.0.0]: https://github.com/yourusername/extract_pptx_elements/releases/tag/V1.0.0
