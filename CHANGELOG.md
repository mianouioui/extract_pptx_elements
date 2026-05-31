# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.2] - 2026-05-31

### Fixed

- Fixed re-running extraction without `--overwrite` from clearing `manifest.csv`; existing assets are still protected, while the manifest is rebuilt with the expected mappings.
- Included resources referenced through slide layouts and slide masters, so layout/master images such as logos and backgrounds are no longer missed.
- Exported slide text by paragraph instead of by individual text run, avoiding unwanted line breaks when PowerPoint splits styled text into multiple runs.
- Changed restore conflict handling so multiple different edits of the same shared PPTX part stop the restore instead of silently choosing the first copy.

## [1.2.1] - 2026-05-31

### Changed

- Renamed the extractor from `PPTX 内容提取器` to `PPTX 元素提取器` (English `PPTX Content Extractor` → `PPTX Element Extractor`) across all entry points and the README.
- Gave the restorer a consistent `PPTX 元素还原器` / `PPTX Element Restorer` name in the README (the launchers already used it).
- Updated the README title to `PPTX 元素提取与还原器 / PPTX Element Extractor & Restorer`.
- Aligned the French README section to the new naming (`Extracteur d'éléments PPTX`) and added a French restorer section (`Restaurateur d'éléments PPTX`).
- Bumped all entry points (`extract_*` and `restore_*` `.py` / `.command` / `.cmd`) to 1.2.1.

## [1.2.0] - 2026-05-31

### Added

- `restore_pptx_elements.py`: new companion tool that writes extracted (and edited) media back into a copy of the original `.pptx`. Because it only swaps media bytes and never touches any XML, every element keeps its exact original position — a strict guarantee for images. It reads `manifest.csv` to map each file to its internal part, and auto-detects both the extracted folder and the original PPTX.
- `restore_pptx_elements.command`: macOS single-file launcher with embedded Python. Double-click it next to the extracted folder for a true one-click restore, or drag the folder onto the window.
- `restore_pptx_elements.cmd`: Windows single-file launcher, mirroring the extractor's entry point — embedded Python preferred, with an embedded PowerShell 5.1+ fallback so no Python install is required.
- Restore options: `--images-only`, `--media-only`, `--pptx`, `-o/--output`, `--overwrite`, `--dry-run`. An image shared across slides resolves to the edited copy, assets identical to the original are skipped, and pointing at the wrong original PPTX is detected without writing a file.

### Changed

- Bumped the whole project to 1.2.0 (extractor `.py` / `.command` / `.cmd`) for a cohesive release. The extractor's behavior is unchanged.

### Notes

- The restore PowerShell fallback in `restore_pptx_elements.cmd` is new (ported from the validated Python logic) and should be verified on a real Windows machine before relying on it for important decks. Restore never modifies the original PPTX — it always writes a separate `_restored.pptx` — so the worst case of an untested path is an unusable output file, never a lost original.

## [1.1.4] - 2026-05-31

### Changed

- Output file naming changed from `{slide}_{TAG}.ext` to `{pptx_name}_{slide}.ext` (e.g. `presentation_001.jpg` instead of `001_JPG.jpg`). Removed `tag_for`/`Get-TagFor` helper functions across all three entry points.

## [1.1.3] - 2026-05-31

### Fixed

- Fixed `--media-only` still exporting slide text — it now correctly implies `--no-text`.
- Fixed re-running without `--overwrite` silently creating `_02`/`_03` duplicate files — existing files are now skipped.
- Fixed Windows `.cmd` crash (flash exit) caused by LF line endings — converted to CRLF and added `.gitattributes` to preserve CRLF in the repository.
- Unified software name to `PPTX 内容提取器` across all entry points (Python source, macOS launcher, Windows launcher).

## [1.1.2] - 2026-05-30

### Changed

- Changed the Windows `.cmd` launcher to use embedded PowerShell extraction logic instead of requiring Python.
- Reworked the README into separate Chinese and English documentation sections, with an English jump link at the top.
- Updated the README title to `PPTX 内容提取器 / PPTX Content Extractor`.

## [1.1.1] - 2026-05-30

### Changed

- Changed the default output location to create `pptx_extracted_elements/` next to the input PPTX file, which is friendlier for drag-and-drop launcher usage.
- Made `extract_pptx_elements.cmd` a self-contained launcher with embedded Python source, so it no longer requires `extract_pptx_elements.py` beside it.
- Refreshed `extract_pptx_elements.command` with the latest embedded core logic.
- Removed the optional `scripts/` build helpers because they are not needed for the one-file quick-use workflow.
- Improved per-file error handling so one bad input does not stop other PPTX files from being processed.
- Updated README runtime documentation with more formal software operation details.

## [1.0.5] - 2026-05-30

### Changed

- Renamed the Chinese README title to `PPTX 内容提取器`.

## [1.0.4] - 2026-05-30

### Changed

- Moved build scripts into `scripts/`.
- Keep generated PyInstaller spec files under `build/specs/`.
- Extracted files are now organized into Chinese type folders such as `图片/`, `视频/`, `音频/`, `图表/`, `图示/`, `嵌入文件/`, and `文本/`.
- Completion output now reminds users that the main output folder can be found in the current folder.

## [1.0.3] - 2026-05-30

### Fixed

- `.cmd`: Windows `>/dev/null` error → `>nul` (4 occurrences)
- `.cmd` and `.command`: empty input now exits gracefully instead of looping

## [1.0.2] - 2026-05-30

### Added

- VERSION variable in `.command` and `.cmd` launchers for single-point version control

### Fixed

- macOS 14.5+ file transfer permission issue: self-healing via `chmod +x` and `xattr -d` quarantine removal
- Cleaned personal references from README

## [1.0.1] - 2026-05-30

### Added

- Self-contained `extract_pptx_elements.command` launcher with embedded Python code — zero external files needed
- `extract_pptx_elements.cmd` launcher for Windows
- Build scripts: `build_macos.sh` and `build_windows.bat`
- Bilingual Chinese/English README

### Fixed

- Intel Mac compatibility via Python source fallback in `.command` launcher

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

[1.2.2]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.2.2
[1.2.1]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.2.1
[1.2.0]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.2.0
[1.1.4]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.1.4
[1.1.3]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.1.3
[1.1.2]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.1.2
[1.1.1]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.1.1
[1.0.5]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.0.5
[1.0.4]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.0.4
[1.0.3]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.0.3
[1.0.2]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.0.2
[1.0.1]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.0.1
[1.0.0]: https://github.com/mianouioui/extract_pptx_elements/releases/tag/V1.0.0
