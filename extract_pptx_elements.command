#!/bin/bash
# ============================================================
#  PPTX 元素提取工具 V1.0.0 - macOS 独立启动器
#  只需要把这一个文件发给你女朋友，双击即可运行！
#  兼容 Intel 和 Apple Silicon Mac，无需安装任何东西。
# ============================================================
set -e

LAUNCHER_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$LAUNCHER_DIR"

# -------- 检查 Python 3 --------
if ! command -v python3 &>/dev/null; then
    echo "============================================"
    echo "  需要 Python 3，您的 Mac 尚未安装。"
    echo "  正在尝试通过 Xcode 命令行工具安装..."
    echo "============================================"
    echo ""
    xcode-select --install 2>/dev/null || true
    echo ""
    echo "如果自动安装失败，请手动安装："
    echo "  https://www.python.org/downloads/"
    echo ""
    echo "安装完成后重新双击本文件即可。"
    read -r -p "按回车关闭..." _
    exit 1
fi

PYTHON="$(command -v python3)"

# -------- 交互式提示（无参数时）--------
if [ $# -eq 0 ]; then
    echo "╔══════════════════════════════════════════╗"
    echo "║     PPTX 元素提取工具 V1.0.0            ║"
    echo "║     extract_pptx_elements               ║"
    echo "╚══════════════════════════════════════════╝"
    echo ""
    echo "用法: 将 .pptx 文件拖拽到此窗口，按回车"
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
    eval "set -- $input_args"
fi

# -------- 将内嵌的 Python 代码写入临时文件并运行 --------
PY_TEMP=$(mktemp /tmp/extract_pptx_elements_XXXXXX.py)
trap 'rm -f "$PY_TEMP"' EXIT

# 从本文件末尾提取 Python 代码（标记后）
sed '1,/^# ---PYTHON_CODE_BELOW---$/d' "$0" > "$PY_TEMP"

echo "→ 运行中..."
"$PYTHON" "$PY_TEMP" "$@"
EXIT_CODE=$?

echo ""
echo "完成！(退出码: $EXIT_CODE)"
echo "输出目录: pptx_extracted_elements/"
read -r -p "按回车关闭窗口..." _
exit $EXIT_CODE

# ---PYTHON_CODE_BELOW---
#!/usr/bin/env python3
"""
Extract slide-level resources from PowerPoint .pptx files.

Examples:
  python3 extract_pptx_elements.py "deck.pptx"
  python3 extract_pptx_elements.py "deck.pptx" -o exported_assets
  python3 extract_pptx_elements.py --with-text

Output naming:
  Slide 1 JPG: 001_JPG.jpg
  Slide 1 MP4: 001_MP4.mp4
  Second JPG on slide 1: 001_JPG_02.jpg
"""

from __future__ import annotations

import argparse
import csv
import posixpath
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable
from xml.etree import ElementTree as ET


PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

IMAGE_EXTS = {
    ".bmp",
    ".emf",
    ".gif",
    ".jfif",
    ".jpeg",
    ".jpg",
    ".png",
    ".svg",
    ".tif",
    ".tiff",
    ".webp",
    ".wmf",
}
VIDEO_EXTS = {
    ".3gp",
    ".asf",
    ".avi",
    ".m4v",
    ".mkv",
    ".mov",
    ".mp4",
    ".mpeg",
    ".mpg",
    ".swf",
    ".webm",
    ".wmv",
}
AUDIO_EXTS = {
    ".aac",
    ".aif",
    ".aiff",
    ".m4a",
    ".mid",
    ".midi",
    ".mp3",
    ".oga",
    ".ogg",
    ".wav",
    ".wma",
}
EMBED_EXTS = {
    ".bin",
    ".csv",
    ".doc",
    ".docx",
    ".html",
    ".pdf",
    ".ppt",
    ".pptx",
    ".rtf",
    ".txt",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".xml",
    ".zip",
}

REL_SKIP_WORDS = (
    "/hyperlink",
    "/notesSlide",
    "/presProps",
    "/printerSettings",
    "/slideLayout",
    "/slideMaster",
    "/theme",
    "/viewProps",
)


@dataclass(frozen=True)
class Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str


@dataclass(frozen=True)
class Resource:
    slide_number: int
    kind: str
    tag: str
    source_part: str
    target_part: str
    rel_id: str
    rel_type: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract images, videos, audio, embedded files, charts, and diagrams "
            "from .pptx files using slide-number-based names."
        )
    )
    parser.add_argument(
        "pptx",
        nargs="*",
        type=Path,
        help=(
            "PPTX file(s) to extract. If omitted, all non-temporary .pptx files "
            "in the current directory are processed."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("pptx_extracted_elements"),
        help="Output directory. Default: pptx_extracted_elements",
    )
    parser.add_argument(
        "--with-text",
        action="store_true",
        help="Also export plain slide text as 001_TXT.txt, 002_TXT.txt, etc.",
    )
    parser.add_argument(
        "--media-only",
        action="store_true",
        help="Only extract images, videos, and audio.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite files in the output directory if they already exist.",
    )
    return parser.parse_args()


def is_temp_pptx(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".~")


def find_input_files(args: argparse.Namespace) -> list[Path]:
    if args.pptx:
        return [path.expanduser().resolve() for path in args.pptx]

    return sorted(
        path.resolve()
        for path in Path.cwd().glob("*.pptx")
        if path.is_file() and not is_temp_pptx(path)
    )


def read_xml(zip_file: zipfile.ZipFile, part: str) -> ET.Element | None:
    try:
        return ET.fromstring(zip_file.read(part))
    except KeyError:
        return None
    except ET.ParseError as exc:
        print(f"Warning: could not parse XML part {part}: {exc}", file=sys.stderr)
        return None


def rels_part_name(part: str) -> str:
    directory = posixpath.dirname(part)
    filename = posixpath.basename(part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def parse_rels(zip_file: zipfile.ZipFile, part: str) -> dict[str, Relationship]:
    root = read_xml(zip_file, rels_part_name(part))
    if root is None:
        return {}

    rels: dict[str, Relationship] = {}
    for item in root.findall(f"{{{PACKAGE_RELS_NS}}}Relationship"):
        rel_id = item.get("Id", "")
        if not rel_id:
            continue

        rels[rel_id] = Relationship(
            rel_id=rel_id,
            rel_type=item.get("Type", ""),
            target=item.get("Target", ""),
            target_mode=item.get("TargetMode", ""),
        )
    return rels


def resolve_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")

    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def natural_slide_sort_key(part: str) -> tuple[int, str]:
    match = re.search(r"/slide(\d+)\.xml$", part)
    if match:
        return (int(match.group(1)), part)
    return (10**9, part)


def slide_parts_in_order(zip_file: zipfile.ZipFile) -> list[str]:
    presentation = read_xml(zip_file, "ppt/presentation.xml")
    presentation_rels = parse_rels(zip_file, "ppt/presentation.xml")
    names = set(zip_file.namelist())

    if presentation is not None and presentation_rels:
        ordered_slides: list[str] = []
        for slide_id in presentation.findall(
            f".//{{{PRESENTATION_NS}}}sldIdLst/{{{PRESENTATION_NS}}}sldId"
        ):
            rel_id = slide_id.get(f"{{{OFFICE_RELS_NS}}}id")
            if not rel_id or rel_id not in presentation_rels:
                continue
            target = resolve_target(
                "ppt/presentation.xml", presentation_rels[rel_id].target
            )
            if target in names:
                ordered_slides.append(target)

        if ordered_slides:
            return ordered_slides

    return sorted(
        (
            name
            for name in names
            if name.startswith("ppt/slides/") and name.endswith(".xml")
        ),
        key=natural_slide_sort_key,
    )


def extension_for(part: str) -> str:
    return PurePosixPath(part).suffix.lower()


def tag_for(kind: str, part: str) -> str:
    ext = extension_for(part).lstrip(".").upper()
    if ext == "JPEG":
        return "JPG"
    if kind == "chart":
        return "CHART"
    if kind == "chart_style":
        return "CHARTSTYLE"
    if kind == "chart_colors":
        return "CHARTCOLORS"
    if kind == "diagram":
        return "DIAGRAM"
    if kind == "ole":
        return "OLE"
    if kind == "unknown":
        return ext or "FILE"
    return ext or kind.upper()


def output_suffix_for(part: str, tag: str) -> str:
    ext = extension_for(part)
    if ext == ".jpeg":
        return ".jpg"
    if ext:
        return ext
    return f".{tag.lower()}"


def classify_part(rel_type: str, part: str) -> str | None:
    lower_rel_type = rel_type.lower()
    lower_part = part.lower()
    ext = extension_for(part)

    if any(word.lower() in lower_rel_type for word in REL_SKIP_WORDS):
        return None

    if ext in IMAGE_EXTS or "/media/image" in lower_part or lower_rel_type.endswith("/image"):
        return "image"
    if ext in VIDEO_EXTS or "video" in lower_rel_type:
        return "video"
    if ext in AUDIO_EXTS or "audio" in lower_rel_type:
        return "audio"
    if "/embeddings/" in lower_part or lower_rel_type.endswith("/package"):
        return "ole" if ext == ".bin" else "embedded"
    if "/charts/style" in lower_part:
        return "chart_style"
    if "/charts/colors" in lower_part:
        return "chart_colors"
    if "/charts/" in lower_part or lower_rel_type.endswith("/chart"):
        return "chart"
    if "/diagrams/" in lower_part:
        return "diagram"

    return None


def should_extract(kind: str, media_only: bool) -> bool:
    if media_only:
        return kind in {"image", "video", "audio"}
    return kind in {
        "audio",
        "chart",
        "chart_colors",
        "chart_style",
        "diagram",
        "embedded",
        "image",
        "ole",
        "video",
    }


def collect_slide_resources(
    zip_file: zipfile.ZipFile,
    slide_part: str,
    slide_number: int,
    *,
    media_only: bool,
) -> list[Resource]:
    names = set(zip_file.namelist())
    resources: list[Resource] = []
    seen_targets: set[str] = set()

    def walk_relationships(source_part: str, depth: int) -> None:
        if depth > 2:
            return

        for rel in parse_rels(zip_file, source_part).values():
            if not rel.target or rel.target_mode.lower() == "external":
                continue

            target_part = resolve_target(source_part, rel.target)
            if target_part not in names:
                continue

            kind = classify_part(rel.rel_type, target_part)
            if kind and should_extract(kind, media_only) and target_part not in seen_targets:
                seen_targets.add(target_part)
                resources.append(
                    Resource(
                        slide_number=slide_number,
                        kind=kind,
                        tag=tag_for(kind, target_part),
                        source_part=source_part,
                        target_part=target_part,
                        rel_id=rel.rel_id,
                        rel_type=rel.rel_type,
                    )
                )

            if kind in {"chart", "diagram"} and not media_only:
                walk_relationships(target_part, depth + 1)

    walk_relationships(slide_part, 0)
    return resources


def extract_slide_text(zip_file: zipfile.ZipFile, slide_part: str) -> str:
    root = read_xml(zip_file, slide_part)
    if root is None:
        return ""

    text_runs: list[str] = []
    for item in root.iter(f"{{{DRAWING_NS}}}t"):
        if item.text:
            text = item.text.strip()
            if text:
                text_runs.append(text)

    return "\n".join(text_runs)


def unique_output_path(
    output_dir: Path,
    slide_number: int,
    tag: str,
    suffix: str,
    counters: dict[tuple[int, str], int],
    *,
    overwrite: bool,
) -> Path:
    key = (slide_number, tag)
    counters[key] = counters.get(key, 0) + 1
    index = counters[key]
    stem = f"{slide_number:03d}_{tag}" if index == 1 else f"{slide_number:03d}_{tag}_{index:02d}"
    candidate = output_dir / f"{stem}{suffix}"

    if overwrite:
        return candidate

    collision_index = index
    while candidate.exists():
        collision_index += 1
        candidate = output_dir / f"{slide_number:03d}_{tag}_{collision_index:02d}{suffix}"

    return candidate


def extract_file(
    zip_file: zipfile.ZipFile,
    member: str,
    destination: Path,
    *,
    overwrite: bool,
) -> None:
    if destination.exists() and not overwrite:
        raise FileExistsError(destination)

    destination.parent.mkdir(parents=True, exist_ok=True)
    with zip_file.open(member) as source, destination.open("wb") as target:
        shutil.copyfileobj(source, target)


def output_dir_for(pptx_path: Path, base_output_dir: Path, multi_input: bool) -> Path:
    if multi_input:
        return base_output_dir / pptx_path.stem
    return base_output_dir


def write_manifest(manifest_path: Path, rows: list[dict[str, str]]) -> None:
    fieldnames = [
        "slide",
        "output_file",
        "kind",
        "source_part",
        "target_part",
        "relationship_id",
        "relationship_type",
    ]
    with manifest_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def extract_pptx(
    pptx_path: Path,
    output_dir: Path,
    *,
    media_only: bool,
    with_text: bool,
    overwrite: bool,
) -> tuple[int, int]:
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)
    if not zipfile.is_zipfile(pptx_path):
        raise ValueError(f"Not a valid .pptx/zip file: {pptx_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    counters: dict[tuple[int, str], int] = {}
    manifest_rows: list[dict[str, str]] = []
    extracted_count = 0

    with zipfile.ZipFile(pptx_path) as pptx_zip:
        slide_parts = slide_parts_in_order(pptx_zip)

        for slide_number, slide_part in enumerate(slide_parts, start=1):
            resources = collect_slide_resources(
                pptx_zip,
                slide_part,
                slide_number,
                media_only=media_only,
            )

            for resource in resources:
                suffix = output_suffix_for(resource.target_part, resource.tag)
                destination = unique_output_path(
                    output_dir,
                    resource.slide_number,
                    resource.tag,
                    suffix,
                    counters,
                    overwrite=overwrite,
                )
                extract_file(
                    pptx_zip,
                    resource.target_part,
                    destination,
                    overwrite=overwrite,
                )
                extracted_count += 1
                manifest_rows.append(
                    {
                        "slide": f"{resource.slide_number:03d}",
                        "output_file": destination.name,
                        "kind": resource.kind,
                        "source_part": resource.source_part,
                        "target_part": resource.target_part,
                        "relationship_id": resource.rel_id,
                        "relationship_type": resource.rel_type,
                    }
                )

            if with_text:
                slide_text = extract_slide_text(pptx_zip, slide_part)
                if slide_text:
                    text_path = unique_output_path(
                        output_dir,
                        slide_number,
                        "TXT",
                        ".txt",
                        counters,
                        overwrite=overwrite,
                    )
                    text_path.write_text(slide_text + "\n", encoding="utf-8")
                    extracted_count += 1
                    manifest_rows.append(
                        {
                            "slide": f"{slide_number:03d}",
                            "output_file": text_path.name,
                            "kind": "text",
                            "source_part": slide_part,
                            "target_part": slide_part,
                            "relationship_id": "",
                            "relationship_type": "",
                        }
                    )

    write_manifest(output_dir / "manifest.csv", manifest_rows)
    return len(slide_parts), extracted_count


def main() -> int:
    args = parse_args()
    input_files = find_input_files(args)

    if not input_files:
        print("No .pptx files found. Pass a file path or run this in a PPTX folder.", file=sys.stderr)
        return 1

    output_base = args.output.expanduser().resolve()
    multi_input = len(input_files) > 1

    for pptx_path in input_files:
        if is_temp_pptx(pptx_path):
            print(f"Skip temporary PowerPoint lock file: {pptx_path.name}")
            continue

        destination_dir = output_dir_for(pptx_path, output_base, multi_input)
        slide_count, extracted_count = extract_pptx(
            pptx_path,
            destination_dir,
            media_only=args.media_only,
            with_text=args.with_text,
            overwrite=args.overwrite,
        )
        print(
            f"{pptx_path.name}: {slide_count} slides, "
            f"{extracted_count} files -> {destination_dir}"
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
