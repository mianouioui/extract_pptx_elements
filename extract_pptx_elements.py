#!/usr/bin/env python3
"""
PPTX 内容提取器 - Extract slide-level resources from PowerPoint .pptx files.

Examples:
  python3 extract_pptx_elements.py "deck.pptx"
  python3 extract_pptx_elements.py "deck.pptx" -o exported_assets
  python3 extract_pptx_elements.py --no-text

Output naming:
  Slide 1 JPG: 图片/presentation_001.jpg
  Slide 1 MP4: 视频/presentation_001.mp4
  Second JPG on slide 1: 图片/presentation_001_02.jpg
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

VERSION = "1.2.0"
DEFAULT_OUTPUT_DIR_NAME = "pptx_extracted_elements"

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

KIND_FOLDER_NAMES = {
    "audio": "音频",
    "chart": "图表",
    "chart_colors": "图表",
    "chart_style": "图表",
    "diagram": "图示",
    "embedded": "嵌入文件",
    "image": "图片",
    "ole": "嵌入文件",
    "text": "文本",
    "video": "视频",
}


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
    source_part: str
    target_part: str
    rel_id: str
    rel_type: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "PPTX 内容提取器 - Extract images, videos, audio, embedded files, charts, "
            "and diagrams from .pptx files into Chinese type folders using "
            "slide-number-based names."
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
        default=None,
        help=(
            "Output directory. Default: create pptx_extracted_elements next to "
            "the PPTX file being processed."
        ),
    )
    text_group = parser.add_mutually_exclusive_group()
    text_group.add_argument(
        "--with-text",
        dest="with_text",
        action="store_true",
        default=True,
        help="Export plain slide text as 001_TXT.txt, 002_TXT.txt, etc. Enabled by default.",
    )
    text_group.add_argument(
        "--no-text",
        dest="with_text",
        action="store_false",
        help="Do not export plain slide text.",
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
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
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


def output_suffix_for(part: str) -> str:
    ext = extension_for(part)
    if ext == ".jpeg":
        return ".jpg"
    if ext:
        return ext
    return ".bin"


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
    pptx_stem: str,
    slide_number: int,
    suffix: str,
    counters: dict[tuple[Path, int], int],
    *,
    overwrite: bool,
) -> Path | None:
    key = (output_dir, slide_number)
    counters[key] = counters.get(key, 0) + 1
    index = counters[key]
    stem = f"{pptx_stem}_{slide_number:03d}" if index == 1 else f"{pptx_stem}_{slide_number:03d}_{index:02d}"
    candidate = output_dir / f"{stem}{suffix}"

    if overwrite:
        return candidate

    if candidate.exists():
        return None

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


def output_dir_for(
    pptx_path: Path,
    output_arg: Path | None,
    multi_input: bool,
) -> Path:
    if output_arg is None:
        base_output_dir = pptx_path.parent / DEFAULT_OUTPUT_DIR_NAME
    else:
        base_output_dir = output_arg.expanduser()
        if not base_output_dir.is_absolute():
            base_output_dir = Path.cwd() / base_output_dir

    if multi_input:
        return base_output_dir / pptx_path.stem
    return base_output_dir


def kind_output_dir(output_dir: Path, kind: str) -> Path:
    return output_dir / KIND_FOLDER_NAMES.get(kind, "其他")


def relative_output_file(path: Path, output_dir: Path) -> str:
    return path.relative_to(output_dir).as_posix()


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
        raise ValueError(f"不是有效的 .pptx/zip 文件：{pptx_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    pptx_stem = pptx_path.stem
    counters: dict[tuple[Path, int], int] = {}
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
                suffix = output_suffix_for(resource.target_part)
                resource_output_dir = kind_output_dir(output_dir, resource.kind)
                destination = unique_output_path(
                    resource_output_dir,
                    pptx_stem,
                    resource.slide_number,
                    suffix,
                    counters,
                    overwrite=overwrite,
                )
                if destination is None:
                    continue
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
                        "output_file": relative_output_file(destination, output_dir),
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
                        kind_output_dir(output_dir, "text"),
                        pptx_stem,
                        slide_number,
                        ".txt",
                        counters,
                        overwrite=overwrite,
                    )
                    if text_path is not None:
                        text_path.parent.mkdir(parents=True, exist_ok=True)
                        text_path.write_text(slide_text + "\n", encoding="utf-8")
                        extracted_count += 1
                        manifest_rows.append(
                            {
                                "slide": f"{slide_number:03d}",
                                "output_file": relative_output_file(text_path, output_dir),
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
        print(
            "未找到 .pptx 文件。请拖入/传入 PPTX 文件，或把本工具放到 PPTX 所在文件夹运行。",
            file=sys.stderr,
        )
        return 1

    multi_input = len(input_files) > 1
    output_dirs: list[Path] = []
    success_count = 0
    failure_count = 0

    for pptx_path in input_files:
        if is_temp_pptx(pptx_path):
            print(f"跳过 PowerPoint 临时锁文件：{pptx_path.name}")
            continue

        destination_dir = output_dir_for(pptx_path, args.output, multi_input).resolve()
        try:
            slide_count, extracted_count = extract_pptx(
                pptx_path,
                destination_dir,
                media_only=args.media_only,
                with_text=args.with_text and not args.media_only,
                overwrite=args.overwrite,
            )
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            failure_count += 1
            print(f"提取失败：{pptx_path} -> {exc}", file=sys.stderr)
            continue

        success_count += 1
        output_dirs.append(destination_dir)
        print(
            f"{pptx_path.name}: {slide_count} slides, "
            f"{extracted_count} files -> {destination_dir}"
        )

    if output_dirs:
        unique_output_dirs = list(dict.fromkeys(output_dirs))
        if len(unique_output_dirs) == 1:
            print(f"提醒：输出文件夹在这里：{unique_output_dirs[0]}")
        else:
            print("提醒：输出文件夹在这里：")
            for output_dir in unique_output_dirs:
                print(f"  - {output_dir}")

    if success_count == 0:
        return 1

    return 1 if failure_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
