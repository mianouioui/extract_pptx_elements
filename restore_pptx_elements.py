#!/usr/bin/env python3
"""
PPTX 元素还原器 - Rebuild a .pptx from extracted media without moving anything.

This is the companion to extract_pptx_elements.py. It reads the manifest.csv
produced during extraction and writes each (possibly edited) media file back
into its exact part inside a *copy* of the original .pptx.

Why positions never change:
  An element's position and size live in the slide XML (<a:off>/<a:ext>), not
  in the image bytes. This tool only swaps the bytes of each media part and
  never touches any XML, so every element keeps its original position exactly.
  This is a strict guarantee for images and, in practice, for every other
  element too.

Typical workflow:
  1. extract_pptx_elements.py "deck.pptx"        -> pptx_extracted_elements/
  2. Batch-edit images in 图片/ (e.g. remove a watermark), keep the file names.
  3. restore_pptx_elements.py                     -> deck_restored.pptx

Examples:
  python3 restore_pptx_elements.py
  python3 restore_pptx_elements.py pptx_extracted_elements
  python3 restore_pptx_elements.py pptx_extracted_elements --pptx deck.pptx
  python3 restore_pptx_elements.py pptx_extracted_elements -o deck_fixed.pptx
  python3 restore_pptx_elements.py --images-only
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Iterable


VERSION = "1.2.0"

MANIFEST_NAME = "manifest.csv"
DEFAULT_EXTRACT_DIR_NAME = "pptx_extracted_elements"
DEFAULT_OUTPUT_SUFFIX = "_restored"

# Kinds whose extracted file is the literal part inside the pptx and can be
# safely written back byte-for-byte. "text" is intentionally excluded: its
# extracted .txt is a lossy plain-text dump of the slide XML and must never be
# written over the slide.
RESTORABLE_KINDS = {
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
MEDIA_KINDS = {"image", "video", "audio"}

# Magic-number signatures used only for a soft "format changed" warning.
IMAGE_SIGNATURES = {
    ".png": (b"\x89PNG\r\n\x1a\n",),
    ".jpg": (b"\xff\xd8\xff",),
    ".jpeg": (b"\xff\xd8\xff",),
    ".gif": (b"GIF87a", b"GIF89a"),
    ".bmp": (b"BM",),
    ".tif": (b"II*\x00", b"MM\x00*"),
    ".tiff": (b"II*\x00", b"MM\x00*"),
    ".webp": (b"RIFF",),
}


@dataclass(frozen=True)
class ManifestRow:
    slide: str
    output_file: str
    kind: str
    source_part: str
    target_part: str
    rel_id: str
    rel_type: str


@dataclass(frozen=True)
class Replacement:
    target_part: str
    source_file: Path
    kind: str


@dataclass
class RestorePlan:
    # target_part -> file that should replace it (only parts that truly change)
    replacements: dict[str, Replacement] = field(default_factory=dict)
    # manifest rows whose file is no longer on disk (original bytes kept)
    missing_files: list[ManifestRow] = field(default_factory=list)
    # target parts present on disk but absent from the pptx (wrong pptx?)
    missing_in_zip: list[str] = field(default_factory=list)
    # parts whose on-disk file is byte-identical to the original (no edit)
    unchanged: list[str] = field(default_factory=list)
    # parts with several edited copies that disagree (target, winner, other)
    conflicts: list[tuple[str, Path, Path]] = field(default_factory=list)
    # distinct in-scope target parts that exist inside the pptx
    present_targets: int = 0

    @property
    def total_targets(self) -> int:
        return self.present_targets + len(self.missing_in_zip)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "PPTX 元素还原器 - Write extracted (and possibly edited) media back "
            "into a copy of the original .pptx, keeping every element in its "
            "original position."
        )
    )
    parser.add_argument(
        "folder",
        nargs="?",
        type=Path,
        default=None,
        help=(
            "Extracted folder containing manifest.csv (the output of "
            "extract_pptx_elements). If omitted, the tool looks for "
            f"'{DEFAULT_EXTRACT_DIR_NAME}' or a manifest.csv in the current folder."
        ),
    )
    parser.add_argument(
        "--pptx",
        type=Path,
        default=None,
        help=(
            "Original .pptx to rebuild from. If omitted, it is auto-detected "
            "next to the extracted folder using the file name recorded in the "
            "manifest."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help=(
            "Output .pptx path. Default: '<original>_restored.pptx' next to the "
            "original file."
        ),
    )
    scope_group = parser.add_mutually_exclusive_group()
    scope_group.add_argument(
        "--images-only",
        action="store_true",
        help="Only write images back; leave every other element untouched.",
    )
    scope_group.add_argument(
        "--media-only",
        action="store_true",
        help="Only write images, videos, and audio back.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite the output .pptx if it already exists.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be written without creating the output file.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
    )
    return parser.parse_args(argv)


def is_temp_pptx(path: Path) -> bool:
    return path.name.startswith("~$") or path.name.startswith(".~")


def find_extract_dir(folder_arg: Path | None) -> Path:
    """Resolve the folder that holds manifest.csv."""
    candidates: list[Path] = []
    if folder_arg is not None:
        candidates.append(folder_arg.expanduser())
    else:
        cwd = Path.cwd()
        candidates.append(cwd / DEFAULT_EXTRACT_DIR_NAME)
        candidates.append(cwd)

    for candidate in candidates:
        if (candidate / MANIFEST_NAME).is_file():
            return candidate.resolve()

    # A folder was given but the manifest is one level down (multi-input layout:
    # <out>/<deck-name>/manifest.csv with a single subfolder).
    if folder_arg is not None:
        base = folder_arg.expanduser()
        if base.is_dir():
            subdirs = [p for p in base.iterdir() if (p / MANIFEST_NAME).is_file()]
            if len(subdirs) == 1:
                return subdirs[0].resolve()

    searched = "\n".join(f"  - {c / MANIFEST_NAME}" for c in candidates)
    raise FileNotFoundError(
        "找不到 manifest.csv（提取时生成的清单）。已查找：\n"
        f"{searched}\n"
        "请把提取得到的文件夹拖到本工具，或在该文件夹内运行。"
    )


def read_manifest(manifest_path: Path) -> list[ManifestRow]:
    rows: list[ManifestRow] = []
    with manifest_path.open("r", newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.DictReader(csv_file)
        for raw in reader:
            rows.append(
                ManifestRow(
                    slide=(raw.get("slide") or "").strip(),
                    output_file=(raw.get("output_file") or "").strip(),
                    kind=(raw.get("kind") or "").strip(),
                    source_part=(raw.get("source_part") or "").strip(),
                    target_part=(raw.get("target_part") or "").strip(),
                    rel_id=(raw.get("relationship_id") or "").strip(),
                    rel_type=(raw.get("relationship_type") or "").strip(),
                )
            )
    return rows


def derive_pptx_stem(rows: Iterable[ManifestRow]) -> str | None:
    """Recover the original pptx file stem from output file names.

    Output files are named '<stem>_<NNN>[_<NN>].ext'. Strip the trailing slide
    number (and optional duplicate index) and return the most common stem.
    """
    counts: dict[str, int] = {}
    pattern = re.compile(r"_\d{3}(?:_\d{2})?$")
    for row in rows:
        if not row.output_file:
            continue
        stem = pattern.sub("", PurePosixPath(row.output_file).stem)
        if stem:
            counts[stem] = counts.get(stem, 0) + 1
    if not counts:
        return None
    return max(counts, key=lambda key: (counts[key], key))


def find_source_pptx(
    extract_dir: Path,
    explicit: Path | None,
    stem: str | None,
) -> Path:
    if explicit is not None:
        path = explicit.expanduser()
        if not path.is_file():
            raise FileNotFoundError(f"指定的 PPTX 不存在：{path}")
        return path.resolve()

    search_dirs = [
        extract_dir.parent,         # default layout: folder sits next to the pptx
        extract_dir,
        Path.cwd(),
        extract_dir.parent.parent,  # multi-input layout
    ]
    seen: set[Path] = set()

    if stem:
        for directory in search_dirs:
            directory = directory.resolve()
            if directory in seen:
                continue
            seen.add(directory)
            candidate = directory / f"{stem}.pptx"
            if candidate.is_file() and not is_temp_pptx(candidate):
                return candidate.resolve()

    # Fallback: a single .pptx sitting next to the extracted folder.
    for directory in (extract_dir.parent, extract_dir):
        pptxs = [
            p for p in directory.glob("*.pptx") if p.is_file() and not is_temp_pptx(p)
        ]
        if len(pptxs) == 1:
            return pptxs[0].resolve()

    hint = f"{stem}.pptx" if stem else "原始 PPTX"
    raise FileNotFoundError(
        f"找不到原始 PPTX（{hint}）。\n"
        "请用 --pptx 指定原始文件，或把它放在提取文件夹旁边。"
    )


def default_output_path(source_pptx: Path) -> Path:
    return source_pptx.with_name(f"{source_pptx.stem}{DEFAULT_OUTPUT_SUFFIX}.pptx")


def kinds_in_scope(*, images_only: bool, media_only: bool) -> set[str]:
    if images_only:
        return {"image"}
    if media_only:
        return set(MEDIA_KINDS)
    return set(RESTORABLE_KINDS)


def signature_mismatch(target_part: str, data: bytes) -> bool:
    """Return True only when an image part's bytes clearly don't match its ext."""
    ext = PurePosixPath(target_part).suffix.lower()
    signatures = IMAGE_SIGNATURES.get(ext)
    if not signatures:
        return False
    return not any(data.startswith(sig) for sig in signatures)


def build_plan(
    extract_dir: Path,
    rows: list[ManifestRow],
    scope: set[str],
    zin: zipfile.ZipFile,
) -> RestorePlan:
    """Decide which parts to swap, reading the original pptx for comparison.

    The same part can be referenced by several slides (a shared logo/background).
    Among the on-disk copies of one part, the copy that actually differs from the
    original wins, so editing any single copy is enough. Copies identical to the
    original are reported as unchanged; genuinely conflicting edits are flagged.
    """
    names = set(zin.namelist())
    plan = RestorePlan()
    original_cache: dict[str, bytes] = {}

    # Group on-disk files by their target part, preserving manifest order.
    groups: dict[str, list[Path]] = {}
    kinds: dict[str, str] = {}
    for row in rows:
        if row.kind == "text" or row.kind not in scope:
            continue
        if not row.target_part or not row.output_file:
            continue
        source_file = extract_dir / row.output_file
        if not source_file.is_file():
            plan.missing_files.append(row)
            continue
        groups.setdefault(row.target_part, []).append(source_file)
        kinds.setdefault(row.target_part, row.kind)

    for target, files in groups.items():
        if target not in names:
            plan.missing_in_zip.append(target)
            continue
        plan.present_targets += 1
        original_size = zin.getinfo(target).file_size

        def is_changed(path: Path) -> bool:
            if path.stat().st_size != original_size:
                return True
            if target not in original_cache:
                original_cache[target] = zin.read(target)
            return path.read_bytes() != original_cache[target]

        changed_files = [path for path in files if is_changed(path)]
        if not changed_files:
            plan.unchanged.append(target)
            continue

        winner = changed_files[0]
        if len(changed_files) > 1:
            # Only read the winner's bytes when there are sibling copies to
            # check against; the common single-copy case needs no read here.
            winner_bytes = winner.read_bytes()
            for other in changed_files[1:]:
                if other.read_bytes() != winner_bytes:
                    plan.conflicts.append((target, winner, other))
                    break

        plan.replacements[target] = Replacement(target, winner, kinds[target])

    return plan


def rewrite_pptx(
    source_pptx: Path,
    replacements: dict[str, Replacement],
    output_path: Path,
) -> list[str]:
    """Copy source_pptx to output_path, swapping bytes for replaced parts.

    Every other part - including all slide XML - is copied verbatim, so every
    element keeps its original position and size.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_name(output_path.name + ".tmp")
    replaced: list[str] = []

    with zipfile.ZipFile(source_pptx) as zin:
        try:
            with zipfile.ZipFile(temp_path, "w") as zout:
                for item in zin.infolist():
                    replacement = replacements.get(item.filename)
                    if replacement is not None:
                        # Reuse the original ZipInfo so compress_type, date_time
                        # and flags are preserved for the swapped part. Read the
                        # replacement just-in-time so at most one media file is
                        # held in memory (matters for large video/audio parts).
                        zout.writestr(item, replacement.source_file.read_bytes())
                        replaced.append(item.filename)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        except BaseException:
            temp_path.unlink(missing_ok=True)
            raise

    temp_path.replace(output_path)
    return replaced


def warn_signature_mismatches(plan: RestorePlan) -> None:
    for target, repl in plan.replacements.items():
        with repl.source_file.open("rb") as handle:
            head = handle.read(16)
        if signature_mismatch(target, head):
            print(
                f"  ⚠ 注意：{repl.source_file.name} 的实际格式可能与 "
                f"{PurePosixPath(target).suffix} 不一致，PowerPoint 仍会尝试显示。",
                file=sys.stderr,
            )


def print_report(
    *,
    source_pptx: Path,
    output_path: Path,
    plan: RestorePlan,
    dry_run: bool,
) -> None:
    by_kind: dict[str, int] = {}
    for repl in plan.replacements.values():
        by_kind[repl.kind] = by_kind.get(repl.kind, 0) + 1

    verb = "将写回" if dry_run else "已写回"
    if by_kind:
        summary = "，".join(f"{kind} {count}" for kind, count in sorted(by_kind.items()))
        print(f"{verb} {len(plan.replacements)} 个元素（{summary}）。")
    else:
        print("未检测到与原始不同的素材（输出与原始内容一致）。")

    warn_signature_mismatches(plan)

    if plan.unchanged:
        print(f"  · {len(plan.unchanged)} 个素材与原始一致，保持不变。")
    if plan.missing_files:
        print(
            f"  · {len(plan.missing_files)} 个清单项在文件夹中找不到对应文件，"
            "保留原始内容。"
        )
    if plan.missing_in_zip:
        print(
            f"  ⚠ {len(plan.missing_in_zip)} 个部件在原始 PPTX 中不存在，已跳过。"
            "请确认 --pptx 指向的是同一个文件。",
            file=sys.stderr,
        )
    for target, first, second in plan.conflicts:
        print(
            f"  ⚠ 同一部件 {target} 有多个不同的修改版本："
            f"{first.name} / {second.name}，采用第一个。",
            file=sys.stderr,
        )

    if dry_run:
        print(f"（预演）将生成：{output_path}")
    else:
        print(f"完成 ✅  已生成：{output_path}")
        print(f"原始文件未改动：{source_pptx}")


def restore(
    *,
    folder_arg: Path | None,
    pptx_arg: Path | None,
    output_arg: Path | None,
    images_only: bool,
    media_only: bool,
    overwrite: bool,
    dry_run: bool,
) -> int:
    extract_dir = find_extract_dir(folder_arg)
    manifest_path = extract_dir / MANIFEST_NAME
    rows = read_manifest(manifest_path)
    if not rows:
        print(f"清单为空：{manifest_path}", file=sys.stderr)
        return 1

    stem = derive_pptx_stem(rows)
    source_pptx = find_source_pptx(extract_dir, pptx_arg, stem)
    if not zipfile.is_zipfile(source_pptx):
        print(f"不是有效的 .pptx/zip 文件：{source_pptx}", file=sys.stderr)
        return 1

    if output_arg is not None:
        output_path = output_arg.expanduser()
        if not output_path.is_absolute():
            output_path = (Path.cwd() / output_path).resolve()
    else:
        output_path = default_output_path(source_pptx)

    if not dry_run and output_path.exists() and not overwrite:
        print(
            f"输出文件已存在：{output_path}\n"
            "使用 --overwrite 覆盖，或用 -o 指定其它文件名。",
            file=sys.stderr,
        )
        return 1

    scope = kinds_in_scope(images_only=images_only, media_only=media_only)
    with zipfile.ZipFile(source_pptx) as zin:
        plan = build_plan(extract_dir, rows, scope, zin)

    print(f"原始 PPTX：{source_pptx}")
    print(f"素材文件夹：{extract_dir}")

    if plan.total_targets == 0:
        print(
            "没有找到任何可写回的素材文件。请确认筛选项正确，且文件名未被改动。",
            file=sys.stderr,
        )
        return 1
    if plan.present_targets == 0:
        print(
            f"这些素材（{len(plan.missing_in_zip)} 个）在该 PPTX 中都不存在，"
            "很可能不是同一个文件。请用 --pptx 指定正确的原始 PPTX。",
            file=sys.stderr,
        )
        return 1

    if not dry_run:
        rewrite_pptx(source_pptx, plan.replacements, output_path)

    print_report(
        source_pptx=source_pptx,
        output_path=output_path,
        plan=plan,
        dry_run=dry_run,
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        return restore(
            folder_arg=args.folder,
            pptx_arg=args.pptx,
            output_arg=args.output,
            images_only=args.images_only,
            media_only=args.media_only,
            overwrite=args.overwrite,
            dry_run=args.dry_run,
        )
    except (FileNotFoundError, ValueError, zipfile.BadZipFile) as exc:
        print(f"还原失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
