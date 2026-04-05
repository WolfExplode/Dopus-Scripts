"""
Mark target files with a leading ✔ when a corresponding file exists under the source tree.

Same relative folder + peeled basename must match (e.g. source `foo.mkv` ↔ target `foo.mp4.jpg`).
Also: move all `.jpg` files from source → target preserving the same relative paths (removes them from source).
Also: read `Copy and Transfer.txt` in the source folder; each line is a `.mp4.jpg` / `.wmv.jpg` thumbnail name — strip `.jpg` and copy the corresponding video into the target tree (same relative path; does not remove from source).
Also: strip chosen characters from filenames under the target (stem only), trim spaces, remove trailing dots before the extension,
collapse a duplicated final extension (e.g. ``.mp4.mp4`` → ``.mp4``; case-insensitive),
and remove trailing copy suffixes like `` (1)`` / `` (23)`` (1–3 digits; avoids `` (2024)``-style years).
Also: under the target tree, append `` [immediate parent folder]`` before the file extension (files directly under the target root are skipped).

GUI: edit source/target folders and strip characters; settings are stored under %APPDATA%\\OrganizeFiles.
"""

from __future__ import annotations

import json
import os
import re
import shutil
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

CHECK = "\u2714"  # HEAVY CHECK MARK ✔

COPY_TRANSFER_LIST_NAME = "Copy and Transfer.txt"

# Trailing `` (n)`` / `` (n) (m)`` duplicate markers (Explorer-style); 1–3 digits so `` (2024)`` is kept.
_TITLE_STRIP_COPY_SUFFIX = re.compile(r"(?:\s+\(\d{1,3}\))+\Z")

# Optional space before ``[...]`` at end of stem (parent-folder tag normalization).
_TRAILING_BRACKET_TAG_END = re.compile(r"(?:\s?)\[([^\]]*)\]\Z")

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "OrganizeFiles"
CONFIG_PATH = CONFIG_DIR / "settings.json"


def config_load_defaults() -> tuple[str, str, str]:
    """Load saved source/target paths and strip-title characters, or empty strings if missing."""
    if not CONFIG_PATH.is_file():
        return "", "", ""
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        src = data.get("source") or ""
        tgt = data.get("target") or ""
        strip = data.get("strip_title_chars") or ""
        return str(src), str(tgt), str(strip)
    except (OSError, json.JSONDecodeError):
        return "", "", ""


def config_save(source: str, target: str, strip_title_chars: str = "") -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(
            {"source": source, "target": target, "strip_title_chars": strip_title_chars},
            indent=2,
            ensure_ascii=False,
        )
        + "\n",
        encoding="utf-8",
    )


def is_resolved_subpath(parent: Path, child: Path) -> bool:
    """True if child is equal to or inside parent (both resolved)."""
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def stem_variants(filename: str) -> set[str]:
    """Peel extensions from the right (.stem) until stable; collect each level for matching."""
    out: set[str] = set()
    cur = filename
    while True:
        stem = Path(cur).stem
        out.add(stem)
        if stem == cur:
            break
        cur = stem
    return out


def index_source(source_root: Path) -> dict[tuple[str, ...], set[str]]:
    """Map relative parent dir (as path parts) -> set of source file stems in that dir."""
    idx: dict[tuple[str, ...], set[str]] = defaultdict(set)
    for path in source_root.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(source_root)
        parent = rel.parent.parts
        idx[parent].add(path.stem)
    return idx


def target_parent_key(target_root: Path, path: Path) -> tuple[str, ...]:
    rel = path.relative_to(target_root)
    return rel.parent.parts


def has_source_match(
    source_index: dict[tuple[str, ...], set[str]],
    parent_key: tuple[str, ...],
    target_filename: str,
) -> bool:
    stems = source_index.get(parent_key)
    if not stems:
        return False
    return bool(stems & stem_variants(target_filename))


def is_jpg_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".jpg"


@dataclass
class JpgMoveScan:
    moves: list[tuple[Path, Path]]
    skipped_exists: list[tuple[Path, Path]]


@dataclass
class ScanResult:
    planned: list[tuple[Path, Path]]
    skipped_exists: list[Path]
    skipped_checked: int
    skipped_no_match: int
    skipped_under_source: int


@dataclass
class CopyTransferScan:
    copies: list[tuple[Path, Path]]
    skipped_exists: list[tuple[Path, Path]]
    missing: list[str]
    ambiguous: list[tuple[str, list[Path]]]
    list_missing: bool


@dataclass
class TitleStripScan:
    planned: list[tuple[Path, Path]]
    skipped_unchanged: int
    skipped_collision: list[tuple[Path, Path]]
    skipped_under_source: int


@dataclass
class ParentFolderTagScan:
    planned: list[tuple[Path, Path]]
    skipped_at_target_root: int
    skipped_already_tagged: int
    skipped_empty_folder_name: int
    skipped_collision: list[tuple[Path, Path]]
    skipped_under_source: int


def sanitize_folder_tag_segment(name: str) -> str:
    """Make folder name safe inside `` […] `` in a filename (Windows-forbidden chars)."""
    out = name
    for ch in '\\/:*?"<>|':
        out = out.replace(ch, "_")
    return out.strip().rstrip(".")


_ANY_BRACKET_TAG = re.compile(r"\s?\[([^\]]*)\]")


def remove_all_matching_bracket_tags(stem: str, tag_inner: str) -> str:
    """Remove every ``[tag]`` / `` [tag]`` in the stem whose inner text matches ``tag_inner`` (case-insensitive)."""
    want = tag_inner.casefold()
    to_remove: list[tuple[int, int]] = []
    for m in _ANY_BRACKET_TAG.finditer(stem):
        inner = m.group(1).strip().casefold()
        if inner == want:
            to_remove.append((m.start(), m.end()))
    if not to_remove:
        return stem
    parts: list[str] = []
    last = 0
    for start, end in to_remove:
        parts.append(stem[last:start])
        last = end
    parts.append(stem[last:])
    merged = "".join(parts)
    return re.sub(r" {2,}", " ", merged).strip()


def stem_final_trailing_bracket_inner(stem: str) -> Optional[str]:
    """Inner text of ``[...]`` at end of stem, or None."""
    m = _TRAILING_BRACKET_TAG_END.search(stem)
    return m.group(1).strip() if m else None


def parent_folder_tag_new_name(path: Path, target_root: Path) -> Optional[str]:
    """
    ``foo/file.mkv`` with parent ``foo`` → ``foo/file [foo].mkv``.
    Removes every ``[parent]`` tag that matches the folder (case-insensitive), anywhere in the stem,
    then appends a single normalized `` [parent]`` at the end (spacing and capitalization from folder).
    Files directly under ``target_root`` return None.
    If the stem ends with a different ``[...]`` tag after that cleanup, returns None (do not add another).
    """
    if not path.name:
        return None
    try:
        rel = path.relative_to(target_root)
    except ValueError:
        return None
    if not rel.parent.parts:
        return None

    tag_inner = sanitize_folder_tag_segment(path.parent.name)
    if not tag_inner:
        return None

    suffix = path.suffix
    stem = path.stem
    stem_clean = remove_all_matching_bracket_tags(stem, tag_inner)

    tail = stem_final_trailing_bracket_inner(stem_clean)
    if tail is not None and tail.casefold() != tag_inner.casefold():
        return None

    base = stem_clean.rstrip()
    if base:
        new_stem = base + f" [{tag_inner}]"
    else:
        new_stem = f"[{tag_inner}]"
    new_name = new_stem + suffix
    if new_name == path.name:
        return None
    return new_name


def scan_parent_folder_tag(source_root: Path, target_root: Path) -> ParentFolderTagScan:
    """Plan renames: normalize `` [immediate parent folder name]`` before extension. Skips target-root files."""
    planned: list[tuple[Path, Path]] = []
    skipped_collision: list[tuple[Path, Path]] = []
    skipped_at_target_root = 0
    skipped_already_tagged = 0
    skipped_empty_folder_name = 0
    skipped_under_source = 0

    for path in sorted(target_root.rglob("*")):
        if not path.is_file():
            continue
        resolved = path.resolve()
        if not is_resolved_subpath(target_root, resolved):
            continue
        if is_resolved_subpath(source_root, resolved):
            skipped_under_source += 1
            continue

        try:
            rel = path.relative_to(target_root)
        except ValueError:
            continue
        if not rel.parent.parts:
            skipped_at_target_root += 1
            continue

        tag_inner = sanitize_folder_tag_segment(path.parent.name)
        if not tag_inner:
            skipped_empty_folder_name += 1
            continue

        new_name = parent_folder_tag_new_name(path, target_root)
        if new_name is None:
            skipped_already_tagged += 1
            continue

        new_path = path.with_name(new_name)
        resolved_new = new_path.resolve()
        if not is_resolved_subpath(target_root, resolved_new):
            continue
        if is_resolved_subpath(source_root, resolved_new):
            skipped_under_source += 1
            continue
        if new_path.exists():
            skipped_collision.append((path, new_path))
            continue
        planned.append((path, new_path))

    return ParentFolderTagScan(
        planned=planned,
        skipped_at_target_root=skipped_at_target_root,
        skipped_already_tagged=skipped_already_tagged,
        skipped_empty_folder_name=skipped_empty_folder_name,
        skipped_collision=skipped_collision,
        skipped_under_source=skipped_under_source,
    )


def normalize_duplicate_trailing_extensions(basename: str) -> str:
    """
    Collapse repeated final extension: ``foo.mp4.mp4`` → ``foo.mp4`` (any depth; extension match is case-insensitive).
    """
    if not basename or basename.endswith(("/", "\\")):
        return basename
    p = Path(basename)
    name = basename
    while p.suffix and p.stem:
        if Path(p.stem).suffix.casefold() != p.suffix.casefold():
            break
        name = p.stem
        p = Path(name)
    return name


def transform_title_filename(filename: str, strip_chars: str) -> Optional[str]:
    """
    Remove each character in strip_chars from the stem; remove trailing `` (n)`` copy suffixes
    (1–3 digit n); trim spaces; strip trailing dots before the extension.
    Collapse a duplicated final extension (e.g. ``.mp4.mp4`` → ``.mp4``) before other stem rules.
    Returns new full filename if changed, else None. None if stem becomes empty.
    """
    if not filename or filename.endswith(("/", "\\")):
        return None
    original = Path(filename).name
    collapsed = normalize_duplicate_trailing_extensions(original)
    p = Path(collapsed)
    suffix = p.suffix
    stem = p.stem
    if not stem and not suffix:
        return None
    new_stem = stem
    for ch in strip_chars:
        if ch:
            new_stem = new_stem.replace(ch, "")
    new_stem = new_stem.rstrip().rstrip(".")
    new_stem = _TITLE_STRIP_COPY_SUFFIX.sub("", new_stem)
    new_stem = new_stem.rstrip().rstrip(".")
    if not new_stem:
        return None
    new_name = new_stem + suffix
    if new_name == original:
        return None
    return new_name


def scan_title_strip(source_root: Path, target_root: Path, strip_chars: str) -> TitleStripScan:
    """Plan renames under target_root only (stem cleanup). Skips paths inside source_root."""
    planned: list[tuple[Path, Path]] = []
    skipped_collision: list[tuple[Path, Path]] = []
    skipped_unchanged = 0
    skipped_under_source = 0

    for path in sorted(target_root.rglob("*")):
        if not path.is_file():
            continue
        resolved = path.resolve()
        if not is_resolved_subpath(target_root, resolved):
            continue
        if is_resolved_subpath(source_root, resolved):
            skipped_under_source += 1
            continue
        new_name = transform_title_filename(path.name, strip_chars)
        if new_name is None:
            skipped_unchanged += 1
            continue
        new_path = path.with_name(new_name)
        resolved_new = new_path.resolve()
        if not is_resolved_subpath(target_root, resolved_new):
            continue
        if is_resolved_subpath(source_root, resolved_new):
            skipped_under_source += 1
            continue
        if new_path.exists():
            skipped_collision.append((path, new_path))
            continue
        planned.append((path, new_path))

    return TitleStripScan(
        planned=planned,
        skipped_unchanged=skipped_unchanged,
        skipped_collision=skipped_collision,
        skipped_under_source=skipped_under_source,
    )


def validate_roots(source_root: Path, target_root: Path) -> Optional[str]:
    if not str(source_root).strip():
        return "Source folder is not set."
    if not str(target_root).strip():
        return "Target folder is not set."
    if source_root == target_root:
        return "Source and target directories must not be the same path."
    if is_resolved_subpath(source_root, target_root):
        return (
            "Target folder is inside the source folder; "
            "that would rename files under the source tree. Choose different paths."
        )
    if is_resolved_subpath(target_root, source_root):
        return (
            "Source folder is inside the target folder; "
            "copy/move operations could write into or through the source tree. Choose different paths."
        )
    if not source_root.is_dir():
        return f"Source is not a directory:\n{source_root}"
    if not target_root.is_dir():
        return f"Target is not a directory:\n{target_root}"
    return None


def scan_target(source_root: Path, target_root: Path) -> ScanResult:
    source_index = index_source(source_root)
    skipped_checked = 0
    skipped_no_match = 0
    skipped_under_source = 0
    planned: list[tuple[Path, Path]] = []
    skipped_exists: list[Path] = []

    for path in sorted(target_root.rglob("*")):
        if not path.is_file():
            continue
        resolved = path.resolve()
        if not is_resolved_subpath(target_root, resolved):
            continue
        if is_resolved_subpath(source_root, resolved):
            skipped_under_source += 1
            continue
        name = path.name
        if name.startswith(CHECK):
            skipped_checked += 1
            continue
        parent_key = target_parent_key(target_root, path)
        if not has_source_match(source_index, parent_key, name):
            skipped_no_match += 1
            continue
        new_path = path.with_name(f"{CHECK}{name}")
        resolved_new = new_path.resolve()
        if not is_resolved_subpath(target_root, resolved_new):
            continue
        if is_resolved_subpath(source_root, resolved_new):
            skipped_under_source += 1
            continue
        if new_path.exists():
            skipped_exists.append(new_path)
            continue
        planned.append((path, new_path))

    return ScanResult(
        planned=planned,
        skipped_exists=skipped_exists,
        skipped_checked=skipped_checked,
        skipped_no_match=skipped_no_match,
        skipped_under_source=skipped_under_source,
    )


def apply_renames(planned: list[tuple[Path, Path]]) -> tuple[int, list[str]]:
    """Returns (success_count, error_messages)."""
    errors: list[str] = []
    n = 0
    for old_path, new_path in planned:
        try:
            old_path.rename(new_path)
            n += 1
        except OSError as e:
            errors.append(f"{old_path}\n  {e}")
    return n, errors


def scan_jpg_moves(source_root: Path, target_root: Path) -> JpgMoveScan:
    """
    For every .jpg under source_root, plan a move to target_root / relative_path.
    Skips destinations that already exist (does not overwrite).
    """
    moves: list[tuple[Path, Path]] = []
    skipped_exists: list[tuple[Path, Path]] = []

    for path in sorted(source_root.rglob("*")):
        if not is_jpg_file(path):
            continue
        resolved = path.resolve()
        if not is_resolved_subpath(source_root, resolved):
            continue
        rel = path.relative_to(source_root)
        dest = (target_root / rel).resolve()
        if not is_resolved_subpath(target_root, dest):
            continue
        if dest.exists():
            skipped_exists.append((path, dest))
            continue
        moves.append((path, dest))

    return JpgMoveScan(moves=moves, skipped_exists=skipped_exists)


def apply_jpg_moves(moves: list[tuple[Path, Path]]) -> tuple[int, list[str]]:
    """Returns (success_count, error_messages). Removes each file from source after move."""
    errors: list[str] = []
    n = 0
    for src, dst in moves:
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(os.fspath(src), os.fspath(dst))
            n += 1
        except OSError as e:
            errors.append(f"{src}\n  -> {dst}\n  {e}")
    return n, errors


def video_basename_from_transfer_line(line: str) -> str:
    """Thumbnail list lines end with `.jpg` (e.g. `clip.mp4.jpg` → `clip.mp4`). Lines without `.jpg` are used as-is."""
    s = line.strip()
    if not s:
        return ""
    if s.lower().endswith(".jpg"):
        return s[:-4]
    return s


def index_basenames_under(root: Path) -> dict[str, list[Path]]:
    idx: dict[str, list[Path]] = defaultdict(list)
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if path.name == COPY_TRANSFER_LIST_NAME:
            continue
        idx[os.path.normcase(path.name)].append(path)
    return idx


def scan_copy_transfer(source_root: Path, target_root: Path) -> CopyTransferScan:
    """
    Read COPY_TRANSFER_LIST_NAME under source_root. Each line names a `.jpg` thumbnail;
    the actual video is the same name without the trailing `.jpg`. Plan copies into target_root
    preserving relative paths. Skips lines with no match, multiple matches, or existing destination.
    """
    list_path = source_root / COPY_TRANSFER_LIST_NAME
    if not list_path.is_file():
        return CopyTransferScan(
            copies=[],
            skipped_exists=[],
            missing=[],
            ambiguous=[],
            list_missing=True,
        )

    wanted: list[str] = []
    seen_line: set[str] = set()
    for line in list_path.read_text(encoding="utf-8", errors="replace").splitlines():
        name = video_basename_from_transfer_line(line)
        if not name or name in seen_line:
            continue
        seen_line.add(name)
        wanted.append(name)

    by_name = index_basenames_under(source_root)
    copies: list[tuple[Path, Path]] = []
    skipped_exists: list[tuple[Path, Path]] = []
    missing: list[str] = []
    ambiguous: list[tuple[str, list[Path]]] = []

    for base in wanted:
        matches = by_name.get(os.path.normcase(base), [])
        if not matches:
            missing.append(base)
            continue
        if len(matches) > 1:
            ambiguous.append((base, matches))
            continue
        src = matches[0]
        rel = src.relative_to(source_root)
        dst = (target_root / rel).resolve()
        if not is_resolved_subpath(target_root, dst):
            continue
        if dst.exists():
            skipped_exists.append((src, dst))
            continue
        copies.append((src, dst))

    return CopyTransferScan(
        copies=copies,
        skipped_exists=skipped_exists,
        missing=missing,
        ambiguous=ambiguous,
        list_missing=False,
    )


def apply_copy_transfer(copies: list[tuple[Path, Path]]) -> tuple[int, list[str]]:
    errors: list[str] = []
    n = 0
    for src, dst in copies:
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(os.fspath(src), os.fspath(dst))
            n += 1
        except OSError as e:
            errors.append(f"{src}\n  -> {dst}\n  {e}")
    return n, errors


def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class App(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title("Organize files")
            self.geometry("900x720")
            self.minsize(640, 480)

            s_default, t_default, strip_default = config_load_defaults()
            self.var_source = tk.StringVar(value=s_default)
            self.var_target = tk.StringVar(value=t_default)
            self.var_strip_title_chars = tk.StringVar(value=strip_default)

            frm = ttk.Frame(self, padding=8)
            frm.pack(fill=tk.BOTH, expand=True)

            ttk.Label(frm, text="Source:").grid(row=0, column=0, sticky="w")
            ttk.Entry(frm, textvariable=self.var_source).grid(
                row=1, column=0, sticky="ew", pady=(0, 6)
            )
            ttk.Button(frm, text="Browse…", command=self.browse_source).grid(
                row=1, column=1, sticky="e", padx=(6, 0), pady=(0, 6)
            )

            ttk.Label(frm, text="Target:").grid(row=2, column=0, sticky="w")
            ttk.Entry(frm, textvariable=self.var_target).grid(
                row=3, column=0, sticky="ew", pady=(0, 6)
            )
            ttk.Button(frm, text="Browse…", command=self.browse_target).grid(
                row=3, column=1, sticky="e", padx=(6, 0), pady=(0, 6)
            )

            ttk.Label(
                frm,
                text=("Rename target files: prepend ✔ to files in target based on the base name"),
                wraplength=820,
                justify=tk.LEFT,
            ).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 4))

            btn_row = ttk.Frame(frm)
            btn_row.grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 6))
            ttk.Button(btn_row, text="Scan / preview", command=self.on_preview).pack(
                side=tk.LEFT, padx=(0, 8)
            )
            ttk.Button(btn_row, text="Apply renames…", command=self.on_apply).pack(side=tk.LEFT)

            ttk.Separator(frm, orient=tk.HORIZONTAL).grid(
                row=6, column=0, columnspan=2, sticky="ew", pady=8
            )
            ttk.Label(
                frm,
                text=(
                    "Strip from titles (Target only): remove each character you type below from the "
                    "filename stem; trim spaces; remove trailing dots before the extension; collapse a "
                    "duplicated final extension (e.g. video.mp4.mp4 → video.mp4); and remove "
                    "trailing copy suffixes like \" (1)\" or \" (12)\" (up to 3 digits in parens, "
                    "so years like \" (2024)\" are not removed). "
                    "Example: 「Juno Bike Exercise」..mp4 with 「」 stripped → Juno Bike Exercise.mp4; "
                    "vdo-333-0123 (1).mp4 → vdo-333-0123.mp4."
                ),
                wraplength=820,
                justify=tk.LEFT,
            ).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(0, 2))
            ttk.Label(
                frm,
                text="Characters to strip from stem (optional; duplicate extension + copy-number suffix always checked):",
            ).grid(row=8, column=0, sticky="w", pady=(0, 2))
            title_strip_row = ttk.Frame(frm)
            title_strip_row.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(0, 6))
            title_strip_row.columnconfigure(0, weight=1)
            ttk.Entry(title_strip_row, textvariable=self.var_strip_title_chars).grid(
                row=0, column=0, sticky="ew", padx=(0, 8)
            )
            ttk.Button(title_strip_row, text="Preview title strip", command=self.on_preview_title_strip).grid(
                row=0, column=1, sticky="e", padx=(0, 8)
            )
            ttk.Button(title_strip_row, text="Apply title strip…", command=self.on_apply_title_strip).grid(
                row=0, column=2, sticky="e"
            )

            ttk.Separator(frm, orient=tk.HORIZONTAL).grid(
                row=10, column=0, columnspan=2, sticky="ew", pady=8
            )
            ttk.Label(
                frm,
                text=(
                    "Parent folder tag (Target only): for files inside a subfolder of the target, ensure "
                    "one trailing \" [parent]\" before the extension (immediate folder only). "
                    "Fixes no space before \"[\", wrong capitalization vs the folder, duplicate tags, and any "
                    "same-name \"[parent]\" earlier in the title — those are removed and one tag is reapplied at the end "
                    "(e.g. Ultrasound[NQ].mp4 → Ultrasound [NQ].mp4). "
                    "Files in the target root are skipped; a different trailing \"[tag]\" is left unchanged."
                ),
                wraplength=820,
                justify=tk.LEFT,
            ).grid(row=11, column=0, columnspan=2, sticky="ew", pady=(0, 2))
            parent_tag_row = ttk.Frame(frm)
            parent_tag_row.grid(row=12, column=0, columnspan=2, sticky="w", pady=(0, 6))
            ttk.Button(parent_tag_row, text="Preview parent folder tag", command=self.on_preview_parent_tag).pack(
                side=tk.LEFT, padx=(0, 8)
            )
            ttk.Button(parent_tag_row, text="Apply parent folder tag…", command=self.on_apply_parent_tag).pack(
                side=tk.LEFT
            )

            ttk.Separator(frm, orient=tk.HORIZONTAL).grid(
                row=13, column=0, columnspan=2, sticky="ew", pady=8
            )
            ttk.Label(
                frm,
                text=(
                    "Move .jpg: move all .jpg files from Source to Target using the same relative "
                    "paths (folders are created as needed; files leave Source):"
                ),
                wraplength=820,
                justify=tk.LEFT,
            ).grid(row=14, column=0, columnspan=2, sticky="ew")
            jpg_row = ttk.Frame(frm)
            jpg_row.grid(row=15, column=0, columnspan=2, sticky="w", pady=(4, 6))
            ttk.Button(jpg_row, text="Preview JPG moves", command=self.on_preview_jpg).pack(
                side=tk.LEFT, padx=(0, 8)
            )
            ttk.Button(jpg_row, text="Move JPGs…", command=self.on_apply_jpg).pack(side=tk.LEFT)

            ttk.Separator(frm, orient=tk.HORIZONTAL).grid(
                row=16, column=0, columnspan=2, sticky="ew", pady=8
            )
            ttk.Label(
                frm,
                text=(
                    f'Place "{COPY_TRANSFER_LIST_NAME}" in Source directory to copy videos from source to target. '
                    "each line is a thumbnail name ending in .jpg"
                ),
                wraplength=820,
                justify=tk.LEFT,
            ).grid(row=17, column=0, columnspan=2, sticky="ew")
            xfer_row = ttk.Frame(frm)
            xfer_row.grid(row=18, column=0, columnspan=2, sticky="w", pady=(4, 6))
            ttk.Button(
                xfer_row, text="Preview copy from list", command=self.on_preview_copy_transfer
            ).pack(side=tk.LEFT, padx=(0, 8))
            ttk.Button(
                xfer_row, text="Copy videos…", command=self.on_apply_copy_transfer
            ).pack(side=tk.LEFT)

            ttk.Label(frm, text="Preview:").grid(row=19, column=0, columnspan=2, sticky="w")
            self.text = tk.Text(frm, height=16, wrap=tk.NONE, font=("Consolas", 9))
            self.text.grid(row=20, column=0, columnspan=2, sticky="nsew")
            sy = ttk.Scrollbar(frm, orient=tk.VERTICAL, command=self.text.yview)
            sy.grid(row=20, column=2, sticky="ns")
            self.text.configure(yscrollcommand=sy.set)
            sx = ttk.Scrollbar(frm, orient=tk.HORIZONTAL, command=self.text.xview)
            sx.grid(row=21, column=0, columnspan=2, sticky="ew")
            self.text.configure(xscrollcommand=sx.set)

            frm.columnconfigure(0, weight=1)
            frm.rowconfigure(20, weight=1)

            self.protocol("WM_DELETE_WINDOW", self.on_close)

        def browse_source(self) -> None:
            p = filedialog.askdirectory(title="Source folder")
            if p:
                self.var_source.set(p)

        def browse_target(self) -> None:
            p = filedialog.askdirectory(title="Target folder")
            if p:
                self.var_target.set(p)

        def persist_paths(self) -> None:
            config_save(
                self.var_source.get().strip(),
                self.var_target.get().strip(),
                self.var_strip_title_chars.get(),
            )

        def on_close(self) -> None:
            try:
                self.persist_paths()
            except OSError:
                pass
            self.destroy()

        def on_preview(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, "Scanning…\n")
            self.update_idletasks()

            try:
                result = scan_target(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                self.text.delete("1.0", tk.END)
                return

            self.text.delete("1.0", tk.END)
            lines: list[str] = []
            lines.append("Files to rename (add ✔ prefix):\n")
            if result.planned:
                for old_path, new_path in result.planned:
                    lines.append(f"  {old_path}\n")
                    lines.append(f"    -> {new_path}\n")
            else:
                lines.append("  (none)\n")

            if result.skipped_exists:
                lines.append("\nNot included (✔ name already exists):\n")
                for p in result.skipped_exists:
                    lines.append(f"  {p}\n")

            lines.append(
                f"\nSummary — to rename: {len(result.planned)}, "
                f"already marked: {result.skipped_checked}, "
                f"no source match: {result.skipped_no_match}, "
                f"skipped (under source tree): {result.skipped_under_source}\n"
            )
            self.text.insert(tk.END, "".join(lines))

        def on_preview_title_strip(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            strip_chars = self.var_strip_title_chars.get()
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, "Scanning target for title strip…\n")
            self.update_idletasks()

            try:
                ts = scan_title_strip(src_r, tgt_r, strip_chars)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                self.text.delete("1.0", tk.END)
                return

            self.text.delete("1.0", tk.END)
            lines: list[str] = []
            lines.append("Target files to rename (strip characters from stem):\n")
            if ts.planned:
                for old_path, new_path in ts.planned:
                    lines.append(f"  {old_path}\n")
                    lines.append(f"    -> {new_path.name}\n")
            else:
                lines.append("  (none)\n")

            if ts.skipped_collision:
                lines.append("\nSkipped (destination name already exists):\n")
                for old_path, dest_path in ts.skipped_collision:
                    lines.append(f"  {old_path.name}\n")
                    lines.append(f"    (exists: {dest_path})\n")

            lines.append(
                f"\nSummary — to rename: {len(ts.planned)}, "
                f"unchanged: {ts.skipped_unchanged}, "
                f"collision: {len(ts.skipped_collision)}, "
                f"skipped (under source tree): {ts.skipped_under_source}\n"
            )
            self.text.insert(tk.END, "".join(lines))

        def on_apply_title_strip(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            strip_chars = self.var_strip_title_chars.get()
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                ts = scan_title_strip(src_r, tgt_r, strip_chars)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                return

            if not ts.planned:
                messagebox.showinfo(
                    "Nothing to do",
                    "No target files need renaming (no strip chars / duplicate extension / copy suffix changes apply). "
                    "Use Preview for details.",
                )
                return

            if not messagebox.askyesno(
                "Confirm title strip",
                f"Rename {len(ts.planned)} file(s) under the target folder?\n\n"
                "Only the filename changes (same folder): optional characters removed from the stem, "
                "duplicated final extensions collapsed (e.g. .mp4.mp4 → .mp4), "
                "trailing \" (n)\" copy suffixes removed (1–3 digits), trailing dots trimmed.\n\n"
                f"Skipped this run: {len(ts.skipped_collision)} collision(s) (name already exists).",
            ):
                return

            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            n, errors = apply_renames(ts.planned)
            if errors:
                messagebox.showwarning(
                    "Some renames failed",
                    f"Renamed: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                messagebox.showinfo("Done", f"Renamed {n} file(s) in the target.")

            self.on_preview_title_strip()

        def on_preview_parent_tag(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, "Scanning target for parent folder tags…\n")
            self.update_idletasks()

            try:
                pt = scan_parent_folder_tag(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                self.text.delete("1.0", tk.END)
                return

            self.text.delete("1.0", tk.END)
            lines: list[str] = []
            lines.append("Target files to rename (normalize [parent folder] before extension):\n")
            if pt.planned:
                for old_path, new_path in pt.planned:
                    lines.append(f"  {old_path}\n")
                    lines.append(f"    -> {new_path.name}\n")
            else:
                lines.append("  (none)\n")

            if pt.skipped_collision:
                lines.append("\nSkipped (destination name already exists):\n")
                for old_path, dest_path in pt.skipped_collision:
                    lines.append(f"  {old_path}\n")
                    lines.append(f"    (exists: {dest_path})\n")

            lines.append(
                f"\nSummary — to rename: {len(pt.planned)}, "
                f"at target root (skipped): {pt.skipped_at_target_root}, "
                f"already tagged / no change: {pt.skipped_already_tagged}, "
                f"empty folder name: {pt.skipped_empty_folder_name}, "
                f"collision: {len(pt.skipped_collision)}, "
                f"skipped (under source tree): {pt.skipped_under_source}\n"
            )
            self.text.insert(tk.END, "".join(lines))

        def on_apply_parent_tag(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                pt = scan_parent_folder_tag(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                return

            if not pt.planned:
                messagebox.showinfo(
                    "Nothing to do",
                    "No files need a parent folder tag. Use Preview for details (e.g. files only at target root).",
                )
                return

            if not messagebox.askyesno(
                "Confirm parent folder tag",
                f"Rename {len(pt.planned)} file(s) under the target folder?\n\n"
                "Each file in a subfolder gets a single normalized \" [parent]\" tag (spacing, capitalization, "
                "dedupe). Files directly under the target root are not included.\n\n"
                f"Skipped this run: {len(pt.skipped_collision)} collision(s) (name already exists).",
            ):
                return

            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            n, errors = apply_renames(pt.planned)
            if errors:
                messagebox.showwarning(
                    "Some renames failed",
                    f"Renamed: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                messagebox.showinfo("Done", f"Renamed {n} file(s) in the target.")

            self.on_preview_parent_tag()

        def on_preview_jpg(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, "Scanning .jpg files…\n")
            self.update_idletasks()

            try:
                jpg = scan_jpg_moves(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                self.text.delete("1.0", tk.END)
                return

            self.text.delete("1.0", tk.END)
            lines: list[str] = []
            lines.append("JPG files to move (source → target, relative path preserved):\n")
            if jpg.moves:
                for old_path, new_path in jpg.moves:
                    lines.append(f"  {old_path}\n")
                    lines.append(f"    -> {new_path}\n")
            else:
                lines.append("  (none)\n")

            if jpg.skipped_exists:
                lines.append("\nSkipped (file already exists at destination):\n")
                for s_path, d_path in jpg.skipped_exists:
                    lines.append(f"  {s_path}\n")
                    lines.append(f"    (exists: {d_path})\n")

            lines.append(
                f"\nSummary — to move: {len(jpg.moves)}, skipped (dest exists): {len(jpg.skipped_exists)}\n"
            )
            self.text.insert(tk.END, "".join(lines))

        def on_apply_jpg(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                jpg = scan_jpg_moves(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                return

            if not jpg.moves:
                messagebox.showinfo("Nothing to do", "No .jpg files to move.")
                return

            if not messagebox.askyesno(
                "Confirm move",
                f"Move {len(jpg.moves)} .jpg file(s) from source to target?\n\n"
                "Relative folders will be created under the target. "
                "Files will be removed from the source (moved, not copied).\n\n"
                "Skipped: destinations that already exist.",
            ):
                return

            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            n, errors = apply_jpg_moves(jpg.moves)
            if errors:
                messagebox.showwarning(
                    "Some moves failed",
                    f"Moved: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                messagebox.showinfo("Done", f"Moved {n} .jpg file(s) to target.")

            self.on_preview_jpg()

        def on_preview_copy_transfer(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            self.text.delete("1.0", tk.END)
            self.text.insert(tk.END, "Reading copy list…\n")
            self.update_idletasks()

            try:
                xfer = scan_copy_transfer(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                self.text.delete("1.0", tk.END)
                return

            self.text.delete("1.0", tk.END)
            lines: list[str] = []
            list_path = src_r / COPY_TRANSFER_LIST_NAME

            if xfer.list_missing:
                lines.append(
                    f"List file not found (expected at):\n  {list_path}\n"
                )
                self.text.insert(tk.END, "".join(lines))
                return

            lines.append(
                f'Videos to copy (from "{COPY_TRANSFER_LIST_NAME}"; source → target):\n'
            )
            if xfer.copies:
                for old_path, new_path in xfer.copies:
                    lines.append(f"  {old_path}\n")
                    lines.append(f"    -> {new_path}\n")
            else:
                lines.append("  (none)\n")

            if xfer.skipped_exists:
                lines.append("\nSkipped (file already exists at destination):\n")
                for s_path, d_path in xfer.skipped_exists:
                    lines.append(f"  {s_path}\n")
                    lines.append(f"    (exists: {d_path})\n")

            if xfer.missing:
                lines.append(f"\nNot found under source ({len(xfer.missing)}):\n")
                for name in xfer.missing[:80]:
                    lines.append(f"  {name}\n")
                if len(xfer.missing) > 80:
                    lines.append(f"  … and {len(xfer.missing) - 80} more\n")

            if xfer.ambiguous:
                lines.append("\nSkipped (multiple files with same name under source):\n")
                for base, paths in xfer.ambiguous[:30]:
                    lines.append(f"  {base}\n")
                    for p in paths[:5]:
                        lines.append(f"    {p}\n")
                    if len(paths) > 5:
                        lines.append(f"    … and {len(paths) - 5} more\n")

            lines.append(
                f"\nSummary — to copy: {len(xfer.copies)}, "
                f"dest exists: {len(xfer.skipped_exists)}, "
                f"not found: {len(xfer.missing)}, "
                f"ambiguous: {len(xfer.ambiguous)}\n"
            )
            self.text.insert(tk.END, "".join(lines))

        def on_apply_copy_transfer(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                xfer = scan_copy_transfer(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                return

            if xfer.list_missing:
                messagebox.showerror(
                    "List missing",
                    f'Could not find "{COPY_TRANSFER_LIST_NAME}" under the source folder:\n\n'
                    f"{src_r}",
                )
                return

            if not xfer.copies:
                messagebox.showinfo(
                    "Nothing to do",
                    "No videos to copy (all missing, ambiguous, or already at destination). "
                    "Use Preview for details.",
                )
                return

            if not messagebox.askyesno(
                "Confirm copy",
                f"Copy {len(xfer.copies)} video file(s) from source to target?\n\n"
                "Relative folders will be created under the target. "
                "Source files are not removed.\n\n"
                "Skipped: missing, ambiguous names, or destination already exists.",
            ):
                return

            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            n, errors = apply_copy_transfer(xfer.copies)
            if errors:
                messagebox.showwarning(
                    "Some copies failed",
                    f"Copied: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                messagebox.showinfo("Done", f"Copied {n} video file(s) to target.")

            self.on_preview_copy_transfer()

        def on_apply(self) -> None:
            src = Path(self.var_source.get().strip())
            tgt = Path(self.var_target.get().strip())
            err = validate_roots(src, tgt)
            if err:
                messagebox.showerror("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                result = scan_target(src_r, tgt_r)
            except OSError as e:
                messagebox.showerror("Scan failed", str(e))
                return

            if not result.planned:
                messagebox.showinfo("Nothing to do", "No renames pending.")
                return

            if not messagebox.askyesno(
                "Confirm",
                f"Rename {len(result.planned)} file(s) in the target folder?\n\n"
                "This step only adds a ✔ prefix to names under the target; "
                "it does not move or delete files.",
            ):
                return

            try:
                self.persist_paths()
            except OSError as e:
                messagebox.showwarning("Could not save settings", str(e))

            n, errors = apply_renames(result.planned)
            if errors:
                messagebox.showwarning(
                    "Some renames failed",
                    f"Renamed: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                messagebox.showinfo("Done", f"Renamed {n} file(s).")

            self.on_preview()

    App().mainloop()


if __name__ == "__main__":
    run_gui()
