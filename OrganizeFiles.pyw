"""
Mark target files with a leading ✔ when a corresponding file exists under the source tree.

Same relative folder + peeled basename must match (e.g. source `foo.mkv` ↔ target `foo.mp4.jpg`).
Also: move all `.jpg` files from source → target preserving the same relative paths (removes them from source).
Also: read `Copy and Transfer.txt` in the source folder; each line is a `.mp4.jpg` / `.wmv.jpg` thumbnail name — strip `.jpg` and copy the corresponding video into the target tree (same relative path; does not remove from source).
Also: strip chosen characters from filenames under the target (stem only), trim spaces, remove trailing dots before the extension,
collapse a duplicated final extension (e.g. ``.mp4.mp4`` → ``.mp4``; case-insensitive),
and remove trailing copy suffixes like `` (1)`` / `` (23)`` (1–3 digits; avoids `` (2024)``-style years).
Also: under the target tree, append `` [immediate parent folder]`` before the file extension (files directly under the target root are skipped).

GUI (Dear PyGui): edit source/target folders and strip characters; settings are stored under %APPDATA%\\OrganizeFiles.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import unicodedata
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

# CJK / symbol UI fonts (first match under %WINDIR%\Fonts).
_UNICODE_UI_FONT_NAMES = (
    "NotoSansSC-VF.ttf",
    "msyh.ttc",
    "msyhbd.ttc",
    "simsun.ttc",
    "mingliu.ttc",
    "msjh.ttc",
    "segoeui.ttf",
)
def normalize_strip_chars(strip_chars: str) -> str:
    """NFC-normalize so pasted CJK / fullwidth brackets match filenames on disk."""
    return unicodedata.normalize("NFC", strip_chars or "")


def windows_fonts_dir() -> Path:
    return Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"


def pick_unicode_ui_font() -> Optional[Path]:
    """Best installed font for Chinese, Japanese, Korean, and Latin UI text."""
    fonts_dir = windows_fonts_dir()
    if not fonts_dir.is_dir():
        return None
    for name in _UNICODE_UI_FONT_NAMES:
        path = fonts_dir / name
        if path.is_file():
            return path
    return None


def config_load_defaults() -> tuple[str, str, str]:
    """Load saved source/target paths and strip-title characters, or empty strings if missing."""
    if not CONFIG_PATH.is_file():
        return "", "", ""
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        src = data.get("source") or ""
        tgt = data.get("target") or ""
        strip = normalize_strip_chars(data.get("strip_title_chars") or "")
        return str(src), str(tgt), strip
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


def scan_parent_folder_tag(
    source_root: Path, target_root: Path, only: Optional[set[Path]] = None
) -> ParentFolderTagScan:
    """Plan renames: normalize `` [immediate parent folder name]`` before extension. Skips target-root files."""
    planned: list[tuple[Path, Path]] = []
    skipped_collision: list[tuple[Path, Path]] = []
    skipped_at_target_root = 0
    skipped_already_tagged = 0
    skipped_empty_folder_name = 0
    skipped_under_source = 0

    for path in iter_files_in_tree(target_root, only):
        resolved = path.resolve()
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
    strip_chars = normalize_strip_chars(strip_chars)
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


def scan_title_strip(
    source_root: Path,
    target_root: Path,
    strip_chars: str,
    only: Optional[set[Path]] = None,
) -> TitleStripScan:
    """Plan renames under target_root only (stem cleanup). Skips paths inside source_root."""
    planned: list[tuple[Path, Path]] = []
    skipped_collision: list[tuple[Path, Path]] = []
    skipped_unchanged = 0
    skipped_under_source = 0

    for path in iter_files_in_tree(target_root, only):
        resolved = path.resolve()
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


def read_only_paths(
    list_path: Optional[str] = None, file_args: Optional[list[str]] = None
) -> Optional[set[Path]]:
    """Resolved file paths to limit scans; None means process the full tree."""
    out: set[Path] = set()
    for s in file_args or []:
        p = Path(s.strip())
        if p.is_file():
            out.add(p.resolve())
    if list_path:
        lp = Path(list_path)
        if lp.is_file():
            for line in lp.read_text(encoding="utf-8", errors="replace").splitlines():
                s = line.strip()
                if not s:
                    continue
                p = Path(s)
                if p.is_file():
                    out.add(p.resolve())
    return out if out else None


def iter_files_in_tree(root: Path, only: Optional[set[Path]] = None):
    """Yield files under root, or only the resolved paths in only when set."""
    if only is not None:
        for path in sorted(only):
            if not path.is_file():
                continue
            resolved = path.resolve()
            if is_resolved_subpath(root, resolved):
                yield path
        return
    for path in sorted(root.rglob("*")):
        if path.is_file():
            yield path


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


def scan_target(
    source_root: Path, target_root: Path, only: Optional[set[Path]] = None
) -> ScanResult:
    source_index = index_source(source_root)
    skipped_checked = 0
    skipped_no_match = 0
    skipped_under_source = 0
    planned: list[tuple[Path, Path]] = []
    skipped_exists: list[Path] = []

    for path in iter_files_in_tree(target_root, only):
        resolved = path.resolve()
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


def scan_jpg_moves(
    source_root: Path, target_root: Path, only: Optional[set[Path]] = None
) -> JpgMoveScan:
    """
    For every .jpg under source_root, plan a move to target_root / relative_path.
    Skips destinations that already exist (does not overwrite).
    """
    moves: list[tuple[Path, Path]] = []
    skipped_exists: list[tuple[Path, Path]] = []

    for path in iter_files_in_tree(source_root, only):
        if not is_jpg_file(path):
            continue
        resolved = path.resolve()
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


def scan_copy_transfer(
    source_root: Path, target_root: Path, only: Optional[set[Path]] = None
) -> CopyTransferScan:
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
        if only is not None:
            matches = [m for m in matches if m.resolve() in only]
            if not matches:
                continue
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


def run_gui(
    initial_source: Optional[str] = None,
    initial_target: Optional[str] = None,
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    import dearpygui.dearpygui as dpg

    class App:
        TAG_SOURCE = "source_input"
        TAG_TARGET = "target_input"
        TAG_STRIP = "strip_chars_input"
        TAG_ONLY = "only_selected_chk"
        TAG_PREVIEW = "preview_text"
        TAG_DLG_SOURCE = "dlg_source"
        TAG_DLG_TARGET = "dlg_target"

        TIP_FOLDERS = (
            "Source and target roots for every operation.\n"
            "Paths are saved when you close the app (%APPDATA%\\OrganizeFiles)."
        )
        TIP_SOURCE = "Folder tree to match against and to pull .jpg / copy-list videos from."
        TIP_TARGET = "Folder tree where renames, tags, and incoming files are applied."
        TIP_ONLY = (
            "When launched from Directory Opus with a file list, process only those paths "
            "(not whole folder trees). Disabled if no list was passed."
        )
        TIP_MARK = (
            "Target only: prepend ✔ to a filename when the same relative folder contains a source file "
            "with a matching peeled basename (e.g. source foo.mkv ↔ target foo.mp4.jpg).\n"
            "Does not move or delete files."
        )
        TIP_TITLE = (
            "Target only: clean filenames — optional characters removed from the stem, spaces trimmed, "
            "trailing dots before the extension removed, duplicated final extension collapsed "
            "(e.g. video.mp4.mp4 → video.mp4), and Explorer-style copy suffixes removed "
            '(" (1)" / " (12)", 1–3 digits; years like " (2024)" are kept).\n'
            "Example: 「Juno Bike Exercise」..mp4 with 「」 stripped → Juno Bike Exercise.mp4."
        )
        TIP_STRIP_CHARS = (
            "Characters removed from each stem (optional). Paste any Unicode here — e.g. corner "
            "brackets 「」 (U+300C / U+300D), fullwidth quotes, or emoji. Each character is removed "
            "wherever it appears in the stem. Duplicate-extension and copy-number suffix cleanup "
            "always runs even when this is empty."
        )
        TIP_PARENT = (
            'Target only: for files in subfolders, ensure one trailing " [parent]" before the extension '
            "(immediate parent folder only). Fixes missing space before [, wrong capitalization, "
            "duplicate tags, and stray [parent] earlier in the name — then reapplies one tag at the end "
            "(e.g. Ultrasound[NQ].mp4 → Ultrasound [NQ].mp4).\n"
            "Files at the target root are skipped; a different trailing [tag] is left as-is."
        )
        TIP_JPG = (
            "Move every .jpg under Source into Target using the same relative paths "
            "(creates folders as needed; files are removed from Source)."
        )
        TIP_COPY = (
            f'Put "{COPY_TRANSFER_LIST_NAME}" in the Source folder — one thumbnail filename per line '
            "(.mp4.jpg / .wmv.jpg). The .jpg is stripped to find the video under Source; "
            "videos are copied into Target (same relative path; Source is not deleted)."
        )
        TIP_PREVIEW_PANEL = "Output from the last Preview or Apply scan."

        def __init__(self) -> None:
            s_default, t_default, strip_default = config_load_defaults()
            if initial_source and str(initial_source).strip():
                s_default = str(initial_source).strip()
            if initial_target and str(initial_target).strip():
                t_default = str(initial_target).strip()
            self._only_paths = read_only_paths(initial_only_list, initial_only_files)
            self._only_default = bool(self._only_paths)
            self._theme_apply = "theme_btn_apply"
            self._theme_preview = "theme_btn_preview"

            dpg.create_context()
            self._build_themes()

            with dpg.file_dialog(
                directory_selector=True,
                show=False,
                modal=True,
                callback=self._on_dir_dialog,
                tag=self.TAG_DLG_SOURCE,
                width=700,
                height=400,
            ):
                pass
            with dpg.file_dialog(
                directory_selector=True,
                show=False,
                modal=True,
                callback=self._on_dir_dialog,
                tag=self.TAG_DLG_TARGET,
                width=700,
                height=400,
            ):
                pass

            with dpg.window(tag="primary_window", label="Organize Files", no_title_bar=True):
                self._build_main_layout(s_default, t_default, strip_default)
            self._build_fonts()

            dpg.create_viewport(
                title="Organize Files",
                width=1080,
                height=740,
                min_width=780,
                min_height=520,
            )
            dpg.setup_dearpygui()
            dpg.show_viewport()
            dpg.set_primary_window("primary_window", True)
            dpg.set_exit_callback(self.on_close)

        def _build_themes(self) -> None:
            accent = (72, 168, 190)
            accent_h = (92, 198, 220)
            apply_bg = (52, 128, 108)
            apply_h = (68, 158, 132)

            with dpg.theme(tag="app_theme"):
                with dpg.theme_component(dpg.mvAll):
                    dpg.add_theme_color(dpg.mvThemeCol_WindowBg, (16, 18, 24))
                    dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (22, 25, 32))
                    dpg.add_theme_color(dpg.mvThemeCol_PopupBg, (26, 30, 40))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (34, 38, 50))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, (44, 50, 66))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, (52, 60, 78))
                    dpg.add_theme_color(dpg.mvThemeCol_Button, (48, 54, 70))
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (60, 68, 88))
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, accent)
                    dpg.add_theme_color(dpg.mvThemeCol_Header, (38, 44, 58))
                    dpg.add_theme_color(dpg.mvThemeCol_HeaderHovered, (50, 58, 76))
                    dpg.add_theme_color(dpg.mvThemeCol_HeaderActive, (58, 72, 92))
                    dpg.add_theme_color(dpg.mvThemeCol_Text, (228, 232, 240))
                    dpg.add_theme_color(dpg.mvThemeCol_TextDisabled, (110, 118, 135))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (52, 58, 74))
                    dpg.add_theme_color(dpg.mvThemeCol_Separator, (48, 54, 70))
                    dpg.add_theme_color(dpg.mvThemeCol_CheckMark, accent_h)
                    dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg, (20, 22, 30))
                    dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrab, (56, 62, 80))
                    dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrabHovered, accent)
                    dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 10)
                    dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 8)
                    dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6)
                    dpg.add_theme_style(dpg.mvStyleVar_GrabRounding, 6)
                    dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 14, 14)
                    dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 10, 6)
                    dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 8)
                    dpg.add_theme_style(dpg.mvStyleVar_ScrollbarSize, 14)

            with dpg.theme(tag=self._theme_preview):
                with dpg.theme_component(dpg.mvButton):
                    dpg.add_theme_color(dpg.mvThemeCol_Button, (42, 48, 62))
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (56, 64, 82))
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (70, 80, 102))

            with dpg.theme(tag=self._theme_apply):
                with dpg.theme_component(dpg.mvButton):
                    dpg.add_theme_color(dpg.mvThemeCol_Button, apply_bg)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, apply_h)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (80, 178, 150))

            with dpg.theme(tag="theme_preview_panel"):
                with dpg.theme_component(dpg.mvChildWindow):
                    dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (12, 14, 20))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (40, 72, 88))

            with dpg.theme(tag="theme_actions_panel"):
                with dpg.theme_component(dpg.mvChildWindow):
                    dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (20, 23, 30))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (44, 50, 66))

            with dpg.theme(tag="theme_modal"):
                with dpg.theme_component(dpg.mvWindowAppItem):
                    dpg.add_theme_color(dpg.mvThemeCol_TitleBg, (32, 38, 52))
                    dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, (48, 88, 108))

            dpg.bind_theme("app_theme")

        def _make_app_font(self, size: int) -> Optional[int | str]:
            ui_font_path = pick_unicode_ui_font()
            if not ui_font_path:
                return None
            return dpg.add_font(str(ui_font_path), size)

        def _build_fonts(self) -> None:
            with dpg.font_registry():
                ui_font = self._make_app_font(14)
                if ui_font is not None:
                    dpg.bind_font(ui_font)
                    title_font = self._make_app_font(18)
                    if title_font is not None:
                        dpg.bind_item_font("title_main", title_font)
                    preview_font = self._make_app_font(13)
                    if preview_font is not None:
                        dpg.bind_item_font(self.TAG_PREVIEW, preview_font)
                    return
                segoe = windows_fonts_dir() / "segoeui.ttf"
                if segoe.is_file():
                    dpg.bind_font(dpg.add_font(str(segoe), 14))

        def _hover_tip(self, parent: int | str, text: str, delay: float = 0.4) -> None:
            with dpg.tooltip(parent, delay=delay):
                dpg.add_text(text, wrap=400, color=(200, 208, 220))

        def _preview_apply_row(
            self,
            preview_label: str,
            apply_label: str,
            preview_cb,
            apply_cb,
            tip_preview: str,
            tip_apply: str,
        ) -> None:
            with dpg.group(horizontal=True):
                btn_p = dpg.add_button(label=preview_label, callback=preview_cb, width=128)
                btn_a = dpg.add_button(label=apply_label, callback=apply_cb, width=128)
            dpg.bind_item_theme(btn_p, self._theme_preview)
            dpg.bind_item_theme(btn_a, self._theme_apply)
            self._hover_tip(btn_p, tip_preview)
            self._hover_tip(btn_a, tip_apply)

        def _path_field(
            self,
            label: str,
            input_tag: str,
            dialog_tag: str,
            default: str,
            tip: str,
        ) -> None:
            dpg.add_text(label, color=(150, 158, 175))
            with dpg.group(horizontal=True):
                dpg.add_input_text(tag=input_tag, default_value=default, width=-48)
                btn = dpg.add_button(
                    label="…",
                    width=40,
                    callback=lambda: dpg.show_item(dialog_tag),
                )
            self._hover_tip(input_tag, tip)
            self._hover_tip(btn, f"Browse for {label.lower()}")

        def _section(
            self,
            title: str,
            tip: str,
            default_open: bool,
        ):
            hdr = dpg.add_collapsing_header(
                label=title,
                default_open=default_open,
                tag=dpg.generate_uuid(),
            )
            self._hover_tip(hdr, tip)
            return hdr

        def _build_main_layout(
            self,
            s_default: str,
            t_default: str,
            strip_default: str,
        ) -> None:
            dpg.add_text("Organize Files", tag="title_main", color=(120, 200, 220))
            dpg.add_text(
                "Source ↔ target workflow",
                color=(130, 138, 155),
            )
            dpg.add_spacer(height=6)

            with dpg.group(horizontal=True):
                with dpg.child_window(width=400, height=-1, border=True, tag="panel_actions"):
                    dpg.bind_item_theme("panel_actions", "theme_actions_panel")

                    hdr_folders = self._section("Folders", self.TIP_FOLDERS, default_open=True)
                    with dpg.group(parent=hdr_folders):
                        self._path_field("Source", self.TAG_SOURCE, self.TAG_DLG_SOURCE, s_default, self.TIP_SOURCE)
                        dpg.add_spacer(height=4)
                        self._path_field("Target", self.TAG_TARGET, self.TAG_DLG_TARGET, t_default, self.TIP_TARGET)
                        dpg.add_spacer(height=4)
                        chk = dpg.add_checkbox(
                            label="Only listed files",
                            tag=self.TAG_ONLY,
                            default_value=self._only_default,
                            enabled=bool(self._only_paths),
                        )
                        self._hover_tip(chk, self.TIP_ONLY)

                    dpg.add_spacer(height=8)

                    hdr_mark = self._section(f"{CHECK}  Match marks", self.TIP_MARK, default_open=True)
                    with dpg.group(parent=hdr_mark):
                        self._preview_apply_row(
                            "Preview",
                            "Apply",
                            self.on_preview,
                            self.on_apply,
                            "Scan target tree and list files that would get a ✔ prefix.",
                            "Add ✔ prefix to matched target files (no moves/deletes).",
                        )

                    hdr_title = self._section("Title cleanup", self.TIP_TITLE, default_open=False)
                    with dpg.group(parent=hdr_title):
                        dpg.add_text("Strip characters", color=(150, 158, 175))
                        dpg.add_input_text(
                            tag=self.TAG_STRIP,
                            default_value=strip_default,
                            width=-1,
                            hint="e.g. 「」",
                        )
                        self._hover_tip(self.TAG_STRIP, self.TIP_STRIP_CHARS)
                        dpg.add_spacer(height=4)
                        self._preview_apply_row(
                            "Preview",
                            "Apply",
                            self.on_preview_title_strip,
                            self.on_apply_title_strip,
                            "List target renames from title cleanup rules.",
                            "Rename target files (stem cleanup only).",
                        )

                    hdr_parent = self._section("Parent folder tags", self.TIP_PARENT, default_open=False)
                    with dpg.group(parent=hdr_parent):
                        self._preview_apply_row(
                            "Preview",
                            "Apply",
                            self.on_preview_parent_tag,
                            self.on_apply_parent_tag,
                            'List files that would get a normalized " [parent]" tag.',
                            "Apply parent-folder tags on target subfolders.",
                        )

                    hdr_jpg = self._section("JPG transfer", self.TIP_JPG, default_open=False)
                    with dpg.group(parent=hdr_jpg):
                        self._preview_apply_row(
                            "Preview",
                            "Move",
                            self.on_preview_jpg,
                            self.on_apply_jpg,
                            "List .jpg files that would move source → target.",
                            "Move .jpg files (removed from source).",
                        )

                    hdr_copy = self._section("Copy from list", self.TIP_COPY, default_open=False)
                    with dpg.group(parent=hdr_copy):
                        self._preview_apply_row(
                            "Preview",
                            "Copy",
                            self.on_preview_copy_transfer,
                            self.on_apply_copy_transfer,
                            f"Read {COPY_TRANSFER_LIST_NAME} and list videos to copy.",
                            "Copy listed videos into target (source kept).",
                        )

                with dpg.child_window(width=-1, height=-1, border=True, tag="panel_preview"):
                    dpg.bind_item_theme("panel_preview", "theme_preview_panel")
                    with dpg.group(horizontal=True):
                        dpg.add_text("Preview", color=(120, 200, 220))
                        dpg.add_text("  ·  ", color=(80, 88, 100))
                        hint = dpg.add_text("hover controls for help", color=(100, 108, 125))
                        self._hover_tip(hint, self.TIP_PREVIEW_PANEL)
                    dpg.add_spacer(height=4)
                    dpg.add_input_text(
                        tag=self.TAG_PREVIEW,
                        multiline=True,
                        readonly=True,
                        width=-1,
                        height=-1,
                        tab_input=False,
                        default_value=(
                            "Pick a section on the left, then Preview.\n\n"
                            "Results show up here."
                        ),
                    )

        def _on_dir_dialog(self, _sender, app_data) -> None:
            path = app_data.get("file_path_name") or app_data.get("current_path") or ""
            if not path:
                return
            if _sender == self.TAG_DLG_SOURCE:
                dpg.set_value(self.TAG_SOURCE, path)
            elif _sender == self.TAG_DLG_TARGET:
                dpg.set_value(self.TAG_TARGET, path)

        def _set_preview(self, text: str) -> None:
            dpg.set_value(self.TAG_PREVIEW, text)
            if dpg.is_dearpygui_running():
                dpg.render_dearpygui_frame()

        def _msg(self, title: str, message: str) -> None:
            modal_tag = dpg.generate_uuid()

            def close() -> None:
                dpg.delete_item(modal_tag)

            with dpg.window(label=title, modal=True, popup=True, tag=modal_tag, no_close=True):
                dpg.bind_item_theme(modal_tag, "theme_modal")
                dpg.add_text(message, wrap=500)
                dpg.add_button(label="OK", callback=close, width=100)
            while dpg.does_item_exist(modal_tag) and dpg.is_dearpygui_running():
                dpg.render_dearpygui_frame()

        def _msg_error(self, title: str, message: str) -> None:
            self._msg(title, message)

        def _msg_warning(self, title: str, message: str) -> None:
            self._msg(title, message)

        def _msg_info(self, title: str, message: str) -> None:
            self._msg(title, message)

        def _ask_yesno(self, title: str, message: str) -> bool:
            modal_tag = dpg.generate_uuid()
            result: list[Optional[bool]] = [None]

            def close(value: bool) -> None:
                result[0] = value
                dpg.delete_item(modal_tag)

            with dpg.window(label=title, modal=True, popup=True, tag=modal_tag, no_close=True):
                dpg.bind_item_theme(modal_tag, "theme_modal")
                dpg.add_text(message, wrap=500)
                with dpg.group(horizontal=True):
                    yes_btn = dpg.add_button(label="Yes", callback=lambda: close(True), width=88)
                    no_btn = dpg.add_button(label="No", callback=lambda: close(False), width=88)
                dpg.bind_item_theme(yes_btn, self._theme_apply)
                dpg.bind_item_theme(no_btn, self._theme_preview)
            while result[0] is None and dpg.is_dearpygui_running():
                dpg.render_dearpygui_frame()
            return bool(result[0])

        def get_only_paths(self) -> Optional[set[Path]]:
            if dpg.get_value(self.TAG_ONLY) and self._only_paths:
                return self._only_paths
            return None

        def persist_paths(self) -> None:
            config_save(
                str(dpg.get_value(self.TAG_SOURCE)).strip(),
                str(dpg.get_value(self.TAG_TARGET)).strip(),
                str(dpg.get_value(self.TAG_STRIP)),
            )

        def on_close(self) -> None:
            try:
                self.persist_paths()
            except OSError:
                pass

        def run(self) -> None:
            dpg.start_dearpygui()
            dpg.destroy_context()

        def on_preview(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                self._msg_warning("Could not save settings", str(e))

            self._set_preview("Scanning…\n")

            try:
                result = scan_target(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                self._set_preview("")
                return

            self._set_preview("")
            lines: list[str] = []
            only = self.get_only_paths()
            if only is not None:
                lines.append(f"Limited to {len(only)} listed file(s).\n\n")
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
            self._set_preview("".join(lines))

        def on_preview_title_strip(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            strip_chars = str(dpg.get_value(self.TAG_STRIP))
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                self._msg_warning("Could not save settings", str(e))

            self._set_preview("Scanning target for title strip…\n")

            try:
                ts = scan_title_strip(src_r, tgt_r, strip_chars, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                self._set_preview("")
                return

            self._set_preview("")
            lines: list[str] = []
            only = self.get_only_paths()
            if only is not None:
                lines.append(f"Limited to {len(only)} listed file(s).\n\n")
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
            self._set_preview("".join(lines))

        def on_apply_title_strip(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            strip_chars = str(dpg.get_value(self.TAG_STRIP))
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                ts = scan_title_strip(src_r, tgt_r, strip_chars, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                return

            if not ts.planned:
                self._msg_info(
                    "Nothing to do",
                    "No target files need renaming (no strip chars / duplicate extension / copy suffix changes apply). "
                    "Use Preview for details.",
                )
                return

            if not self._ask_yesno(
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
                self._msg_warning("Could not save settings", str(e))

            n, errors = apply_renames(ts.planned)
            if errors:
                self._msg_warning(
                    "Some renames failed",
                    f"Renamed: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                self._msg_info("Done", f"Renamed {n} file(s) in the target.")

            self.on_preview_title_strip()

        def on_preview_parent_tag(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                self._msg_warning("Could not save settings", str(e))

            self._set_preview("Scanning target for parent folder tags…\n")

            try:
                pt = scan_parent_folder_tag(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                self._set_preview("")
                return

            self._set_preview("")
            lines: list[str] = []
            only = self.get_only_paths()
            if only is not None:
                lines.append(f"Limited to {len(only)} listed file(s).\n\n")
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
            self._set_preview("".join(lines))

        def on_apply_parent_tag(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                pt = scan_parent_folder_tag(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                return

            if not pt.planned:
                self._msg_info(
                    "Nothing to do",
                    "No files need a parent folder tag. Use Preview for details (e.g. files only at target root).",
                )
                return

            if not self._ask_yesno(
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
                self._msg_warning("Could not save settings", str(e))

            n, errors = apply_renames(pt.planned)
            if errors:
                self._msg_warning(
                    "Some renames failed",
                    f"Renamed: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                self._msg_info("Done", f"Renamed {n} file(s) in the target.")

            self.on_preview_parent_tag()

        def on_preview_jpg(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                self._msg_warning("Could not save settings", str(e))

            self._set_preview("Scanning .jpg files…\n")

            try:
                jpg = scan_jpg_moves(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                self._set_preview("")
                return

            self._set_preview("")
            lines: list[str] = []
            only = self.get_only_paths()
            if only is not None:
                lines.append(f"Limited to {len(only)} listed file(s).\n\n")
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
            self._set_preview("".join(lines))

        def on_apply_jpg(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                jpg = scan_jpg_moves(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                return

            if not jpg.moves:
                self._msg_info("Nothing to do", "No .jpg files to move.")
                return

            if not self._ask_yesno(
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
                self._msg_warning("Could not save settings", str(e))

            n, errors = apply_jpg_moves(jpg.moves)
            if errors:
                self._msg_warning(
                    "Some moves failed",
                    f"Moved: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                self._msg_info("Done", f"Moved {n} .jpg file(s) to target.")

            self.on_preview_jpg()

        def on_preview_copy_transfer(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()
            try:
                self.persist_paths()
            except OSError as e:
                self._msg_warning("Could not save settings", str(e))

            self._set_preview("Reading copy list…\n")

            try:
                xfer = scan_copy_transfer(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                self._set_preview("")
                return

            self._set_preview("")
            lines: list[str] = []
            only = self.get_only_paths()
            if only is not None:
                lines.append(f"Limited to {len(only)} listed file(s).\n\n")
            list_path = src_r / COPY_TRANSFER_LIST_NAME

            if xfer.list_missing:
                lines.append(
                    f"List file not found (expected at):\n  {list_path}\n"
                )
                self._set_preview("".join(lines))
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
            self._set_preview("".join(lines))

        def on_apply_copy_transfer(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                xfer = scan_copy_transfer(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                return

            if xfer.list_missing:
                self._msg_error(
                    "List missing",
                    f'Could not find "{COPY_TRANSFER_LIST_NAME}" under the source folder:\n\n'
                    f"{src_r}",
                )
                return

            if not xfer.copies:
                self._msg_info(
                    "Nothing to do",
                    "No videos to copy (all missing, ambiguous, or already at destination). "
                    "Use Preview for details.",
                )
                return

            if not self._ask_yesno(
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
                self._msg_warning("Could not save settings", str(e))

            n, errors = apply_copy_transfer(xfer.copies)
            if errors:
                self._msg_warning(
                    "Some copies failed",
                    f"Copied: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                self._msg_info("Done", f"Copied {n} video file(s) to target.")

            self.on_preview_copy_transfer()

        def on_apply(self) -> None:
            src = Path(str(dpg.get_value(self.TAG_SOURCE)).strip())
            tgt = Path(str(dpg.get_value(self.TAG_TARGET)).strip())
            err = validate_roots(src, tgt)
            if err:
                self._msg_error("Invalid paths", err)
                return
            src_r, tgt_r = src.resolve(), tgt.resolve()

            try:
                result = scan_target(src_r, tgt_r, only=self.get_only_paths())
            except OSError as e:
                self._msg_error("Scan failed", str(e))
                return

            if not result.planned:
                self._msg_info("Nothing to do", "No renames pending.")
                return

            if not self._ask_yesno(
                "Confirm",
                f"Rename {len(result.planned)} file(s) in the target folder?\n\n"
                "This step only adds a ✔ prefix to names under the target; "
                "it does not move or delete files.",
            ):
                return

            try:
                self.persist_paths()
            except OSError as e:
                self._msg_warning("Could not save settings", str(e))

            n, errors = apply_renames(result.planned)
            if errors:
                self._msg_warning(
                    "Some renames failed",
                    f"Renamed: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]),
                )
            else:
                self._msg_info("Done", f"Renamed {n} file(s).")

            self.on_preview()

    App().run()


def _cli_format_errors(errors: list[str], limit: int = 10) -> str:
    if not errors:
        return ""
    return "\n\n".join(errors[:limit]) + (
        f"\n\n… and {len(errors) - limit} more." if len(errors) > limit else ""
    )


def run_cli(argv: list[str]) -> int:
    import argparse
    import sys

    parser = argparse.ArgumentParser(
        description="Organize files (CLI). Default with no args opens the GUI."
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Open the Dear PyGui GUI (also used when no --action is given).",
    )
    parser.add_argument(
        "--action",
        choices=("mark", "title-strip", "parent-tag", "jpg-move", "copy-transfer"),
        help="Operation to run from the command line.",
    )
    parser.add_argument("--source", help="Source folder path.")
    parser.add_argument("--target", help="Target folder path.")
    parser.add_argument(
        "--strip-chars",
        default="",
        help="Characters to strip from stems (title-strip only; uses saved value if omitted).",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Scan and print a preview only (default when --apply is not set).",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Apply the planned operation after validation.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip interactive confirmation when using --apply (CLI only).",
    )
    parser.add_argument(
        "--only-list",
        metavar="FILE",
        help="Text file with one file path per line; only those files are processed.",
    )
    parser.add_argument(
        "--only-file",
        action="append",
        default=[],
        metavar="PATH",
        help="Limit to this file path (repeatable). Combined with --only-list.",
    )
    args = parser.parse_args(argv)

    if args.gui or not args.action:
        run_gui(
            initial_source=args.source,
            initial_target=args.target,
            initial_only_list=args.only_list,
            initial_only_files=args.only_file,
        )
        return 0

    saved_src, saved_tgt, saved_strip = config_load_defaults()
    src_s = (args.source or saved_src or "").strip()
    tgt_s = (args.target or saved_tgt or "").strip()
    strip_chars = normalize_strip_chars(
        args.strip_chars if args.strip_chars != "" else saved_strip
    )

    src = Path(src_s)
    tgt = Path(tgt_s)
    err = validate_roots(src, tgt)
    if err:
        print(err, file=sys.stderr)
        return 2

    src_r, tgt_r = src.resolve(), tgt.resolve()
    only = read_only_paths(args.only_list, args.only_file)
    do_apply = bool(args.apply)
    do_preview = not do_apply or args.preview

    def cli_only_banner() -> str:
        return f"Limited to {len(only)} listed file(s).\n\n" if only else ""

    try:
        if args.action == "mark":
            result = scan_target(src_r, tgt_r, only=only)
            if do_preview:
                print(cli_only_banner() + "Files to rename (add ✔ prefix):")
                if result.planned:
                    for old_path, new_path in result.planned:
                        print(f"  {old_path}")
                        print(f"    -> {new_path}")
                else:
                    print("  (none)")
                print(
                    f"\nSummary — to rename: {len(result.planned)}, "
                    f"already marked: {result.skipped_checked}, "
                    f"no source match: {result.skipped_no_match}, "
                    f"skipped (under source tree): {result.skipped_under_source}"
                )
            if do_apply:
                if not result.planned:
                    print("Nothing to do.")
                    return 0
                if not args.yes:
                    print(
                        f"Refusing --apply without --yes ({len(result.planned)} renames pending).",
                        file=sys.stderr,
                    )
                    return 2
                n, errors = apply_renames(result.planned)
                print(f"Renamed {n} file(s).")
                if errors:
                    print(_cli_format_errors(errors), file=sys.stderr)
                    return 1

        elif args.action == "title-strip":
            ts = scan_title_strip(src_r, tgt_r, strip_chars, only=only)
            if do_preview:
                print(cli_only_banner() + "Target files to rename (strip characters from stem):")
                if ts.planned:
                    for old_path, new_path in ts.planned:
                        print(f"  {old_path}")
                        print(f"    -> {new_path.name}")
                else:
                    print("  (none)")
                print(
                    f"\nSummary — to rename: {len(ts.planned)}, "
                    f"unchanged: {ts.skipped_unchanged}, "
                    f"collision: {len(ts.skipped_collision)}, "
                    f"skipped (under source tree): {ts.skipped_under_source}"
                )
            if do_apply:
                if not ts.planned:
                    print("Nothing to do.")
                    return 0
                if not args.yes:
                    print(
                        f"Refusing --apply without --yes ({len(ts.planned)} renames pending).",
                        file=sys.stderr,
                    )
                    return 2
                n, errors = apply_renames(ts.planned)
                print(f"Renamed {n} file(s).")
                if errors:
                    print(_cli_format_errors(errors), file=sys.stderr)
                    return 1

        elif args.action == "parent-tag":
            pt = scan_parent_folder_tag(src_r, tgt_r, only=only)
            if do_preview:
                print(
                    cli_only_banner()
                    + "Target files to rename (normalize [parent folder] before extension):"
                )
                if pt.planned:
                    for old_path, new_path in pt.planned:
                        print(f"  {old_path}")
                        print(f"    -> {new_path.name}")
                else:
                    print("  (none)")
                print(
                    f"\nSummary — to rename: {len(pt.planned)}, "
                    f"at target root (skipped): {pt.skipped_at_target_root}, "
                    f"already tagged / no change: {pt.skipped_already_tagged}, "
                    f"empty folder name: {pt.skipped_empty_folder_name}, "
                    f"collision: {len(pt.skipped_collision)}, "
                    f"skipped (under source tree): {pt.skipped_under_source}"
                )
            if do_apply:
                if not pt.planned:
                    print("Nothing to do.")
                    return 0
                if not args.yes:
                    print(
                        f"Refusing --apply without --yes ({len(pt.planned)} renames pending).",
                        file=sys.stderr,
                    )
                    return 2
                n, errors = apply_renames(pt.planned)
                print(f"Renamed {n} file(s).")
                if errors:
                    print(_cli_format_errors(errors), file=sys.stderr)
                    return 1

        elif args.action == "jpg-move":
            jpg = scan_jpg_moves(src_r, tgt_r, only=only)
            if do_preview:
                print(cli_only_banner() + "JPG files to move (source → target, relative path preserved):")
                if jpg.moves:
                    for old_path, new_path in jpg.moves:
                        print(f"  {old_path}")
                        print(f"    -> {new_path}")
                else:
                    print("  (none)")
                print(
                    f"\nSummary — to move: {len(jpg.moves)}, "
                    f"skipped (dest exists): {len(jpg.skipped_exists)}"
                )
            if do_apply:
                if not jpg.moves:
                    print("Nothing to do.")
                    return 0
                if not args.yes:
                    print(
                        f"Refusing --apply without --yes ({len(jpg.moves)} moves pending).",
                        file=sys.stderr,
                    )
                    return 2
                n, errors = apply_jpg_moves(jpg.moves)
                print(f"Moved {n} .jpg file(s).")
                if errors:
                    print(_cli_format_errors(errors), file=sys.stderr)
                    return 1

        elif args.action == "copy-transfer":
            xfer = scan_copy_transfer(src_r, tgt_r, only=only)
            if xfer.list_missing:
                print(
                    f'List file not found (expected "{COPY_TRANSFER_LIST_NAME}" under source):\n  {src_r / COPY_TRANSFER_LIST_NAME}',
                    file=sys.stderr,
                )
                return 2
            if do_preview:
                print(
                    cli_only_banner()
                    + f'Videos to copy (from "{COPY_TRANSFER_LIST_NAME}"; source → target):'
                )
                if xfer.copies:
                    for old_path, new_path in xfer.copies:
                        print(f"  {old_path}")
                        print(f"    -> {new_path}")
                else:
                    print("  (none)")
                print(
                    f"\nSummary — to copy: {len(xfer.copies)}, "
                    f"dest exists: {len(xfer.skipped_exists)}, "
                    f"not found: {len(xfer.missing)}, "
                    f"ambiguous: {len(xfer.ambiguous)}"
                )
            if do_apply:
                if not xfer.copies:
                    print("Nothing to do.")
                    return 0
                if not args.yes:
                    print(
                        f"Refusing --apply without --yes ({len(xfer.copies)} copies pending).",
                        file=sys.stderr,
                    )
                    return 2
                n, errors = apply_copy_transfer(xfer.copies)
                print(f"Copied {n} video file(s).")
                if errors:
                    print(_cli_format_errors(errors), file=sys.stderr)
                    return 1
    except OSError as e:
        print(str(e), file=sys.stderr)
        return 1

    if do_apply:
        try:
            config_save(src_s, tgt_s, strip_chars)
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    import sys

    if sys.platform == "win32":
        for _stream in (sys.stdin, sys.stdout, sys.stderr):
            if _stream is not None and hasattr(_stream, "reconfigure"):
                try:
                    _stream.reconfigure(encoding="utf-8")
                except (OSError, ValueError):
                    pass

    if len(sys.argv) > 1:
        raise SystemExit(run_cli(sys.argv[1:]))
    run_gui()
