"""
Mark target files with a leading ✔ when a corresponding file exists under the source tree.

Same relative folder + peeled basename must match (e.g. source `foo.mkv` ↔ target `foo.mp4.jpg`).
Also: move all `.jpg` files from source → target preserving the same relative paths (removes them from source).
Also: read `Copy and Transfer.txt` in the source folder; each line is a `.mp4.jpg` / `.wmv.jpg` thumbnail name — strip `.jpg` and copy the corresponding video into the target tree (same relative path; does not remove from source).
Also: strip chosen characters from filenames under the target (stem only), trim spaces, remove trailing dots before the extension,
collapse a duplicated final extension (e.g. ``.mp4.mp4`` → ``.mp4``; case-insensitive),
and remove trailing copy suffixes like `` (1)`` / `` (23)`` (1–3 digits; avoids `` (2024)``-style years).
Also: under the target tree, append a chosen `` [tag]`` before the file extension, or strip every ``[...]`` tag from names.

GUI (Dear PyGui): edit source paths (folders and/or files), target folder, and strip characters; drag files/folders from Explorer onto the path fields (Windows); settings are stored under %APPDATA%\\OrganizeFiles.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

CHECK = "\u2714"  # HEAVY CHECK MARK ✔

COPY_TRANSFER_LIST_NAME = "Copy and Transfer.txt"

# Trailing `` (n)`` / `` (n) (m)`` duplicate markers (Explorer-style); 1–3 digits so `` (2024)`` is kept.
_TITLE_STRIP_COPY_SUFFIX = re.compile(r"(?:\s+\(\d{1,3}\))+\Z")

# Optional space before ``[...]`` at end of stem (bracket-tag normalization).
_TRAILING_BRACKET_TAG_END = re.compile(r"(?:\s?)\[([^\]]*)\]\Z")

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "OrganizeFiles"
CONFIG_PATH = CONFIG_DIR / "settings.json"

LAST_ACTIONS = ("mark", "title-strip", "bracket-tag", "jpg-move", "copy-transfer")
LAST_MODES = ("preview", "apply")
DEFAULT_LAST_ACTION = "mark"
DEFAULT_LAST_MODE = "preview"

# GUI collapsing-section keys (saved under settings.json → gui_sections).
GUI_SECTION_DEFAULTS: dict[str, bool] = {
    "folders": True,
    "mark": True,
    "title": False,
    "bracket": False,
    "jpg": False,
    "copy": False,
}
GUI_SECTION_KEYS = tuple(GUI_SECTION_DEFAULTS)

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


def config_read() -> dict:
    if not CONFIG_PATH.is_file():
        return {}
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def config_write(data: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def config_load_last() -> tuple[str, str]:
    """Last GUI/CLI operation: action name and preview or apply."""
    data = config_read()
    action = str(data.get("last_action") or DEFAULT_LAST_ACTION)
    mode = str(data.get("last_mode") or DEFAULT_LAST_MODE)
    if action not in LAST_ACTIONS:
        action = DEFAULT_LAST_ACTION
    if mode not in LAST_MODES:
        mode = DEFAULT_LAST_MODE
    return action, mode


def config_record_last(action: str, mode: str) -> None:
    if action not in LAST_ACTIONS or mode not in LAST_MODES:
        return
    data = config_read()
    data["last_action"] = action
    data["last_mode"] = mode
    try:
        config_write(data)
    except OSError:
        pass


def config_load_defaults() -> tuple[str, str, str, str]:
    """Load saved source paths text, target folder, strip-title characters, and bracket tag text."""
    try:
        data = config_read()
        src = str(data.get("source") or "").strip()
        tgt = str(data.get("target") or "")
        strip = normalize_strip_chars(data.get("strip_title_chars") or "")
        tag = str(data.get("bracket_tag_text") or "").strip()
        return src, tgt, strip, tag
    except OSError:
        return "", "", "", ""


def config_load_gui_sections() -> dict[str, bool]:
    """Which action panels were expanded last time the GUI closed."""
    out = dict(GUI_SECTION_DEFAULTS)
    raw = config_read().get("gui_sections")
    if not isinstance(raw, dict):
        return out
    for key in GUI_SECTION_KEYS:
        if key in raw:
            out[key] = bool(raw[key])
    return out


def config_save(
    source: str,
    target: str,
    strip_title_chars: str = "",
    bracket_tag_text: str = "",
    gui_sections: Optional[dict[str, bool]] = None,
) -> None:
    data = config_read()
    data["source"] = source
    data["target"] = target
    data["strip_title_chars"] = strip_title_chars
    data["bracket_tag_text"] = bracket_tag_text
    if gui_sections is not None:
        data["gui_sections"] = {
            k: bool(gui_sections[k]) for k in GUI_SECTION_KEYS if k in gui_sections
        }
    config_write(data)


def is_resolved_subpath(parent: Path, child: Path) -> bool:
    """True if child is equal to or inside parent (both resolved)."""
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def path_in_source_library(source_library: Optional[Path], path: Path) -> bool:
    """True when path lies under an explicit source-folder line (not file-only scope roots)."""
    if source_library is None:
        return False
    return is_resolved_subpath(source_library, path)


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
    """Map relative parent dir (as path parts) -> case-folded peeled stems for matching."""
    idx: dict[tuple[str, ...], set[str]] = defaultdict(set)
    for path in source_root.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(source_root)
        parent = rel.parent.parts
        idx[parent].update(v.casefold() for v in stem_variants(path.name))
    return idx


def trees_are_separate(source_root: Path, target_root: Path) -> bool:
    """True when neither tree contains the other (typical two-library layout)."""
    if source_root == target_root:
        return False
    return not is_resolved_subpath(source_root, target_root) and not is_resolved_subpath(
        target_root, source_root
    )


def files_in_directory(dir_path: Path) -> set[Path]:
    return {p.resolve() for p in dir_path.rglob("*") if p.is_file()}


def match_parent_key(source_root: Path, target_root: Path, path: Path) -> tuple[str, ...]:
    """Relative parent dir for source↔target matching (works outside target_root)."""
    resolved = path.resolve()
    if is_resolved_subpath(target_root, resolved):
        return resolved.relative_to(target_root).parent.parts
    if is_resolved_subpath(source_root, resolved):
        return resolved.relative_to(source_root).parent.parts
    return resolved.parent.parts


def rename_stays_valid(
    source_root: Path,
    target_root: Path,
    old_path: Path,
    new_path: Path,
    source_library: Optional[Path] = None,
) -> bool:
    """Allow in-place renames; require new path to stay under target when old path did."""
    resolved = old_path.resolve()
    resolved_new = new_path.resolve()
    if path_in_source_library(source_library, resolved_new):
        return False
    if is_resolved_subpath(target_root, resolved):
        return is_resolved_subpath(target_root, resolved_new)
    return True


def has_source_match(
    source_index: dict[tuple[str, ...], set[str]],
    parent_key: tuple[str, ...],
    target_filename: str,
) -> bool:
    stems = source_index.get(parent_key)
    if not stems:
        return False
    target_stems = {v.casefold() for v in stem_variants(target_filename)}
    return bool(stems & target_stems)


def _planned_dest_key(path: Path) -> str:
    """Case-insensitive destination identity for batch rename collision checks."""
    return os.path.normcase(os.fspath(path))


def _try_plan_rename(
    old_path: Path,
    new_path: Path,
    planned: list[tuple[Path, Path]],
    skipped_collision: list[tuple[Path, Path]],
    planned_dest_keys: set[str],
) -> bool:
    """Append a planned rename unless the destination exists or is already taken this batch."""
    if new_path.exists():
        skipped_collision.append((old_path, new_path))
        return False
    key = _planned_dest_key(new_path)
    if key in planned_dest_keys:
        skipped_collision.append((old_path, new_path))
        return False
    planned_dest_keys.add(key)
    planned.append((old_path, new_path))
    return True


def collect_paths_from_lines(lines: list[str]) -> tuple[list[Path], set[Path], Optional[str]]:
    """Parse lines: directories expand to all files inside; files are kept as-is."""
    dirs: list[Path] = []
    files: set[Path] = set()
    for s in lines:
        s = s.strip()
        if not s:
            continue
        p = Path(s)
        if p.is_dir():
            dirs.append(p.resolve())
        elif p.is_file():
            files.add(p.resolve())
        else:
            return [], set(), f"Path not found:\n{p}"
    scope = set(files)
    for d in dirs:
        scope.update(files_in_directory(d))
    return dirs, scope, None


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
    not_in_selection: list[str]
    ambiguous: list[tuple[str, list[Path]]]
    list_missing: bool


@dataclass
class RenameScan:
    planned: list[tuple[Path, Path]]
    skipped_no_change: int
    skipped_collision: list[tuple[Path, Path]]
    skipped_under_source: int


_EMPTY_RENAME_SCAN = RenameScan([], 0, [], 0)


def sanitize_tag_text(text: str) -> str:
    """Make tag text safe inside `` […] `` in a filename (Windows-forbidden chars)."""
    out = text
    for ch in '\\/:*?"<>|':
        out = out.replace(ch, "_")
    return out.strip().rstrip(".")


_ANY_BRACKET_TAG = re.compile(r"\s?\[([^\]]*)\]")


def _strip_bracket_tag_spans(stem: str, spans: list[tuple[int, int]]) -> str:
    if not spans:
        return stem
    parts: list[str] = []
    last = 0
    for start, end in spans:
        parts.append(stem[last:start])
        last = end
    parts.append(stem[last:])
    merged = "".join(parts)
    return re.sub(r" {2,}", " ", merged).strip()


def remove_all_bracket_tags(stem: str) -> str:
    """Remove every ``[...]`` / `` [...]`` bracket tag in the stem."""
    to_remove = [(m.start(), m.end()) for m in _ANY_BRACKET_TAG.finditer(stem)]
    return _strip_bracket_tag_spans(stem, to_remove)


def stem_final_trailing_bracket_inner(stem: str) -> Optional[str]:
    """Inner text of ``[...]`` at end of stem, or None."""
    m = _TRAILING_BRACKET_TAG_END.search(stem)
    return m.group(1).strip() if m else None


def bracket_tag_new_name(path: Path, tag_inner: str) -> Optional[str]:
    """Append `` [tag]`` at the end of the stem; leave other bracket tags unchanged."""
    if not path.name or not tag_inner:
        return None

    suffix = path.suffix
    stem = path.stem
    tail = stem_final_trailing_bracket_inner(stem)
    if tail is not None and tail.casefold() == tag_inner.casefold():
        return None

    base = stem.rstrip()
    new_stem = (base + f" [{tag_inner}]") if base else f"[{tag_inner}]"
    new_name = new_stem + suffix
    if new_name == path.name:
        return None
    return new_name


def bracket_tag_remove_all_name(path: Path) -> Optional[str]:
    """Remove every ``[...]`` bracket tag from the stem."""
    if not path.name:
        return None

    suffix = path.suffix
    stem_clean = remove_all_bracket_tags(path.stem)
    if stem_clean == path.stem:
        return None
    return stem_clean + suffix


def _scan_target_renames(
    source_root: Path,
    target_root: Path,
    new_name_for: Callable[[Path], Optional[str]],
    only: Optional[set[Path]] = None,
    source_library: Optional[Path] = None,
) -> RenameScan:
    planned: list[tuple[Path, Path]] = []
    skipped_collision: list[tuple[Path, Path]] = []
    skipped_no_change = 0
    skipped_under_source = 0
    planned_dest_keys: set[str] = set()

    for path in iter_files_in_tree(target_root, only):
        if path_in_source_library(source_library, path):
            skipped_under_source += 1
            continue
        new_name = new_name_for(path)
        if new_name is None:
            skipped_no_change += 1
            continue
        new_path = path.with_name(new_name)
        if not rename_stays_valid(
            source_root, target_root, path, new_path, source_library=source_library
        ):
            skipped_under_source += 1
            continue
        _try_plan_rename(path, new_path, planned, skipped_collision, planned_dest_keys)

    return RenameScan(
        planned=planned,
        skipped_no_change=skipped_no_change,
        skipped_collision=skipped_collision,
        skipped_under_source=skipped_under_source,
    )


def scan_bracket_tag(
    source_root: Path,
    target_root: Path,
    tag_text: str,
    only: Optional[set[Path]] = None,
    source_library: Optional[Path] = None,
) -> RenameScan:
    """Plan renames: append `` [tag_text]`` at end of each target-tree filename stem."""
    tag_inner = sanitize_tag_text(tag_text)
    if not tag_inner:
        return _EMPTY_RENAME_SCAN
    return _scan_target_renames(
        source_root,
        target_root,
        lambda p: bracket_tag_new_name(p, tag_inner),
        only=only,
        source_library=source_library,
    )


def scan_bracket_tag_remove_all(
    source_root: Path,
    target_root: Path,
    only: Optional[set[Path]] = None,
    source_library: Optional[Path] = None,
) -> RenameScan:
    """Plan renames: strip every ``[...]`` tag from target-tree filenames."""
    return _scan_target_renames(
        source_root,
        target_root,
        bracket_tag_remove_all_name,
        only=only,
        source_library=source_library,
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
    source_library: Optional[Path] = None,
) -> RenameScan:
    """Plan renames under target_root only (stem cleanup). Skips paths inside source_root."""
    return _scan_target_renames(
        source_root,
        target_root,
        lambda p: transform_title_filename(p.name, strip_chars),
        only=only,
        source_library=source_library,
    )


def parse_only_path_lines(lines: list[str]) -> set[Path]:
    """Resolved files from path strings (files or folders — folders expand to all files inside)."""
    _, scope, err = collect_paths_from_lines(lines)
    return scope if not err else set()


def common_root_for_files(files: set[Path]) -> Optional[Path]:
    if not files:
        return None
    if len(files) == 1:
        return next(iter(files)).parent.resolve()
    try:
        common = Path(os.path.commonpath([os.fspath(f) for f in files])).resolve()
        return common.parent.resolve() if common.is_file() else common
    except ValueError:
        return None


def parse_source_input_lines(
    lines: list[str],
    target_root: Optional[Path] = None,
) -> tuple[Optional[Path], Optional[Path], Optional[set[Path]], Optional[str]]:
    """
    One path per line: a file = that file only; a folder = every file in that folder.
    With two separate library trees (source vs target), a lone folder line limits the run
    to files inside that folder. When the folder is the source library and overlaps the
    target tree, ``only`` is None so mark scans the full target tree.
    Returns (source_root, source_library or None, only_files or None for full tree, error).
    ``source_library`` is set only for an explicit folder line; file-only runs use the
    common parent as ``source_root`` for matching but must not treat listed files as
    "under the source tree" for target-only renames.
    """
    dirs: list[Path] = []
    file_lines: set[Path] = set()
    scope: set[Path] = set()
    for s in lines:
        s = s.strip()
        if not s:
            continue
        p = Path(s)
        if p.is_dir():
            d = p.resolve()
            dirs.append(d)
            scope.update(files_in_directory(d))
        elif p.is_file():
            f = p.resolve()
            file_lines.add(f)
            scope.add(f)
        else:
            return None, None, None, f"Path not found:\n{p}"

    if len(dirs) > 1:
        return None, None, None, "Only one source folder per run (one directory line)."

    if dirs:
        source_root = dirs[0]
        source_library = source_root
        if file_lines:
            return source_root, source_library, scope, None
        if target_root is not None and trees_are_separate(source_root, target_root.resolve()):
            return source_root, source_library, scope, None
        return source_root, source_library, None, None

    if scope:
        root = common_root_for_files(scope)
        if root is None:
            return (
                None,
                None,
                None,
                "Listed files must share one folder (same drive on Windows).",
            )
        return root, None, scope, None

    return None, None, None, "Add a source folder or at least one file path."


@dataclass
class WorkPaths:
    source_root: Path
    target_root: Path
    only: Optional[set[Path]]
    source_library: Optional[Path] = None


def resolve_work_paths(
    source_text: str, target_text: str
) -> tuple[Optional[WorkPaths], Optional[str]]:
    tgt_s = target_text.strip()
    if not tgt_s:
        return None, "Target folder is not set."
    tgt = Path(tgt_s)
    if not tgt.is_dir():
        return None, f"Target is not a directory:\n{tgt}"

    source_root, source_library, only, err = parse_source_input_lines(
        source_text.splitlines(), target_root=tgt.resolve()
    )
    if err:
        return None, err

    err = validate_roots(source_root, tgt.resolve())
    if err:
        return None, err

    lib = source_library.resolve() if source_library is not None else None
    return (
        WorkPaths(source_root.resolve(), tgt.resolve(), only, source_library=lib),
        None,
    )


def _only_list_lines(
    list_path: Optional[str] = None, file_args: Optional[list[str]] = None
) -> list[str]:
    """Raw path strings from --only-list / --only-file (no existence check)."""
    lines: list[str] = []
    for s in file_args or []:
        t = str(s).strip()
        if t:
            lines.append(t)
    if list_path:
        lp = Path(list_path)
        if lp.is_file():
            for ln in lp.read_text(encoding="utf-8-sig", errors="replace").splitlines():
                t = ln.strip()
                if t:
                    lines.append(t)
    return lines


def _dedupe_path_lines(lines: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for ln in lines:
        key = ln.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(ln)
    return out


def build_initial_source_text(
    saved_source: str,
    initial_source: Optional[str],
    initial_only_list: Optional[str],
    initial_only_files: Optional[list[str]],
) -> str:
    from_dopus = _dedupe_path_lines(
        _only_list_lines(initial_only_list, initial_only_files)
    )
    if from_dopus:
        return "\n".join(from_dopus)

    lines: list[str] = []
    for ln in saved_source.splitlines():
        s = ln.strip()
        if s:
            lines.append(s)
    if initial_source and str(initial_source).strip():
        lines.append(str(initial_source).strip())
    return "\n".join(_dedupe_path_lines(lines))


def read_only_paths(
    list_path: Optional[str] = None, file_args: Optional[list[str]] = None
) -> Optional[set[Path]]:
    """Resolved file paths to limit scans; None means process the full tree."""
    lines = _only_list_lines(list_path, file_args)
    if not lines:
        return None
    out = parse_only_path_lines(lines)
    return out if out else None


def iter_files_in_tree(root: Path, only: Optional[set[Path]] = None):
    """
    Yield files under root (full tree when only is None).
    When only is set, yield every listed file even if it is outside root.
    """
    if only is not None:
        for path in sorted(only):
            if path.is_file():
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
    source_root: Path,
    target_root: Path,
    only: Optional[set[Path]] = None,
    source_library: Optional[Path] = None,
) -> ScanResult:
    source_index = index_source(source_root)
    skipped_checked = 0
    skipped_no_match = 0
    skipped_under_source = 0
    planned: list[tuple[Path, Path]] = []
    skipped_exists: list[Path] = []

    for path in iter_files_in_tree(target_root, only):
        if path_in_source_library(source_library, path):
            skipped_under_source += 1
            continue
        name = path.name
        if name.startswith(CHECK):
            skipped_checked += 1
            continue
        parent_key = match_parent_key(source_root, target_root, path)
        if not has_source_match(source_index, parent_key, name):
            skipped_no_match += 1
            continue
        new_path = path.with_name(f"{CHECK}{name}")
        if not rename_stays_valid(
            source_root, target_root, path, new_path, source_library=source_library
        ):
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


def _rename_lookup_keys(path: Path) -> list[str]:
    keys = [os.fspath(path).casefold(), path.as_posix().casefold()]
    try:
        resolved = path.resolve()
        keys.append(os.fspath(resolved).casefold())
        keys.append(resolved.as_posix().casefold())
    except OSError:
        pass
    return keys


def remap_path_lines(
    lines: list[str], renames: list[tuple[Path, Path]]
) -> list[str]:
    """Replace source-input lines that match a renamed path with the new path."""
    if not renames:
        return lines
    lookup: dict[str, str] = {}
    for old_path, new_path in renames:
        new_s = os.fspath(new_path)
        for key in _rename_lookup_keys(old_path):
            lookup[key] = new_s
    out: list[str] = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            out.append(line)
            continue
        hit = None
        for key in _rename_lookup_keys(Path(stripped)):
            if key in lookup:
                hit = lookup[key]
                break
        out.append(hit if hit is not None else line)
    return out


def apply_renames(
    planned: list[tuple[Path, Path]],
) -> tuple[int, list[str], list[tuple[Path, Path]]]:
    """Returns (success_count, error_messages, successfully_renamed pairs)."""
    errors: list[str] = []
    done: list[tuple[Path, Path]] = []
    n = 0
    for old_path, new_path in planned:
        try:
            old_path.rename(new_path)
            n += 1
            done.append((old_path, new_path))
        except OSError as e:
            errors.append(f"{old_path}\n  {e}")
    return n, errors, done


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
        if not is_resolved_subpath(source_root, resolved):
            continue
        rel = resolved.relative_to(source_root)
        dest = (target_root / rel).resolve()
        if not is_resolved_subpath(target_root, dest):
            continue
        if dest.exists():
            skipped_exists.append((path, dest))
            continue
        moves.append((path, dest))

    return JpgMoveScan(moves=moves, skipped_exists=skipped_exists)


def _apply_file_pairs(
    pairs: list[tuple[Path, Path]], *, copy: bool
) -> tuple[int, list[str]]:
    errors: list[str] = []
    n = 0
    for src, dst in pairs:
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            if copy:
                shutil.copy2(os.fspath(src), os.fspath(dst))
            else:
                shutil.move(os.fspath(src), os.fspath(dst))
            n += 1
        except OSError as e:
            errors.append(f"{src}\n  -> {dst}\n  {e}")
    return n, errors


def apply_jpg_moves(moves: list[tuple[Path, Path]]) -> tuple[int, list[str]]:
    """Returns (success_count, error_messages). Removes each file from source after move."""
    return _apply_file_pairs(moves, copy=False)


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
            not_in_selection=[],
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
    not_in_selection: list[str] = []
    ambiguous: list[tuple[str, list[Path]]] = []

    for base in wanted:
        all_matches = by_name.get(os.path.normcase(base), [])
        if not all_matches:
            missing.append(base)
            continue
        if len(all_matches) > 1:
            ambiguous.append((base, all_matches))
            continue
        src = all_matches[0]
        if only is not None and src.resolve() not in only:
            not_in_selection.append(base)
            continue
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
        not_in_selection=not_in_selection,
        ambiguous=ambiguous,
        list_missing=False,
    )


def _browse_initial_dir(hint: str) -> str:
    hint = hint.strip()
    if not hint:
        return ""
    p = Path(hint)
    if p.is_dir():
        return str(p)
    if p.is_file():
        return str(p.parent)
    parent = p.parent
    return str(parent) if parent.is_dir() else ""


def _pick_native_dialog(folder: bool, title: str, initial: str = "") -> list[str]:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    kwargs: dict = {"title": title, "parent": root}
    init = _browse_initial_dir(initial)
    if init:
        kwargs["initialdir"] = init
    if folder:
        kwargs["mustexist"] = True
        path = filedialog.askdirectory(**kwargs)
        root.destroy()
        return [path] if path else []
    paths = filedialog.askopenfilenames(**kwargs)
    root.destroy()
    return list(paths) if paths else []


def pick_native_folder(title: str, initial: str = "") -> Optional[str]:
    """Windows folder picker (tkinter uses the system dialog on win32)."""
    paths = _pick_native_dialog(True, title, initial)
    return paths[0] if paths else None


def pick_native_files(title: str, initial: str = "") -> list[str]:
    """Windows multi-file picker (tkinter uses the system dialog on win32)."""
    return _pick_native_dialog(False, title, initial)


def apply_copy_transfer(copies: list[tuple[Path, Path]]) -> tuple[int, list[str]]:
    return _apply_file_pairs(copies, copy=True)


def _scope_banner(only: Optional[set[Path]]) -> str:
    return f"Limited to {len(only)} listed file(s).\n\n" if only else ""


def _format_rename_collision_block(
    collisions: list[tuple[Path, Path]], *, basename_only: bool
) -> str:
    if not collisions:
        return ""
    lines = [
        "\nSkipped (destination already exists on disk or duplicate in this batch):\n"
    ]
    for old_path, dest_path in collisions:
        left = old_path.name if basename_only else str(old_path)
        lines.append(f"  {left}\n")
        lines.append(f"    -> {dest_path.name}\n")
    return "".join(lines)


def format_preview_mark(result: ScanResult, only: Optional[set[Path]]) -> str:
    lines = [_scope_banner(only)]
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
    return "".join(lines)


def format_preview_rename(
    rs: RenameScan,
    only: Optional[set[Path]],
    heading: str,
    skip_label: str,
    *,
    collision_basename_only: bool,
) -> str:
    lines = [_scope_banner(only), heading]
    if rs.planned:
        for old_path, new_path in rs.planned:
            lines.append(f"  {old_path}\n")
            lines.append(f"    -> {new_path.name}\n")
    else:
        lines.append("  (none)\n")
    lines.append(_format_rename_collision_block(rs.skipped_collision, basename_only=collision_basename_only))
    lines.append(
        f"\nSummary — to rename: {len(rs.planned)}, "
        f"{skip_label}: {rs.skipped_no_change}, "
        f"collision: {len(rs.skipped_collision)}, "
        f"skipped (under source tree): {rs.skipped_under_source}\n"
    )
    return "".join(lines)


def format_preview_jpg(jpg: JpgMoveScan, only: Optional[set[Path]]) -> str:
    lines = [_scope_banner(only), "JPG files to move (source → target, relative path preserved):\n"]
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
        f"\nSummary — to move: {len(jpg.moves)}, "
        f"skipped (dest exists): {len(jpg.skipped_exists)}\n"
    )
    return "".join(lines)


def format_preview_copy_transfer(
    xfer: CopyTransferScan, only: Optional[set[Path]], source_root: Path
) -> str:
    lines = [_scope_banner(only)]
    list_path = source_root / COPY_TRANSFER_LIST_NAME
    if xfer.list_missing:
        lines.append(f"List file not found (expected at):\n  {list_path}\n")
        return "".join(lines)
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
    if xfer.not_in_selection:
        lines.append(
            f"\nFound under source but not in selected files ({len(xfer.not_in_selection)}):\n"
        )
        for name in xfer.not_in_selection[:80]:
            lines.append(f"  {name}\n")
        if len(xfer.not_in_selection) > 80:
            lines.append(f"  … and {len(xfer.not_in_selection) - 80} more\n")
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
        f"not in selection: {len(xfer.not_in_selection)}, "
        f"ambiguous: {len(xfer.ambiguous)}\n"
    )
    return "".join(lines)


def _cli_apply_items(
    action: str, items: list[tuple[Path, Path]]
) -> tuple[int, list[str], int]:
    """Run apply for action; return (success_count, errors, exit_code)."""
    if not items:
        print("Nothing to do.")
        return 0, [], 0
    if action in ("mark", "title-strip", "bracket-tag"):
        n, errors, _ = apply_renames(items)
        print(f"Renamed {n} file(s).")
        return n, errors, 1 if errors else 0
    if action == "jpg-move":
        n, errors = apply_jpg_moves(items)
        print(f"Moved {n} .jpg file(s).")
        return n, errors, 1 if errors else 0
    n, errors = apply_copy_transfer(items)
    print(f"Copied {n} video file(s).")
    return n, errors, 1 if errors else 0


def _cli_scan_and_format(
    action: str,
    work: WorkPaths,
    strip_chars: str,
    tag_text: str,
) -> tuple[object, str, Optional[int]]:
    """Scan for CLI action. Returns (result, preview_text, early_exit_code)."""
    src_r, tgt_r, only, src_lib = (
        work.source_root,
        work.target_root,
        work.only,
        work.source_library,
    )
    if action == "mark":
        result = scan_target(src_r, tgt_r, only=only, source_library=src_lib)
        return result, format_preview_mark(result, only), None
    if action == "title-strip":
        result = scan_title_strip(
            src_r, tgt_r, strip_chars, only=only, source_library=src_lib
        )
        return (
            result,
            format_preview_rename(
                result,
                only,
                "Target files to rename (strip characters from stem):\n",
                "unchanged",
                collision_basename_only=True,
            ),
            None,
        )
    if action == "bracket-tag":
        if not tag_text:
            return None, "", 2
        result = scan_bracket_tag(
            src_r, tgt_r, tag_text, only=only, source_library=src_lib
        )
        return (
            result,
            format_preview_rename(
                result,
                only,
                f'Target files to rename (append " [{tag_text}]" at end):\n',
                "already has tag at end / no change",
                collision_basename_only=False,
            ),
            None,
        )
    if action == "jpg-move":
        result = scan_jpg_moves(src_r, tgt_r, only=only)
        return result, format_preview_jpg(result, only), None
    result = scan_copy_transfer(src_r, tgt_r, only=only)
    if result.list_missing:
        print(
            f'List file not found (expected "{COPY_TRANSFER_LIST_NAME}" under source):\n'
            f"  {src_r / COPY_TRANSFER_LIST_NAME}",
            file=sys.stderr,
        )
        return result, "", 2
    return (
        result,
        format_preview_copy_transfer(result, only, src_r),
        None,
    )


def _cli_planned_items(action: str, result: object) -> list[tuple[Path, Path]]:
    if action == "mark":
        return result.planned  # type: ignore[attr-defined]
    if action in ("title-strip", "bracket-tag"):
        return result.planned  # type: ignore[attr-defined]
    if action == "jpg-move":
        return result.moves  # type: ignore[attr-defined]
    return result.copies  # type: ignore[attr-defined]


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
        TAG_BRACKET = "bracket_tag_input"
        TAG_PREVIEW = "preview_text"

        TIP_FOLDERS = (
            "Source paths and target folder for every operation.\n"
            "Paths are saved when you close the app (%APPDATA%\\OrganizeFiles)."
        )
        TIP_SOURCE = (
            "One path per line. A folder = every file inside that folder. "
            "A file = only that file (paths can be outside the target folder). "
            "Launch from Directory Opus to fill this from your selection. "
            "Drag files or folders from Explorer onto this box to append paths."
        )
        TIP_TARGET = (
            "Folder tree where renames, tags, and incoming files are applied. "
            "Drag a folder from Explorer onto this box (dropping a file uses its parent folder)."
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
        TIP_BRACKET = (
            'Target only: type any tag text (e.g. NQ or Episode 3).\n'
            'Preview / Apply — append " [tag]" at the end (other bracket tags stay; skipped if '
            "that tag is already the final trailing tag).\n"
            "Remove all tags — strip every […] from filenames."
        )
        TIP_BRACKET_TEXT = (
            "Text inside the brackets (without the brackets). Windows-forbidden characters "
            'are replaced with _. Example: NQ → "movie [NQ].mp4".'
        )
        TIP_BRACKET_APPLY = (
            'Append " [tag]" at the end of each target filename. Other bracket tags are left '
            "as-is; skipped if that tag is already the final trailing tag."
        )
        TIP_BRACKET_REMOVE_ALL = (
            "Remove every […] bracket tag from target filenames (any tag text, anywhere in the name)."
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
            s_default, t_default, strip_default, tag_default = config_load_defaults()
            if initial_target and str(initial_target).strip():
                t_default = str(initial_target).strip()
            s_default = build_initial_source_text(
                s_default, initial_source, initial_only_list, initial_only_files
            )
            self._theme_apply = "theme_btn_apply"
            self._theme_preview = "theme_btn_preview"
            self._gui_sections = config_load_gui_sections()
            self._section_tags: dict[str, int | str] = {}

            dpg.create_context()
            self._init_os_drag_drop()
            self._build_themes()

            with dpg.window(tag="primary_window", label="Organize Files", no_title_bar=True):
                self._build_main_layout(s_default, t_default, strip_default, tag_default)
            self._build_fonts()

            dpg.create_viewport(
                title="Organize Files",
                width=1080,
                height=740,
                min_width=780,
                min_height=520,
            )
            self._register_os_drag_drop_handlers()
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

            with dpg.theme(tag="theme_drop_hover"):
                with dpg.theme_component(dpg.mvInputText):
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (40, 68, 82))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, (52, 88, 108))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (72, 168, 190))

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
            default: str,
            tip: str,
            browse_callback,
        ) -> None:
            dpg.add_text(label, color=(150, 158, 175))
            with dpg.group(horizontal=True):
                dpg.add_input_text(tag=input_tag, default_value=default, width=-48)
                btn = dpg.add_button(label="…", width=40, callback=browse_callback)
            self._hover_tip(input_tag, tip)
            self._hover_tip(btn, f"Browse for {label.lower()}")

        def _section(self, title: str, tip: str, section_key: str):
            hdr = dpg.add_collapsing_header(
                label=title,
                default_open=self._gui_sections.get(
                    section_key, GUI_SECTION_DEFAULTS[section_key]
                ),
                tag=dpg.generate_uuid(),
            )
            self._section_tags[section_key] = hdr
            self._hover_tip(hdr, tip)
            return hdr

        def _build_main_layout(
            self,
            s_default: str,
            t_default: str,
            strip_default: str,
            tag_default: str,
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

                    hdr_folders = self._section("Input/Output", self.TIP_FOLDERS, "folders")
                    with dpg.group(parent=hdr_folders):
                        dpg.add_text("Paths", color=(150, 158, 175))
                        src_input = dpg.add_input_text(
                            tag=self.TAG_SOURCE,
                            default_value=s_default,
                            multiline=True,
                            width=-1,
                            height=100,
                            tab_input=False,
                        )
                        self._hover_tip(src_input, self.TIP_SOURCE)
                        with dpg.group(horizontal=True):
                            folder_btn = dpg.add_button(
                                label="Add folder…",
                                callback=self._browse_add_source_folder,
                            )
                            files_btn = dpg.add_button(
                                label="Add files…",
                                callback=self._browse_add_source_files,
                            )
                            clear_btn = dpg.add_button(
                                label="Clear",
                                callback=self._on_clear_source_paths,
                            )
                        self._hover_tip(folder_btn, "Append a source folder to the list.")
                        self._hover_tip(files_btn, "Append one or more file paths.")
                        self._hover_tip(clear_btn, "Clear all source paths.")
                        dpg.add_spacer(height=4)
                        self._path_field(
                            "Target",
                            self.TAG_TARGET,
                            t_default,
                            self.TIP_TARGET,
                            self._browse_target_folder,
                        )

                    dpg.add_spacer(height=8)

                    hdr_mark = self._section(f"{CHECK}  Match marks", self.TIP_MARK, "mark")
                    with dpg.group(parent=hdr_mark):
                        self._preview_apply_row(
                            "Preview",
                            "Apply",
                            self.on_preview,
                            self.on_apply,
                            "Scan target tree and list files that would get a ✔ prefix.",
                            "Add ✔ prefix to matched target files (no moves/deletes).",
                        )

                    hdr_title = self._section("Title cleanup", self.TIP_TITLE, "title")
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

                    hdr_bracket = self._section("Filename tags", self.TIP_BRACKET, "bracket")
                    with dpg.group(parent=hdr_bracket):
                        dpg.add_text("Tag text", color=(150, 158, 175))
                        dpg.add_input_text(
                            tag=self.TAG_BRACKET,
                            default_value=tag_default,
                            width=-1,
                            hint="e.g. NQ",
                        )
                        self._hover_tip(self.TAG_BRACKET, self.TIP_BRACKET_TEXT)
                        dpg.add_spacer(height=4)
                        self._preview_apply_row(
                            "Preview",
                            "Apply",
                            self.on_preview_bracket_tag,
                            self.on_apply_bracket_tag,
                            'List files that would get " [tag]" appended at the end.',
                            self.TIP_BRACKET_APPLY,
                        )
                        dpg.add_spacer(height=4)
                        btn_remove = dpg.add_button(
                            label="Remove all tags",
                            callback=self.on_apply_bracket_tag_remove_all,
                            width=-1,
                        )
                        dpg.bind_item_theme(btn_remove, self._theme_apply)
                        self._hover_tip(btn_remove, self.TIP_BRACKET_REMOVE_ALL)

                    hdr_jpg = self._section("JPG transfer", self.TIP_JPG, "jpg")
                    with dpg.group(parent=hdr_jpg):
                        self._preview_apply_row(
                            "Preview",
                            "Move",
                            self.on_preview_jpg,
                            self.on_apply_jpg,
                            "List .jpg files that would move source → target.",
                            "Move .jpg files (removed from source).",
                        )

                    hdr_copy = self._section("Copy from list", self.TIP_COPY, "copy")
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

        def _source_browse_hint(self) -> str:
            for line in self._source_list_lines():
                s = line.strip()
                if s:
                    return s
            return str(dpg.get_value(self.TAG_TARGET)).strip()

        def _browse_add_source_folder(self) -> None:
            if sys.platform != "win32":
                self._msg("Browse", "Folder picker is only supported on Windows.")
                return
            path = pick_native_folder("Select source folder", self._source_browse_hint())
            if path:
                self._merge_source_paths([path])

        def _browse_add_source_files(self) -> None:
            if sys.platform != "win32":
                self._msg("Browse", "File picker is only supported on Windows.")
                return
            paths = pick_native_files("Select files", self._source_browse_hint())
            if paths:
                self._merge_source_paths(paths)

        def _browse_target_folder(self) -> None:
            if sys.platform != "win32":
                self._msg("Browse", "Folder picker is only supported on Windows.")
                return
            current = str(dpg.get_value(self.TAG_TARGET)).strip()
            path = pick_native_folder("Select target folder", current)
            if path:
                dpg.set_value(self.TAG_TARGET, path)

        def _init_os_drag_drop(self) -> None:
            if sys.platform != "win32":
                return
            import DearPyGui_DragAndDrop as dpg_dnd

            dpg_dnd.initialize()

        def _register_os_drag_drop_handlers(self) -> None:
            if sys.platform != "win32":
                return
            import DearPyGui_DragAndDrop as dpg_dnd

            dpg_dnd.set_drop(self._on_os_drop)
            dpg_dnd.set_drag_over(self._on_os_drag_over)
            dpg_dnd.set_drag_leave(self._on_os_drag_leave)

        def _path_as_target_folder(self, path: str) -> Optional[str]:
            p = Path(path)
            if p.is_dir():
                return os.fspath(p)
            if p.is_file():
                return os.fspath(p.parent)
            return None

        def _apply_target_drop(self, paths: list[str]) -> None:
            for raw in paths:
                folder = self._path_as_target_folder(raw)
                if folder:
                    dpg.set_value(self.TAG_TARGET, folder)
                    return

        def _clear_drop_hover_themes(self) -> None:
            for tag in (self.TAG_SOURCE, self.TAG_TARGET):
                if dpg.does_item_exist(tag):
                    dpg.bind_item_theme(tag, None)

        def _on_os_drop(self, data, keys) -> None:
            if not isinstance(data, list):
                return
            paths = [str(p).strip() for p in data if str(p).strip()]
            if not paths:
                return
            if dpg.is_item_hovered(self.TAG_TARGET):
                self._apply_target_drop(paths)
            elif dpg.is_item_hovered(self.TAG_SOURCE):
                self._merge_source_paths(paths)
            self._clear_drop_hover_themes()

        def _on_os_drag_over(self, keys) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd

            if dpg.is_item_hovered(self.TAG_TARGET):
                dpg.bind_item_theme(self.TAG_TARGET, "theme_drop_hover")
                dpg.bind_item_theme(self.TAG_SOURCE, None)
                dpg_dnd.set_drop_effect(dpg_dnd.DROPEFFECT.MOVE)
            elif dpg.is_item_hovered(self.TAG_SOURCE):
                dpg.bind_item_theme(self.TAG_SOURCE, "theme_drop_hover")
                dpg.bind_item_theme(self.TAG_TARGET, None)
                dpg_dnd.set_drop_effect(dpg_dnd.DROPEFFECT.MOVE)
            else:
                self._clear_drop_hover_themes()
                dpg_dnd.set_drop_effect()

        def _on_os_drag_leave(self) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd

            self._clear_drop_hover_themes()
            dpg_dnd.set_drop_effect()

        def _source_list_lines(self) -> list[str]:
            return str(dpg.get_value(self.TAG_SOURCE)).splitlines()

        def _merge_source_paths(self, new_paths: list[str]) -> None:
            lines = [ln.strip() for ln in self._source_list_lines() if ln.strip()]
            for p in new_paths:
                s = str(p).strip()
                if s:
                    lines.append(s)
            dpg.set_value(self.TAG_SOURCE, "\n".join(_dedupe_path_lines(lines)))

        def _on_clear_source_paths(self) -> None:
            dpg.set_value(self.TAG_SOURCE, "")

        def resolve_work_paths(self) -> tuple[Optional[WorkPaths], Optional[str]]:
            return resolve_work_paths(
                str(dpg.get_value(self.TAG_SOURCE)),
                str(dpg.get_value(self.TAG_TARGET)),
            )

        def _process_frame(self) -> None:
            dpg.render_dearpygui_frame()
            jobs = dpg.get_callback_queue()
            dpg.run_callbacks(jobs)

        def _set_preview(self, text: str) -> None:
            dpg.set_value(self.TAG_PREVIEW, text)
            if dpg.is_dearpygui_running():
                self._process_frame()

        def _msg(self, title: str, message: str) -> None:
            self._set_preview(f"{title}\n\n{message}")

        def _gui_sections_snapshot(self) -> dict[str, bool]:
            sections: dict[str, bool] = {}
            for key, tag in self._section_tags.items():
                if dpg.does_item_exist(tag):
                    sections[key] = bool(dpg.get_value(tag))
            return sections

        def persist_paths(self) -> None:
            config_save(
                str(dpg.get_value(self.TAG_SOURCE)),
                str(dpg.get_value(self.TAG_TARGET)).strip(),
                str(dpg.get_value(self.TAG_STRIP)),
                str(dpg.get_value(self.TAG_BRACKET)).strip(),
                gui_sections=self._gui_sections_snapshot(),
            )

        def _bracket_tag_text(self) -> str:
            return sanitize_tag_text(str(dpg.get_value(self.TAG_BRACKET)))

        def _update_source_paths_after_renames(
            self, renames: list[tuple[Path, Path]]
        ) -> None:
            if not renames:
                return
            lines = self._source_list_lines()
            updated = remap_path_lines(lines, renames)
            if updated != lines:
                dpg.set_value(self.TAG_SOURCE, "\n".join(updated))
                self.persist_paths()

        def _apply_renames_gui(
            self, planned: list[tuple[Path, Path]]
        ) -> tuple[int, list[str]]:
            n, errors, done = apply_renames(planned)
            self._update_source_paths_after_renames(done)
            return n, errors

        def _work_tuple(self, work: WorkPaths):
            return work.source_root, work.target_root, work.only, work.source_library

        def _gui_preview(
            self,
            action: str,
            busy: str,
            scan,
            format_result,
            *,
            pre_check=None,
        ) -> None:
            config_record_last(action, "preview")
            if pre_check is not None:
                err = pre_check()
                if err:
                    title, msg = err
                    self._msg(title, msg)
                    return
            work, err = self.resolve_work_paths()
            if err:
                self._msg("Invalid paths", err)
                return
            self.persist_paths()
            self._set_preview(busy)
            try:
                result = scan(work)
            except OSError as e:
                self._msg("Scan failed", str(e))
                self._set_preview("")
                return
            self._set_preview(format_result(result, work))

        def _gui_apply(
            self,
            action: str,
            scan,
            get_items,
            apply_items,
            refresh_preview,
            *,
            pre_check=None,
            fail_label: str = "Renamed",
            error_title: str = "Some renames failed",
            list_missing=None,
        ) -> None:
            config_record_last(action, "apply")
            if pre_check is not None:
                err = pre_check()
                if err:
                    title, msg = err
                    self._msg(title, msg)
                    return
            work, err = self.resolve_work_paths()
            if err:
                self._msg("Invalid paths", err)
                return
            try:
                result = scan(work)
            except OSError as e:
                self._msg("Scan failed", str(e))
                return
            if list_missing is not None and list_missing(result, work):
                return
            items = get_items(result)
            if not items:
                return
            self.persist_paths()
            n, errors = apply_items(items)
            if errors:
                self._set_preview(
                    f"{error_title}\n\n{fail_label}: {n}\nFailed: {len(errors)}\n\n"
                    + "\n\n".join(errors[:10])
                )
            else:
                refresh_preview()

        def on_close(self) -> None:
            try:
                self.persist_paths()
            except OSError:
                pass

        def run(self) -> None:
            while dpg.is_dearpygui_running():
                self._process_frame()
            dpg.destroy_context()

        def _bracket_tag_required(self, when: str):
            tag = self._bracket_tag_text()
            if tag:
                return None
            return ("Tag text required", f"Enter tag text (e.g. NQ) before {when}.")

        def on_preview(self) -> None:
            self._gui_preview(
                "mark",
                "Scanning…\n",
                lambda w: scan_target(*self._work_tuple(w)[:3], source_library=w.source_library),
                lambda r, w: format_preview_mark(r, w.only),
            )

        def on_apply(self) -> None:
            self._gui_apply(
                "mark",
                lambda w: scan_target(*self._work_tuple(w)[:3], source_library=w.source_library),
                lambda r: r.planned,
                self._apply_renames_gui,
                self.on_preview,
                fail_label="Renamed",
            )

        def on_preview_title_strip(self) -> None:
            strip = str(dpg.get_value(self.TAG_STRIP))

            def scan(w):
                return scan_title_strip(
                    w.source_root,
                    w.target_root,
                    strip,
                    only=w.only,
                    source_library=w.source_library,
                )

            self._gui_preview(
                "title-strip",
                "Scanning target for title strip…\n",
                scan,
                lambda r, w: format_preview_rename(
                    r,
                    w.only,
                    "Target files to rename (strip characters from stem):\n",
                    "unchanged",
                    collision_basename_only=True,
                ),
            )

        def on_apply_title_strip(self) -> None:
            strip = str(dpg.get_value(self.TAG_STRIP))

            def scan(w):
                return scan_title_strip(
                    w.source_root,
                    w.target_root,
                    strip,
                    only=w.only,
                    source_library=w.source_library,
                )

            self._gui_apply(
                "title-strip",
                scan,
                lambda r: r.planned,
                self._apply_renames_gui,
                self.on_preview_title_strip,
            )

        def on_preview_bracket_tag(self) -> None:
            tag = self._bracket_tag_text()

            def scan(w):
                return scan_bracket_tag(
                    w.source_root,
                    w.target_root,
                    tag,
                    only=w.only,
                    source_library=w.source_library,
                )

            self._gui_preview(
                "bracket-tag",
                f'Scanning target for tag "[{tag}]"…\n',
                scan,
                lambda r, w: format_preview_rename(
                    r,
                    w.only,
                    f'Target files to rename (append " [{tag}]" at end):\n',
                    "already has tag at end / no change",
                    collision_basename_only=False,
                ),
                pre_check=lambda: self._bracket_tag_required("preview"),
            )

        def on_apply_bracket_tag(self) -> None:
            tag = self._bracket_tag_text()

            def scan(w):
                return scan_bracket_tag(
                    w.source_root,
                    w.target_root,
                    tag,
                    only=w.only,
                    source_library=w.source_library,
                )

            self._gui_apply(
                "bracket-tag",
                scan,
                lambda r: r.planned,
                self._apply_renames_gui,
                self.on_preview_bracket_tag,
                pre_check=lambda: self._bracket_tag_required("apply"),
            )

        def on_preview_bracket_tag_remove_all(self) -> None:
            self._gui_preview(
                "bracket-tag",
                "Scanning target to remove all bracket tags…\n",
                lambda w: scan_bracket_tag_remove_all(
                    w.source_root,
                    w.target_root,
                    only=w.only,
                    source_library=w.source_library,
                ),
                lambda r, w: format_preview_rename(
                    r,
                    w.only,
                    "Target files to rename (remove all […] tags):\n",
                    "no bracket tags / no change",
                    collision_basename_only=False,
                ),
            )

        def on_apply_bracket_tag_remove_all(self) -> None:
            self._gui_apply(
                "bracket-tag",
                lambda w: scan_bracket_tag_remove_all(
                    w.source_root,
                    w.target_root,
                    only=w.only,
                    source_library=w.source_library,
                ),
                lambda r: r.planned,
                self._apply_renames_gui,
                self.on_preview_bracket_tag_remove_all,
            )

        def on_preview_jpg(self) -> None:
            self._gui_preview(
                "jpg-move",
                "Scanning .jpg files…\n",
                lambda w: scan_jpg_moves(w.source_root, w.target_root, only=w.only),
                lambda r, w: format_preview_jpg(r, w.only),
            )

        def on_apply_jpg(self) -> None:
            self._gui_apply(
                "jpg-move",
                lambda w: scan_jpg_moves(w.source_root, w.target_root, only=w.only),
                lambda r: r.moves,
                apply_jpg_moves,
                self.on_preview_jpg,
                fail_label="Moved",
                error_title="Some moves failed",
            )

        def on_preview_copy_transfer(self) -> None:
            self._gui_preview(
                "copy-transfer",
                "Reading copy list…\n",
                lambda w: scan_copy_transfer(
                    w.source_root, w.target_root, only=w.only
                ),
                lambda r, w: format_preview_copy_transfer(r, w.only, w.source_root),
            )

        def on_apply_copy_transfer(self) -> None:
            def list_missing(xfer, work):
                if not xfer.list_missing:
                    return False
                self._msg(
                    "List missing",
                    f'Could not find "{COPY_TRANSFER_LIST_NAME}" under the source folder:\n\n'
                    f"{work.source_root}",
                )
                return True

            self._gui_apply(
                "copy-transfer",
                lambda w: scan_copy_transfer(
                    w.source_root, w.target_root, only=w.only
                ),
                lambda r: r.copies,
                apply_copy_transfer,
                self.on_preview_copy_transfer,
                fail_label="Copied",
                error_title="Some copies failed",
                list_missing=list_missing,
            )

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
        choices=LAST_ACTIONS,
        help="Operation to run from the command line.",
    )
    parser.add_argument(
        "--repeat",
        action="store_true",
        help="Run the last saved action and mode (preview or apply) from settings.",
    )
    parser.add_argument("--source", help="Source folder path.")
    parser.add_argument("--target", help="Target folder path.")
    parser.add_argument(
        "--strip-chars",
        default="",
        help="Characters to strip from stems (title-strip only; uses saved value if omitted).",
    )
    parser.add_argument(
        "--tag-text",
        default="",
        help='Bracket tag text (bracket-tag only; uses saved value if omitted).',
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

    if args.repeat:
        last_action, last_mode = config_load_last()
        args.action = last_action
        # --apply / --preview on the command line win (DOpus Ctrl+click passes --apply).
        if not args.apply and not args.preview:
            if last_mode == "apply":
                args.apply = True
            else:
                args.preview = True

    if not args.action:
        run_gui(
            initial_source=args.source,
            initial_target=args.target,
            initial_only_list=args.only_list,
            initial_only_files=args.only_file,
        )
        return 0

    saved_src, saved_tgt, saved_strip, saved_tag = config_load_defaults()
    tgt_s = (args.target or saved_tgt or "").strip()
    strip_chars = normalize_strip_chars(
        args.strip_chars if args.strip_chars != "" else saved_strip
    )
    tag_text = sanitize_tag_text(
        args.tag_text if args.tag_text != "" else saved_tag
    )

    # Match GUI / DOpus: --only-list or --only-file → only those paths, not saved folders too.
    src_s = build_initial_source_text(
        saved_src, args.source, args.only_list, args.only_file
    )

    work, err = resolve_work_paths(src_s, tgt_s)
    if err:
        print(err, file=sys.stderr)
        return 2

    do_apply = bool(args.apply)
    do_preview = not do_apply or args.preview
    config_record_last(args.action, "apply" if do_apply else "preview")

    try:
        result, preview_text, early = _cli_scan_and_format(
            args.action, work, strip_chars, tag_text
        )
        if early == 2:
            if args.action == "bracket-tag":
                print(
                    "Tag text required (--tag-text or saved bracket_tag_text).",
                    file=sys.stderr,
                )
            return 2
        if early is not None:
            return early
        if do_preview and preview_text:
            print(preview_text.rstrip("\n"))
        if do_apply:
            _, errors, code = _cli_apply_items(
                args.action, _cli_planned_items(args.action, result)
            )
            if errors:
                print(_cli_format_errors(errors), file=sys.stderr)
            if code:
                return code
    except OSError as e:
        print(str(e), file=sys.stderr)
        return 1

    if do_apply:
        try:
            only_lines = _only_list_lines(args.only_list, args.only_file)
            persist_source = saved_src if only_lines else src_s
            config_save(persist_source, tgt_s, strip_chars, tag_text)
        except OSError:
            pass
    return 0


def _configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except (OSError, ValueError):
                pass


if __name__ == "__main__":
    _configure_stdio_utf8()
    if len(sys.argv) > 1:
        raise SystemExit(run_cli(sys.argv[1:]))
    run_gui()
