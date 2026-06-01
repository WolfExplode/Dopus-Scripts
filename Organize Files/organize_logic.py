"""Core file-organize operations (scan, apply, config, previews, CLI helpers)."""

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

# gallery-dl / editor artifacts: ``-cut-merged-<digits>`` at end of stem.
_TITLE_STRIP_CUT_MERGED_SUFFIX = re.compile(r"-cut-merged-\d+\Z")

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
    "compare": False,
}
GUI_SECTION_KEYS = tuple(GUI_SECTION_DEFAULTS)


def normalize_strip_chars(strip_chars: str) -> str:
    """NFC-normalize so pasted CJK / fullwidth brackets match filenames on disk."""
    return unicodedata.normalize("NFC", strip_chars or "")


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


_COMPARE_THRESHOLD_DEFAULT = 80


def config_load_compare_threshold() -> int:
    """Saved fuzzy threshold for text compare."""
    default = _COMPARE_THRESHOLD_DEFAULT
    try:
        from compare_text import FUZZY_THRESHOLD

        default = FUZZY_THRESHOLD
    except ImportError:
        pass
    try:
        data = config_read()
        raw_thr = data.get("compare_threshold")
        try:
            threshold = int(raw_thr) if raw_thr is not None else default
        except (TypeError, ValueError):
            threshold = default
        return max(0, min(100, threshold))
    except OSError:
        return default


def config_load_compare_debug() -> bool:
    """Whether text compare writes fuzzy-matches.txt and fuzzy-mismatches.txt."""
    try:
        return bool(config_read().get("compare_debug"))
    except OSError:
        return False


def config_load_compare_missing() -> bool:
    """Whether text compare writes missing-from-*.txt per side."""
    try:
        raw = config_read().get("compare_missing")
        return True if raw is None else bool(raw)
    except OSError:
        return True


def config_load_compare_shared() -> bool:
    """Whether text compare writes shared.txt (lines in A that matched B)."""
    try:
        raw = config_read().get("compare_shared")
        return True if raw is None else bool(raw)
    except OSError:
        return True


def config_load_compare_strip_extensions() -> bool:
    """Whether text compare strips audio extensions before matching."""
    try:
        raw = config_read().get("compare_strip_extensions")
        return True if raw is None else bool(raw)
    except OSError:
        return True


def config_load_compare_romaji() -> bool:
    """Whether text compare also scores Japanese lines converted to romaji."""
    try:
        raw = config_read().get("compare_romaji")
        return True if raw is None else bool(raw)
    except OSError:
        return True


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
    *,
    compare_threshold: Optional[int] = None,
    compare_debug: Optional[bool] = None,
    compare_missing: Optional[bool] = None,
    compare_shared: Optional[bool] = None,
    compare_strip_extensions: Optional[bool] = None,
    compare_romaji: Optional[bool] = None,
) -> None:
    data = config_read()
    data["source"] = source
    data["target"] = target
    data["strip_title_chars"] = strip_title_chars
    data["bracket_tag_text"] = bracket_tag_text
    if compare_threshold is not None:
        data["compare_threshold"] = max(0, min(100, int(compare_threshold)))
    if compare_debug is not None:
        data["compare_debug"] = bool(compare_debug)
    if compare_missing is not None:
        data["compare_missing"] = bool(compare_missing)
    if compare_shared is not None:
        data["compare_shared"] = bool(compare_shared)
    if compare_strip_extensions is not None:
        data["compare_strip_extensions"] = bool(compare_strip_extensions)
    if compare_romaji is not None:
        data["compare_romaji"] = bool(compare_romaji)
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


def _scan_input_renames(
    root: Path,
    new_name_for: Callable[[Path], Optional[str]],
    only: Optional[set[Path]] = None,
) -> RenameScan:
    """Plan in-place renames for files under Source paths (folder tree or explicit file list)."""
    planned: list[tuple[Path, Path]] = []
    skipped_collision: list[tuple[Path, Path]] = []
    skipped_no_change = 0
    planned_dest_keys: set[str] = set()

    for path in iter_files_in_tree(root, only):
        new_name = new_name_for(path)
        if new_name is None:
            skipped_no_change += 1
            continue
        new_path = path.with_name(new_name)
        _try_plan_rename(path, new_path, planned, skipped_collision, planned_dest_keys)

    return RenameScan(
        planned=planned,
        skipped_no_change=skipped_no_change,
        skipped_collision=skipped_collision,
        skipped_under_source=0,
    )


def scan_bracket_tag(
    source_root: Path,
    tag_text: str,
    only: Optional[set[Path]] = None,
) -> RenameScan:
    """Plan renames: append `` [tag_text]`` at end of each source-path filename stem."""
    tag_inner = sanitize_tag_text(tag_text)
    if not tag_inner:
        return _EMPTY_RENAME_SCAN
    return _scan_input_renames(
        source_root,
        lambda p: bracket_tag_new_name(p, tag_inner),
        only=only,
    )


def scan_bracket_tag_remove_all(
    source_root: Path,
    only: Optional[set[Path]] = None,
) -> RenameScan:
    """Plan renames: strip every ``[...]`` tag from source-path filenames."""
    return _scan_input_renames(source_root, bracket_tag_remove_all_name, only=only)


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
    new_stem = _TITLE_STRIP_CUT_MERGED_SUFFIX.sub("", new_stem)
    new_stem = new_stem.rstrip().rstrip(".")
    if not new_stem:
        return None
    new_name = new_stem + suffix
    if new_name == original:
        return None
    return new_name


def scan_title_strip(
    source_root: Path,
    strip_chars: str,
    only: Optional[set[Path]] = None,
) -> RenameScan:
    """Plan in-place renames under source paths (stem cleanup). No target folder."""
    return _scan_input_renames(
        source_root,
        lambda p: transform_title_filename(p.name, strip_chars),
        only=only,
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
    With two separate library trees (source vs target), a lone folder line uses the full
    source and target trees (``only`` is None). Use explicit file lines or --only-list to
    limit which paths are processed. When the source folder overlaps the target tree,
    ``only`` is also None.
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


@dataclass
class BracketWorkPaths:
    source_root: Path
    only: Optional[set[Path]]


def resolve_bracket_work_paths(
    source_text: str,
) -> tuple[Optional[BracketWorkPaths], Optional[str]]:
    source_root, _source_library, only, err = parse_source_input_lines(
        source_text.splitlines()
    )
    if err:
        return None, err
    if source_root is None:
        return None, "Add a source folder or at least one file path."
    root = source_root.resolve()
    if only is None and not root.is_dir():
        return None, f"Source is not a directory:\n{root}"
    return BracketWorkPaths(root, only), None


def resolve_compare_paths(
    source_text: str, target_text: str
) -> tuple[Optional[Path], Optional[Path], Optional[Path], Optional[str]]:
    """
    Text compare: exactly two files in Source (file A, then file B).
    Report txt files are written under Target.
    """
    tgt_s = target_text.strip()
    if not tgt_s:
        return None, None, None, "Target folder is not set."
    tgt = Path(tgt_s)
    if not tgt.is_dir():
        return None, None, None, f"Target is not a directory:\n{tgt}"

    files: list[Path] = []
    for s in source_text.splitlines():
        s = s.strip()
        if not s:
            continue
        p = Path(s)
        if p.is_dir():
            return (
                None,
                None,
                None,
                "Compare needs two text files in Source paths (not folders).\n"
                "First line = file A, second line = file B.",
            )
        if not p.is_file():
            return None, None, None, f"Path not found:\n{p}"
        files.append(p.resolve())

    if len(files) != 2:
        return (
            None,
            None,
            None,
            "Compare needs exactly two file paths in Source paths:\n"
            "file A, then file B (one path per line).",
        )

    return files[0], files[1], tgt.resolve(), None


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


def dedupe_path_lines(lines: list[str]) -> list[str]:
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
    from_dopus = dedupe_path_lines(
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
    return "\n".join(dedupe_path_lines(lines))


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
    work: WorkPaths | BracketWorkPaths,
    strip_chars: str,
    tag_text: str,
) -> tuple[object, str, Optional[int]]:
    """Scan for CLI action. Returns (result, preview_text, early_exit_code)."""
    src_r = work.source_root
    only = work.only
    if isinstance(work, BracketWorkPaths):
        tgt_r = src_r
        src_lib = None
    else:
        tgt_r = work.target_root
        src_lib = work.source_library
    if action == "mark":
        result = scan_target(src_r, tgt_r, only=only, source_library=src_lib)
        return result, format_preview_mark(result, only), None
    if action == "title-strip":
        result = scan_title_strip(src_r, strip_chars, only=only)
        return (
            result,
            format_preview_rename(
                result,
                only,
                "Files to rename (strip characters from stem):\n",
                "unchanged",
                collision_basename_only=False,
            ),
            None,
        )
    if action == "bracket-tag":
        if not tag_text:
            return None, "", 2
        result = scan_bracket_tag(src_r, tag_text, only=only)
        return (
            result,
            format_preview_rename(
                result,
                only,
                f'Files to rename (append " [{tag_text}]" at end):\n',
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


def _cli_format_errors(errors: list[str], limit: int = 10) -> str:
    if not errors:
        return ""
    return "\n\n".join(errors[:limit]) + (
        f"\n\n… and {len(errors) - limit} more." if len(errors) > limit else ""
    )


def run_cli(argv: list[str]) -> int:
    import argparse

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
        if not args.apply and not args.preview:
            if last_mode == "apply":
                args.apply = True
            else:
                args.preview = True

    if not args.action:
        from organize_gui import run_gui

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

    src_s = build_initial_source_text(
        saved_src, args.source, args.only_list, args.only_file
    )

    if args.action in ("bracket-tag", "title-strip"):
        work, err = resolve_bracket_work_paths(src_s)
    else:
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

