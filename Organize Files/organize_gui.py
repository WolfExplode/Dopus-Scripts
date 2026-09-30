"""PySide6 front-end for Organize Files.

Every operation follows the same shape: Preview scans and prints what would
change; Apply scans again, performs it, then re-runs Preview to show what is
left. Scans run on a worker thread so large trees never freeze the window.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Optional

from organize_logic import (
    CHECK,
    CONFIG_DIR,
    COPY_TRANSFER_LIST_NAME,
    apply_copy_transfer,
    apply_jpg_moves,
    apply_renames,
    build_initial_source_text,
    config_load_compare_debug,
    config_load_compare_missing,
    config_load_compare_romaji,
    config_load_compare_shared,
    config_load_compare_strip_extensions,
    config_load_compare_threshold,
    config_load_defaults,
    config_load_last,
    config_record_last,
    config_save,
    format_preview_copy_transfer,
    format_preview_jpg,
    format_preview_mark,
    format_preview_rename,
    remap_path_lines,
    resolve_bracket_work_paths,
    resolve_compare_paths,
    resolve_work_paths,
    sanitize_tag_text,
    scan_bracket_tag,
    scan_bracket_tag_remove_all,
    scan_copy_transfer,
    scan_jpg_moves,
    scan_target,
    scan_title_strip,
)

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))

import uikit as ui  # noqa: E402
from uikit import icons  # noqa: E402

ACCENT = "#34D399"


@dataclass
class Report:
    ok: bool
    summary: str = ""
    text: str = ""
    renames: list = field(default_factory=list)
    refresh: bool = False


@dataclass
class Operation:
    """One Preview/Apply pair."""

    action: str                       # config_record_last key
    busy: str                         # console header while scanning
    scan: Callable[[Any], Any]        # work paths -> scan result
    fmt: Callable[[Any, Any], str]    # (result, work) -> preview text
    items: Callable[[Any], list]      # result -> planned pairs
    apply: Callable[[list], tuple]    # pairs -> (n, errors[, done])
    verb: str = "Renamed"
    source_only: bool = False         # resolve Source only (no Target needed)
    pre_check: Optional[Callable[[], Optional[str]]] = None
    missing: Optional[Callable[[Any, Any], Optional[str]]] = None


class OrganizeWindow(ui.ToolWindow):
    def __init__(self, source_text: str, target: str, strip: str, tag: str):
        self.sources = ui.PathList(
            empty_text="Drop a source folder or files here\nA folder means every file inside it",
            files_title="Select files",
            noun="path",
        )
        self.sources.set_text(source_text)
        self.target = ui.PathField(target, placeholder="Folder where marks, moves and copies land",
                                   title="Select target folder",
                                   tip="Drop a folder here. Dropping a file uses its parent folder.")
        inputs = ui.Card("Source & target",
                         "Source paths feed every operation. Target is needed by marks, JPG transfer, "
                         "copy-from-list and compare.", icons.FOLDER)
        inputs.add(self.sources, 1)
        inputs.add(ui.field("Target folder", self.target))

        console = ui.ConsolePanel("Preview", placeholder="Pick an operation, then Preview. Results show up here.")
        super().__init__(
            title="Organize Files",
            tagline="Source ↔ target workflow",
            glyph=icons.FOLDER,
            config_dir=CONFIG_DIR,
            inputs=inputs,
            console=console,
        )
        self.jobs = ui.JobHost(self, console)
        self._ops = self._operations()
        self._strip_default = strip
        self._tag_default = tag

        self.add_page("mark", "Match marks", icons.CHECK, self._build_mark(), overline="Target only",
                      on_run=lambda: self.preview("mark"))
        self.add_page("title", "Title cleanup", icons.RENAME, self._build_title(), overline="Source paths",
                      on_run=lambda: self.preview("title"))
        self.add_page("tags", "Filename tags", icons.TAG, self._build_tags(), overline="Source paths",
                      on_run=lambda: self.preview("tag"))
        self.add_page("jpg", "JPG transfer", icons.PHOTO, self._build_jpg(), overline="Source → target",
                      on_run=lambda: self.preview("jpg"))
        self.add_page("copy", "Copy from list", icons.LIST, self._build_copy(), overline="Source → target",
                      on_run=lambda: self.preview("copy"))
        self.add_page("compare", "Compare text", icons.COMPARE, self._build_compare(), overline="Two text files",
                      on_run=self.compare)
        self.add_rail_note(ui.RailNote("Tip", "Preview first. Apply always rescans, so what you see is what "
                                               "happens. Ctrl+Enter previews the open page."))
        self.on_close = self.persist

    # ── operations table ──────────────────────────────────────────────────
    def _operations(self) -> dict[str, Operation]:
        def title_scan(w):
            return scan_title_strip(w.source_root, self.strip.text(), only=w.only)

        def tag_scan(w):
            return scan_bracket_tag(w.source_root, self._tag(), only=w.only)

        def need_tag() -> Optional[str]:
            return None if self._tag() else "Enter tag text (e.g. NQ) first."

        def copy_missing(r, w) -> Optional[str]:
            if not r.list_missing:
                return None
            return f'Could not find "{COPY_TRANSFER_LIST_NAME}" under the source folder:\n\n{w.source_root}'

        return {
            "mark": Operation(
                "mark", "Scanning target tree",
                lambda w: scan_target(w.source_root, w.target_root, w.only, source_library=w.source_library),
                lambda r, w: format_preview_mark(r, w.only),
                lambda r: r.planned, apply_renames,
            ),
            "title": Operation(
                "title-strip", "Scanning source paths for title cleanup",
                title_scan,
                lambda r, w: format_preview_rename(
                    r, w.only, "Files to rename (strip characters from stem):\n", "unchanged",
                    collision_basename_only=False),
                lambda r: r.planned, apply_renames, source_only=True,
            ),
            "tag": Operation(
                "bracket-tag", "Scanning source paths for the tag",
                tag_scan,
                lambda r, w: format_preview_rename(
                    r, w.only, f'Files to rename (append " [{self._tag()}]" at end):\n',
                    "already has tag at end / no change", collision_basename_only=False),
                lambda r: r.planned, apply_renames, source_only=True, pre_check=need_tag,
            ),
            "untag": Operation(
                "bracket-tag", "Scanning source paths for bracket tags",
                lambda w: scan_bracket_tag_remove_all(w.source_root, only=w.only),
                lambda r, w: format_preview_rename(
                    r, w.only, "Files to rename (remove all […] tags):\n", "no bracket tags / no change",
                    collision_basename_only=False),
                lambda r: r.planned, apply_renames, source_only=True,
            ),
            "jpg": Operation(
                "jpg-move", "Scanning .jpg files",
                lambda w: scan_jpg_moves(w.source_root, w.target_root, only=w.only),
                lambda r, w: format_preview_jpg(r, w.only),
                lambda r: r.moves, apply_jpg_moves, verb="Moved",
            ),
            "copy": Operation(
                "copy-transfer", "Reading copy list",
                lambda w: scan_copy_transfer(w.source_root, w.target_root, only=w.only),
                lambda r, w: format_preview_copy_transfer(r, w.only, w.source_root),
                lambda r: r.copies, apply_copy_transfer, verb="Copied", missing=copy_missing,
            ),
        }

    # ── pages ──────────────────────────────────────────────────────────────
    def _pair(self, key: str, apply_text: str = "Apply", apply_tip: str = "", preview_tip: str = ""):
        pv = ui.button("Preview", "secondary", icons.FILTER, preview_tip or "Scan and list what would change.",
                       lambda: self.preview(key))
        ap = ui.button(apply_text, "primary", icons.CHECK, apply_tip, lambda: self.apply(key))
        self.jobs.lock_while_running(pv, ap)
        return pv, ap

    def _build_mark(self):
        card = ui.Card(f"{CHECK} Match marks",
                       f"Prepends {CHECK} to a target file when the same relative folder in Source holds a file "
                       "with a matching base name (source foo.mkv ↔ target foo.mp4.jpg). Nothing is moved or "
                       "deleted.", icons.CHECK)
        card.add_actions(*self._pair("mark", "Add marks", f"Add the {CHECK} prefix to matched target files."))
        return ui.page(card)

    def _build_title(self):
        self.strip = ui.line_edit(self._strip_default, "e.g. 「」",
                                  "Every character typed here is removed wherever it appears in a stem.")
        card = ui.Card("Title cleanup",
                       "Cleans source file names in place: removes the characters below, trims spaces and trailing "
                       "dots, collapses doubled extensions (video.mp4.mp4) and drops Explorer copy suffixes like "
                       '" (1)". Years such as " (2024)" are kept.', icons.RENAME)
        card.add(ui.field("Strip characters", self.strip,
                          "Optional. The other clean-ups always run."))
        card.add_actions(*self._pair("title", "Rename", "Rename source files with the cleaned titles."))
        return ui.page(card)

    def _build_tags(self):
        self.tag = ui.line_edit(self._tag_default, "e.g. NQ",
                                'Text inside the brackets. Forbidden characters become "_". NQ → "movie [NQ].mp4".')
        add = ui.Card("Append a tag", 'Adds " [tag]" to the end of each source file name. Skipped when that tag '
                                      "is already the final one; other tags stay.", icons.TAG)
        add.add(ui.field("Tag text", self.tag))
        add.add_actions(*self._pair("tag", "Add tag"))
        rm = ui.Card("Remove every tag", "Strips all […] tags from source file names, wherever they are.",
                     icons.DELETE)
        rm.add_actions(*self._pair("untag", "Remove tags"))
        return ui.page(add, rm)

    def _build_jpg(self):
        card = ui.Card("Move JPGs into target",
                       "Every .jpg under Source moves to the same relative path under Target. Folders are created "
                       "as needed; existing files are never overwritten.", icons.PHOTO)
        card.add_actions(*self._pair("jpg", "Move", "Move the listed .jpg files (removed from Source)."))
        return ui.page(card)

    def _build_copy(self):
        card = ui.Card("Copy videos from a list",
                       f'Put "{COPY_TRANSFER_LIST_NAME}" in the Source folder with one thumbnail name per line '
                       "(.mp4.jpg / .wmv.jpg). The .jpg is dropped to find each video under Source, which is then "
                       "copied to the same relative path in Target. Source is kept.", icons.LIST)
        card.add_actions(*self._pair("copy", "Copy", "Copy the listed videos into Target."))
        return ui.page(card)

    def _build_compare(self):
        thr = config_load_compare_threshold()
        self.cmp_threshold = ui.spin(thr, 0, 100, "Minimum line score. Each word must reach it against a word "
                                                  "in the other line.", suffix=" %")
        self.cmp_missing = ui.OptionRow("Missing", "missing-from-<other>.txt per side: lines only in one file.",
                                        config_load_compare_missing())
        self.cmp_shared = ui.OptionRow("Shared", "shared.txt: lines of file A that matched B, in A's order.",
                                       config_load_compare_shared())
        self.cmp_debug = ui.OptionRow("Debug", "fuzzy-matches.txt and fuzzy-mismatches.txt with scores.",
                                      config_load_compare_debug())
        self.cmp_strip_ext = ui.OptionRow("Strip extensions", "Drop a trailing audio extension (.mp3, .flac …) "
                                                              "before matching.", config_load_compare_strip_extensions())
        self.cmp_romaji = ui.OptionRow("Romaji", "Also score Japanese converted to Hepburn romaji; the higher score "
                                                 "wins.", config_load_compare_romaji())
        card = ui.Card("Compare two text files",
                       "Source paths: exactly two files, A then B, one entry per line. Exact matches first, then "
                       "word-based fuzzy matching. Reports are written to Target.", icons.COMPARE)
        card.add(ui.field("Fuzzy threshold", self.cmp_threshold))
        outputs = ui.grid_fields(self.cmp_missing, self.cmp_shared, self.cmp_debug, columns=3)
        card.add(ui.field("Reports to write", outputs))
        card.add(ui.field("Matching", ui.grid_fields(self.cmp_strip_ext, self.cmp_romaji, columns=2)))
        btn = ui.button("Compare", "primary", icons.COMPARE, "Compare and write the selected reports.", self.compare)
        self.jobs.lock_while_running(btn)
        card.add_actions(btn)
        return ui.page(card)

    # ── helpers ────────────────────────────────────────────────────────────
    def _tag(self) -> str:
        return sanitize_tag_text(self.tag.text())

    def _resolve(self, op: Operation):
        if op.source_only:
            return resolve_bracket_work_paths(self.sources.text())
        return resolve_work_paths(self.sources.text(), self.target.value())

    def persist(self) -> None:
        try:
            config_save(
                self.sources.text(),
                self.target.value(),
                self.strip.text(),
                self.tag.text().strip(),
                compare_threshold=self.cmp_threshold.value(),
                compare_missing=self.cmp_missing.isChecked(),
                compare_debug=self.cmp_debug.isChecked(),
                compare_shared=self.cmp_shared.isChecked(),
                compare_strip_extensions=self.cmp_strip_ext.isChecked(),
                compare_romaji=self.cmp_romaji.isChecked(),
            )
        except OSError:
            pass

    def _prepare(self, key: str, mode: str):
        op = self._ops[key]
        config_record_last(op.action, mode)
        if op.pre_check is not None:
            err = op.pre_check()
            if err:
                self.toast(err)
                return None, None
        work, err = self._resolve(op)
        if err:
            self.console.set_text(f"Invalid paths\n\n{err}", tone="error")
            self.console.set_status("error", "Invalid")
            return None, None
        self.persist()
        return op, work

    # ── preview / apply ────────────────────────────────────────────────────
    def preview(self, key: str) -> None:
        op, work = self._prepare(key, "preview")
        if op is None:
            return

        def job(emit):
            return Report(True, text=op.fmt(op.scan(work), work))

        self.jobs.start(f"{op.busy}…", job, on_done=self._show_report)

    def apply(self, key: str) -> None:
        op, work = self._prepare(key, "apply")
        if op is None:
            return

        def job(emit):
            result = op.scan(work)
            if op.missing is not None:
                msg = op.missing(result, work)
                if msg:
                    return Report(False, "List missing", text=msg)
            items = op.items(result)
            if not items:
                return Report(True, "Nothing to change.", text=op.fmt(result, work))
            emit(f"{op.verb} {len(items)} item(s)…", False)
            out = op.apply(items)
            n, errors = out[0], out[1]
            done = out[2] if len(out) > 2 else []
            if errors:
                text = (f"{op.verb}: {n}\nFailed: {len(errors)}\n\n" + "\n\n".join(errors[:10]))
                return Report(False, f"{op.verb} {n}, {len(errors)} failed", text=text, renames=done)
            return Report(True, f"{op.verb} {n} item(s).", renames=done, refresh=True)

        def done(report) -> None:
            if isinstance(report, Report) and report.renames:
                self._remap_sources(report.renames)
            if isinstance(report, Report) and report.refresh:
                self.preview(key)
            else:
                self._show_report(report)

        self.jobs.start(f"{op.busy} and applying…", job, on_done=done)

    def _show_report(self, report) -> None:
        if isinstance(report, Report) and report.text:
            self.console.set_text(report.text)

    def _remap_sources(self, renames) -> None:
        lines = self.sources.text().splitlines()
        updated = remap_path_lines(lines, renames)
        if updated != lines:
            self.sources.set_paths(updated)
            self.persist()

    def compare(self) -> None:
        try:
            import compare_text
        except ImportError:
            self.console.set_text('Compare needs extra packages.\n\nRun: pip install -r "Organize Files/requirements.txt"',
                                  tone="error")
            return
        file_a, file_b, out_dir, err = resolve_compare_paths(self.sources.text(), self.target.value())
        if err:
            self.console.set_text(f"Text compare\n\n{err}", tone="error")
            return
        missing, shared, debug = self.cmp_missing.isChecked(), self.cmp_shared.isChecked(), self.cmp_debug.isChecked()
        if not (missing or shared or debug):
            self.toast("Choose at least one report to write")
            return
        threshold = self.cmp_threshold.value()
        strip_ext, romaji = self.cmp_strip_ext.isChecked(), self.cmp_romaji.isChecked()
        self.persist()

        def job(emit):
            result = compare_text.run_compare(
                file_a, file_b, output_dir=out_dir, threshold=threshold, write_missing=missing,
                write_shared=shared, debug=debug, strip_extensions=strip_ext, romaji_compare=romaji,
            )
            return Report(True, text=compare_text.format_compare_report(result))

        self.jobs.start(f"Comparing {file_a.name} ↔ {file_b.name}…", job, on_done=self._show_report)


PAGE_FOR_ACTION = {
    "mark": "mark", "title-strip": "title", "bracket-tag": "tags", "jpg-move": "jpg", "copy-transfer": "copy",
}


def run_gui(
    initial_source: Optional[str] = None,
    initial_target: Optional[str] = None,
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    ui.create_app("OrganizeFiles", ACCENT)
    s_default, t_default, strip_default, tag_default = config_load_defaults()
    if initial_target and str(initial_target).strip():
        t_default = str(initial_target).strip()
    source_text = build_initial_source_text(s_default, initial_source, initial_only_list, initial_only_files)
    win = OrganizeWindow(source_text, t_default, strip_default, tag_default)
    last_action, _ = config_load_last()
    win.restore_ui(fallback_page=PAGE_FOR_ACTION.get(last_action, ""))
    win.run_app()
