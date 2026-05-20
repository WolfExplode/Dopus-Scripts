"""Dear PyGui front-end for Organize Files."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional

from organize_logic import *

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
        TAG_COMPARE_THRESHOLD = "compare_threshold_input"
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
        TIP_COMPARE = (
            "Compare two text lists (Artist - Title per line): a streaming playlist export vs a local "
            "file list. Uses exact keys plus RapidFuzz fuzzy matching (ratio, partial, character coverage).\n"
            "Source paths: two list files — first line playlist export (e.g. www.txt), second line local "
            "file list (e.g. Music.txt). Target folder: report files "
            "(missing-from-library.txt, extra-not-in-playlist.txt, fuzzy-matches.txt)."
        )
        TIP_COMPARE_THRESHOLD = (
            "Minimum match score (0–100). A pair matches if any fuzzy metric reaches this value."
        )
        TIP_PREVIEW_PANEL = "Output from the last Preview or Apply scan."

        def __init__(self) -> None:
            s_default, t_default, strip_default, tag_default = config_load_defaults()
            cmp_thr = config_load_compare_threshold()
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
                self._build_main_layout(
                    s_default, t_default, strip_default, tag_default, cmp_thr
                )
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
            cmp_thr_default: int,
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

                    hdr_compare = self._section(
                        "Compare playlists", self.TIP_COMPARE, "compare"
                    )
                    with dpg.group(parent=hdr_compare):
                        dpg.add_text("Fuzzy threshold %", color=(150, 158, 175))
                        thr_input = dpg.add_input_int(
                            tag=self.TAG_COMPARE_THRESHOLD,
                            default_value=cmp_thr_default,
                            min_value=0,
                            max_value=100,
                            width=80,
                        )
                        self._hover_tip(thr_input, self.TIP_COMPARE_THRESHOLD)
                        dpg.add_spacer(height=4)
                        btn_compare = dpg.add_button(
                            label="Compare",
                            callback=self.on_compare_playlists,
                            width=-1,
                        )
                        dpg.bind_item_theme(btn_compare, self._theme_apply)
                        self._hover_tip(
                            btn_compare,
                            "Run compare and show results; writes report files to the target folder.",
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

        def _paths_from_drop(self, data: object) -> list[str]:
            if data is None:
                return []
            if isinstance(data, str):
                lines = [ln.strip() for ln in data.splitlines() if ln.strip()]
                return lines if lines else ([data.strip()] if data.strip() else [])
            if isinstance(data, list):
                return [str(p).strip() for p in data if str(p).strip()]
            s = str(data).strip()
            return [s] if s else []

        def _drop_hover_tags(self) -> tuple[str, ...]:
            return (self.TAG_SOURCE, self.TAG_TARGET)

        def _item_under_mouse(self, tag: str) -> bool:
            if not dpg.does_item_exist(tag):
                return False
            if dpg.is_item_hovered(tag):
                return True
            try:
                mx, my = dpg.get_mouse_pos(local=False)
                rmin = dpg.get_item_rect_min(tag)
                rmax = dpg.get_item_rect_max(tag)
                return rmin[0] <= mx <= rmax[0] and rmin[1] <= my <= rmax[1]
            except Exception:
                return False

        def _drop_target_tag(self) -> str | None:
            for tag in self._drop_hover_tags():
                if self._item_under_mouse(tag):
                    return tag
            return None

        def _apply_target_drop(self, paths: list[str]) -> None:
            for raw in paths:
                folder = self._path_as_target_folder(raw)
                if folder:
                    dpg.set_value(self.TAG_TARGET, folder)
                    return

        def _clear_drop_hover_themes(self) -> None:
            for tag in self._drop_hover_tags():
                if dpg.does_item_exist(tag):
                    dpg.bind_item_theme(tag, None)

        def _bind_drop_hover(self, active_tag: str) -> None:
            for tag in self._drop_hover_tags():
                if dpg.does_item_exist(tag):
                    dpg.bind_item_theme(
                        tag, "theme_drop_hover" if tag == active_tag else None
                    )

        def _on_os_drop(self, data, keys) -> None:
            paths = self._paths_from_drop(data)
            if not paths:
                return
            target = self._drop_target_tag()
            if target == self.TAG_TARGET:
                self._apply_target_drop(paths)
            elif target == self.TAG_SOURCE:
                self._merge_source_paths(paths)
            self._clear_drop_hover_themes()

        def _on_os_drag_over(self, keys) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd

            active = self._drop_target_tag()
            if active:
                self._bind_drop_hover(active)
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
            dpg.set_value(self.TAG_SOURCE, "\n".join(dedupe_path_lines(lines)))

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
                compare_threshold=int(dpg.get_value(self.TAG_COMPARE_THRESHOLD)),
            )

        def on_compare_playlists(self) -> None:
            try:
                import compare_playlists
            except ImportError:
                self._msg(
                    "Playlist compare",
                    "rapidfuzz is not installed.\n\n"
                    'Run: pip install -r "Organize Files/requirements.txt"',
                )
                return
            playlist_p, library_p, output_p, err = resolve_compare_paths(
                str(dpg.get_value(self.TAG_SOURCE)),
                str(dpg.get_value(self.TAG_TARGET)),
            )
            if err:
                self._msg("Playlist compare", err)
                return
            threshold = int(dpg.get_value(self.TAG_COMPARE_THRESHOLD))
            self.persist_paths()
            self._set_preview("Comparing playlists…\n")
            result = compare_playlists.run_compare(
                playlist_p,
                library_p,
                output_dir=output_p,
                threshold=threshold,
            )
            self._set_preview(compare_playlists.format_compare_report(result))

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
