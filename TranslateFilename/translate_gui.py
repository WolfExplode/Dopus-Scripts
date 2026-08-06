"""Dear PyGui front-end for Translate Filename — batch translate + rename."""

from __future__ import annotations

import os
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from translate_logic import (
    Settings,
    build_initial_inputs_text,
    build_output_filename,
    config_load_settings,
    config_save_settings,
    dedupe_lines,
    forget_all_history,
    rename_file_apply,
    translate_name,
    untranslate_file,
)


def _windows_fonts_dir() -> Path:
    return Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"


def _pick_unicode_ui_font() -> Optional[Path]:
    for name in (
        "NotoSansSC-VF.ttf", "msyh.ttc", "msyhbd.ttc", "simsun.ttc",
        "mingliu.ttc", "msjh.ttc", "segoeui.ttf",
    ):
        path = _windows_fonts_dir() / name
        if path.is_file():
            return path
    return None


# Source file names can arrive in almost any script. DPG (2.x) renders each
# widget with exactly one font and no longer supports merging glyphs from
# multiple font files into one (add_font only accepts the font registry as a
# parent now). So instead of merging, each widget's font is picked dynamically
# — whichever script font actually has glyphs for that widget's current text —
# and rebound whenever the text changes.
SCRIPT_FONT_FILES: dict[str, str] = {
    "hangul": "malgun.ttf",
    "kana": "msgothic.ttc",
    "han": "simsun.ttc",
    "arabic": "NotoSansArabic-Regular.ttf",
    "hebrew": "NotoSansHebrew-Regular.ttf",
}


def _classify_char(ch: str) -> Optional[str]:
    cp = ord(ch)
    if 0xAC00 <= cp <= 0xD7A3 or 0x1100 <= cp <= 0x11FF or 0x3130 <= cp <= 0x318F \
            or 0xA960 <= cp <= 0xA97F or 0xD7B0 <= cp <= 0xD7FF:
        return "hangul"
    if 0x3040 <= cp <= 0x30FF or 0x31F0 <= cp <= 0x31FF or 0xFF66 <= cp <= 0xFF9F:
        return "kana"
    if 0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or 0x3000 <= cp <= 0x303F \
            or 0xFF00 <= cp <= 0xFFEF or 0x20000 <= cp <= 0x2A6DF:
        return "han"
    if 0x0600 <= cp <= 0x06FF or 0x0750 <= cp <= 0x077F:
        return "arabic"
    if 0x0590 <= cp <= 0x05FF:
        return "hebrew"
    return None


def _majority_script(text: str) -> Optional[str]:
    """Most frequent non-Latin script bucket in text (None = plain/Latin).

    Editable widgets (Inputs box, translated/text-only fields) get exactly one
    bound font each — DPG has no per-character multi-font rendering for those.
    No installed font covers every script at once (e.g. Malgun Gothic has full
    Hangul but only partial Han, SimSun is the reverse), so a line mixing two
    non-Latin scripts will render its minority script as tofu in these widgets.
    Picking the majority script minimizes how much of the line is affected.
    """
    counts: dict[str, int] = {}
    for ch in text or "":
        bucket = _classify_char(ch)
        if bucket:
            counts[bucket] = counts.get(bucket, 0) + 1
    if not counts:
        return None
    return max(counts.items(), key=lambda kv: kv[1])[0]


def _split_script_runs(text: str) -> list[tuple[str, Optional[str]]]:
    """Split text into (run_text, script_bucket) pairs, for read-only display.

    Used for the per-entry "original name" row, which isn't a native editable
    control — it's built from separate dpg.add_text items, one per run, each
    bound to the font that actually has those glyphs. This is how every
    script in a mixed-script name renders correctly, unlike the single-font
    editable widgets above. Characters with no specific script (ASCII,
    digits, punctuation, spaces) stick to whichever script is currently
    active, so e.g. "[159]_bpm.png" after a Korean word isn't split off.
    """
    runs: list[tuple[str, Optional[str]]] = []
    current_bucket: Optional[str] = None
    current_chars: list[str] = []
    for ch in text:
        bucket = _classify_char(ch)
        if bucket is None:
            bucket = current_bucket
        if bucket != current_bucket and current_chars:
            runs.append(("".join(current_chars), current_bucket))
            current_chars = []
        current_bucket = bucket
        current_chars.append(ch)
    if current_chars:
        runs.append(("".join(current_chars), current_bucket))
    return runs


def _pick_native_files(title: str) -> list[str]:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    paths = filedialog.askopenfilenames(title=title, parent=root)
    root.destroy()
    return list(paths) if paths else []


def _confirm(title: str, message: str) -> bool:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    result = messagebox.askyesno(title, message, parent=root)
    root.destroy()
    return bool(result)


@dataclass
class EntryState:
    raw_line: str
    path: Optional[Path]
    display_original: str
    status: str = "pending"  # "pending" | "ok" | "error"
    translation: str = ""
    text_only: str = ""
    error: str = ""
    warning: str = ""
    apply_error: str = ""
    applied: bool = False
    tag_translated: object = field(default=None)
    tag_textonly: object = field(default=None)


def _resolve_entry_source(line: str) -> tuple[Optional[Path], str]:
    p = Path(line)
    if p.is_file() or p.is_dir():
        return p, p.name
    return None, line


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    import dearpygui.dearpygui as dpg

    class App:
        TAG_INPUTS = "inputs_text"
        TAG_ENTRIES_PANEL = "entries_panel"
        TAG_API_KEY = "api_key_input"
        TAG_MODEL = "model_input"
        TAG_AUTO_RENAME = "auto_rename_check"
        TAG_APPEND_MODE = "append_mode_check"
        TAG_STATUS = "status_text"
        TAG_TRANSLATE_BTN = "translate_btn"
        TAG_APPLY_BTN = "apply_btn"
        TAG_UNTRANSLATE_BTN = "untranslate_btn"

        WINDOW_MIN_HEIGHT = 460
        WINDOW_MAX_HEIGHT = 900

        def __init__(self) -> None:
            self.settings = config_load_settings()
            inputs_default = build_initial_inputs_text(
                self.settings.inputs_text, initial_only_list, initial_only_files
            )
            self._entries: list[EntryState] = []
            self._job_thread: threading.Thread | None = None
            self._job_reported = True
            self._default_font: int | str | None = None
            self._script_fonts: dict[str, int | str] = {}

            dpg.create_context()
            self._init_os_drag_drop()
            self._build_theme()

            with dpg.window(tag="primary_window", label="Translate Filename", no_title_bar=True):
                self._build_layout(inputs_default)

            self._build_fonts()
            dpg.create_viewport(
                title="Translate Filename",
                width=900,
                height=640,
                min_width=680,
                min_height=self.WINDOW_MIN_HEIGHT,
            )
            self._register_os_drag_drop_handlers()
            dpg.setup_dearpygui()
            dpg.show_viewport()
            dpg.set_primary_window("primary_window", True)
            dpg.set_exit_callback(self.on_close)

            self._sync_entries_from_input()

        # -- theming -----------------------------------------------------

        def _build_theme(self) -> None:
            accent = (72, 168, 190)
            apply_bg = (52, 128, 108)
            apply_h = (68, 158, 132)

            with dpg.theme(tag="app_theme"):
                with dpg.theme_component(dpg.mvAll):
                    dpg.add_theme_color(dpg.mvThemeCol_WindowBg, (16, 18, 24))
                    dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (22, 25, 32))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (34, 38, 50))
                    dpg.add_theme_color(dpg.mvThemeCol_Button, (48, 54, 70))
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (60, 68, 88))
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, accent)
                    dpg.add_theme_color(dpg.mvThemeCol_Text, (228, 232, 240))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (52, 58, 74))
                    dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 10)
                    dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6)
                    dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 14, 14)
                    dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 8)

            with dpg.theme(tag="theme_apply"):
                with dpg.theme_component(dpg.mvButton):
                    dpg.add_theme_color(dpg.mvThemeCol_Button, apply_bg)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, apply_h)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (80, 178, 150))

            with dpg.theme(tag="theme_tight_box"):
                with dpg.theme_component(dpg.mvChildWindow):
                    dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 6, 4)
                    dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 0, 0)
                    dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (30, 40, 48))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (52, 88, 100))

            with dpg.theme(tag="theme_translated_input"):
                with dpg.theme_component(dpg.mvInputText):
                    dpg.add_theme_color(dpg.mvThemeCol_Text, (150, 220, 160))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (28, 42, 34))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (60, 110, 80))

            with dpg.theme(tag="theme_textonly_input"):
                with dpg.theme_component(dpg.mvInputText):
                    dpg.add_theme_color(dpg.mvThemeCol_Text, (230, 190, 120))
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (44, 38, 26))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (110, 90, 50))

            with dpg.theme(tag="theme_entry_row"):
                with dpg.theme_component(dpg.mvAll):
                    dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 0, 3)
                    dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 6, 3)

            with dpg.theme(tag="theme_drop_hover"):
                with dpg.theme_component(dpg.mvInputText):
                    dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (40, 68, 82))
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (72, 168, 190))

            dpg.bind_theme("app_theme")

        def _build_fonts(self) -> None:
            fonts_dir = _windows_fonts_dir()
            base_path = fonts_dir / "segoeui.ttf"
            if not base_path.is_file():
                base_path = _pick_unicode_ui_font()

            with dpg.font_registry():
                if base_path:
                    self._default_font = dpg.add_font(str(base_path), 14)
                    dpg.bind_font(self._default_font)
                    title_font = dpg.add_font(str(base_path), 18)
                    dpg.bind_item_font("title_main", title_font)

                for bucket, filename in SCRIPT_FONT_FILES.items():
                    path = fonts_dir / filename
                    if path.is_file():
                        self._script_fonts[bucket] = dpg.add_font(str(path), 14)

        def _font_for_text(self, text: str) -> int | str | None:
            bucket = _majority_script(text)
            if bucket:
                font = self._script_fonts.get(bucket)
                if font is not None:
                    return font
            return self._default_font

        def _wrap_width(self) -> int:
            if dpg.does_item_exist(self.TAG_ENTRIES_PANEL):
                w = dpg.get_item_rect_size(self.TAG_ENTRIES_PANEL)[0]
                if w > 0:
                    return max(120, int(w) - 26)
            return 800

        def _render_multiscript_text(
            self, group_tag, text: str, color=(150, 190, 230), wrap: Optional[int] = None
        ) -> None:
            if not dpg.does_item_exist(group_tag):
                return
            dpg.delete_item(group_tag, children_only=True)
            display_text = text if text else " "
            runs = _split_script_runs(display_text)
            max_width = wrap if wrap is not None else self._wrap_width()

            for run_text, bucket in runs:
                item = dpg.add_text(run_text, parent=group_tag, color=color, wrap=max_width)
                font = self._script_fonts.get(bucket) if bucket else None
                if font is not None:
                    dpg.bind_item_font(item, font)

        def _bind_font_for_text(self, item, text: str) -> None:
            font = self._font_for_text(text)
            if font is not None and dpg.does_item_exist(item):
                dpg.bind_item_font(item, font)

        def _refresh_inputs_font(self) -> None:
            if dpg.does_item_exist(self.TAG_INPUTS):
                self._bind_font_for_text(self.TAG_INPUTS, str(dpg.get_value(self.TAG_INPUTS)))

        def _hover_tip(self, parent, text: str) -> None:
            with dpg.tooltip(parent, delay=0.4):
                dpg.add_text(text, wrap=380, color=(200, 208, 220))

        # -- layout --------------------------------------------------------

        def _build_layout(self, inputs_default: str) -> None:
            dpg.add_text("Translate Filename", tag="title_main", color=(120, 200, 220))
            dpg.add_text("Batch translate file names to English via DeepSeek", color=(130, 138, 155))
            dpg.add_spacer(height=8)

            dpg.add_text("Inputs", color=(150, 158, 175))
            inputs_box = dpg.add_input_text(
                tag=self.TAG_INPUTS,
                default_value=inputs_default,
                multiline=True,
                width=-1,
                height=90,
                tab_input=False,
            )
            self._hover_tip(
                inputs_box,
                "One file path (or plain text) per line. Drag files here, or edit directly — "
                "each line becomes an entry below once you click away.",
            )
            with dpg.group(horizontal=True):
                add_btn = dpg.add_button(label="Browse…", callback=self._browse_add_files)
                clear_btn = dpg.add_button(label="Clear", callback=self._clear_inputs)
            self._hover_tip(add_btn, "Add files to the input list.")
            self._hover_tip(clear_btn, "Clear the input list.")
            dpg.add_spacer(height=8)

            dpg.add_child_window(width=-1, border=True, auto_resize_y=True, tag=self.TAG_ENTRIES_PANEL)
            dpg.add_spacer(height=8)

            with dpg.group(horizontal=True):
                btn_t = dpg.add_button(label="Translate", tag=self.TAG_TRANSLATE_BTN, callback=self._on_translate, width=140)
                dpg.bind_item_theme(btn_t, "theme_apply")
                btn_a = dpg.add_button(label="Rename", tag=self.TAG_APPLY_BTN, callback=self._on_apply, width=140)
                btn_u = dpg.add_button(label="Rename to Original", tag=self.TAG_UNTRANSLATE_BTN, callback=self._on_untranslate, width=140)
            self._hover_tip(btn_t, "Translate every new/changed entry via DeepSeek.")
            self._hover_tip(btn_a, "Rename every entry's file on disk using its current translation.")
            self._hover_tip(btn_u, "Revert every entry's file back to its original name.")
            dpg.add_spacer(height=8)

            with dpg.collapsing_header(label="Settings", default_open=not self.settings.api_key):
                dpg.add_text("DeepSeek API key", color=(150, 158, 175))
                dpg.add_input_text(
                    tag=self.TAG_API_KEY,
                    default_value=self.settings.api_key,
                    password=True,
                    width=-1,
                )
                dpg.add_text("Model", color=(150, 158, 175))
                dpg.add_input_text(
                    tag=self.TAG_MODEL,
                    default_value=self.settings.model,
                    width=-1,
                )
                save_btn = dpg.add_button(label="Save settings", callback=self._on_save_settings, width=140)
                self._hover_tip(save_btn, "Saved to %APPDATA%\\TranslateFilename\\settings.json")
                dpg.add_spacer(height=6)

                dpg.add_checkbox(
                    tag=self.TAG_AUTO_RENAME,
                    label="Auto-rename after Translate",
                    default_value=self.settings.auto_rename,
                    callback=self._on_toggle_setting,
                )
                self._hover_tip(
                    self.TAG_AUTO_RENAME,
                    "When on, clicking Translate also renames each file immediately.\n"
                    "Ctrl+click from Directory Opus always auto-renames regardless of this toggle.",
                )
                dpg.add_checkbox(
                    tag=self.TAG_APPEND_MODE,
                    label="Append mode (keep original name, append translated text)",
                    default_value=self.settings.append_mode,
                    callback=self._on_toggle_setting,
                )
                self._hover_tip(
                    self.TAG_APPEND_MODE,
                    "Output becomes: original name + \" \" + text-only translated phrase + extension,\n"
                    "instead of replacing the name with the full translation.",
                )
                dpg.add_spacer(height=6)
                forget_btn = dpg.add_button(label="Forget all translation history", callback=self._on_forget_all, width=220)
                self._hover_tip(forget_btn, "Deletes all stored original-name history used by Untranslate.")

            dpg.add_spacer(height=8)
            dpg.add_text("", tag=self.TAG_STATUS, wrap=860, color=(150, 158, 175))

        # -- entry list management ------------------------------------------

        def _sync_widget_values_into_state(self) -> None:
            for e in self._entries:
                if e.tag_translated is not None and dpg.does_item_exist(e.tag_translated):
                    e.translation = str(dpg.get_value(e.tag_translated))
                if e.tag_textonly is not None and dpg.does_item_exist(e.tag_textonly):
                    e.text_only = str(dpg.get_value(e.tag_textonly))

        def _sync_entries_from_input(self) -> None:
            self._sync_widget_values_into_state()
            lines = dedupe_lines(
                [ln.strip() for ln in str(dpg.get_value(self.TAG_INPUTS)).splitlines() if ln.strip()]
            )
            old_by_line = {e.raw_line: e for e in self._entries}
            new_entries: list[EntryState] = []
            for ln in lines:
                existing = old_by_line.get(ln)
                if existing is not None:
                    new_entries.append(existing)
                else:
                    path, display = _resolve_entry_source(ln)
                    new_entries.append(EntryState(raw_line=ln, path=path, display_original=display))
            self._entries = new_entries
            self._render_entries()
            self._refresh_inputs_font()

        def _append_mode(self) -> bool:
            return bool(dpg.get_value(self.TAG_APPEND_MODE)) if dpg.does_item_exist(self.TAG_APPEND_MODE) else self.settings.append_mode

        def _render_entries(self) -> None:
            dpg.delete_item(self.TAG_ENTRIES_PANEL, children_only=True)
            append_mode = self._append_mode()
            if not self._entries:
                dpg.add_text(
                    "No inputs yet. Drag files onto the box above, or use Browse….",
                    parent=self.TAG_ENTRIES_PANEL,
                    color=(130, 138, 155),
                )
                return
            for e in self._entries:
                self._render_entry(e, append_mode)

        def _render_entry(self, e: EntryState, append_mode: bool) -> None:
            parent = self.TAG_ENTRIES_PANEL
            entry_group = dpg.generate_uuid()
            with dpg.group(parent=parent, tag=entry_group):
                dpg.bind_item_theme(entry_group, "theme_entry_row")
                orig_box = dpg.generate_uuid()
                orig_group = dpg.generate_uuid()
                with dpg.child_window(width=-1, border=True, auto_resize_y=True, tag=orig_box):
                    dpg.bind_item_theme(orig_box, "theme_tight_box")
                    with dpg.group(tag=orig_group):
                        pass
                self._render_multiscript_text(orig_group, e.display_original)

                e.tag_translated = dpg.generate_uuid()
                dpg.add_input_text(tag=e.tag_translated, default_value=e.translation, width=-1)
                dpg.bind_item_theme(e.tag_translated, "theme_translated_input")
                self._bind_font_for_text(e.tag_translated, e.translation)

                e.tag_textonly = dpg.generate_uuid()
                dpg.add_input_text(tag=e.tag_textonly, default_value=e.text_only, width=-1)
                dpg.bind_item_theme(e.tag_textonly, "theme_textonly_input")
                self._bind_font_for_text(e.tag_textonly, e.text_only)

                notes: list[str] = []
                if e.status == "error":
                    notes.append(("error", f"Translate error: {e.error}"))
                if e.apply_error:
                    notes.append(("error", e.apply_error))
                if e.status == "ok":
                    is_dir = e.path.is_dir() if e.path is not None else False
                    _, warn = build_output_filename(
                        e.display_original, e.translation, e.text_only, append_mode, is_dir=is_dir
                    )
                    if warn:
                        notes.append(("warn", warn))
                if e.applied:
                    notes.append(("info", f"Renamed \u2192 {e.path.name if e.path else ''}"))
                if e.path is None:
                    notes.append(("info", "No backing file — translate-only (Apply/Untranslate skip this line)."))
                for kind, text in notes:
                    color = {"error": (220, 100, 100), "warn": (230, 160, 90), "info": (130, 160, 190)}[kind]
                    note_group = dpg.generate_uuid()
                    dpg.add_group(parent=entry_group, tag=note_group)
                    self._render_multiscript_text(note_group, text, color=color, wrap=840)

                dpg.add_spacer(height=2)
                dpg.add_separator()
                dpg.add_spacer(height=2)

        # -- status --------------------------------------------------------

        def _set_status(self, text: str) -> None:
            dpg.set_value(self.TAG_STATUS, text)

        # -- inputs actions --------------------------------------------------

        def _merge_inputs(self, new_lines: list[str]) -> None:
            existing = [ln.strip() for ln in str(dpg.get_value(self.TAG_INPUTS)).splitlines() if ln.strip()]
            for ln in new_lines:
                s = str(ln).strip()
                if s:
                    existing.append(s)
            dpg.set_value(self.TAG_INPUTS, "\n".join(dedupe_lines(existing)))
            self._sync_entries_from_input()

        def _browse_add_files(self) -> None:
            if sys.platform != "win32":
                self._set_status("File picker is only supported on Windows.")
                return
            paths = _pick_native_files("Select files")
            if paths:
                self._merge_inputs(paths)

        def _clear_inputs(self) -> None:
            dpg.set_value(self.TAG_INPUTS, "")
            self._entries = []
            self._render_entries()
            self._refresh_inputs_font()

        # -- translate / apply / untranslate ---------------------------------

        def _on_translate(self) -> None:
            self._start_translate()

        def _start_translate(self) -> None:
            if self._job_thread and self._job_thread.is_alive():
                return
            self._sync_entries_from_input()
            pending = [e for e in self._entries if e.status in ("pending", "error")]
            if not pending:
                self._set_status("Nothing to translate — every entry is already up to date.")
                return

            settings = Settings(
                api_key=str(dpg.get_value(self.TAG_API_KEY)).strip() if dpg.does_item_exist(self.TAG_API_KEY) else self.settings.api_key,
                model=str(dpg.get_value(self.TAG_MODEL)).strip() if dpg.does_item_exist(self.TAG_MODEL) else self.settings.model,
            )
            auto_rename = self.settings.auto_rename
            if dpg.does_item_exist(self.TAG_AUTO_RENAME):
                auto_rename = bool(dpg.get_value(self.TAG_AUTO_RENAME))
            append_mode = self._append_mode()

            dpg.configure_item(self.TAG_TRANSLATE_BTN, enabled=False)
            self._set_status(f"Translating {len(pending)} entr{'y' if len(pending) == 1 else 'ies'}…")
            self._job_reported = False

            def worker() -> None:
                for e in pending:
                    result = translate_name(e.display_original, settings)
                    if result.ok:
                        e.translation = result.translation
                        e.text_only = result.text_only
                        e.status = "ok"
                        e.error = ""
                        if auto_rename and e.path is not None:
                            self._apply_entry(e, append_mode)
                    else:
                        e.status = "error"
                        e.error = result.error

            self._job_thread = threading.Thread(target=worker, daemon=True)
            self._job_thread.start()

        def _apply_entry(self, e: EntryState, append_mode: bool) -> None:
            if e.path is None or not e.path.exists():
                return
            is_dir = e.path.is_dir()
            new_name, warning = build_output_filename(
                e.path.name, e.translation, e.text_only, append_mode, is_dir=is_dir
            )
            new_path, err = rename_file_apply(e.path, new_name)
            if err:
                e.apply_error = err
                return
            e.path = new_path
            e.applied = True
            e.warning = warning or ""
            e.apply_error = ""

        def _on_apply(self) -> None:
            if self._job_thread and self._job_thread.is_alive():
                self._set_status("A translate job is still running — wait for it to finish.")
                return
            self._sync_entries_from_input()
            append_mode = self._append_mode()
            candidates = [e for e in self._entries if e.status == "ok" and e.path is not None and e.path.exists()]
            skipped_no_file = sum(1 for e in self._entries if e.status == "ok" and e.path is None)
            if not candidates:
                self._set_status("Nothing to rename — translate some entries with a backing file first.")
                return
            renamed = 0
            failed = 0
            for e in candidates:
                before = e.applied
                self._apply_entry(e, append_mode)
                if e.apply_error:
                    failed += 1
                elif not before:
                    renamed += 1
            self._render_entries()
            parts = [f"Renamed {renamed}"]
            if failed:
                parts.append(f"{failed} failed")
            if skipped_no_file:
                parts.append(f"{skipped_no_file} skipped (no backing file)")
            self._set_status(", ".join(parts) + ".")

        def _on_untranslate(self) -> None:
            if self._job_thread and self._job_thread.is_alive():
                self._set_status("A translate job is still running — wait for it to finish.")
                return
            self._sync_entries_from_input()
            candidates = [e for e in self._entries if e.path is not None and e.path.exists()]
            if not candidates:
                self._set_status("Nothing to revert — no entries have a backing file.")
                return
            reverted = 0
            no_history = 0
            failed = 0
            for e in candidates:
                new_path, err = untranslate_file(e.path)
                if err == "No translation history for this file.":
                    no_history += 1
                    continue
                if err:
                    failed += 1
                    e.apply_error = err
                    continue
                e.path = new_path
                e.applied = False
                e.warning = ""
                e.apply_error = ""
                reverted += 1
            self._render_entries()
            parts = [f"Reverted {reverted}"]
            if no_history:
                parts.append(f"{no_history} had no history")
            if failed:
                parts.append(f"{failed} failed")
            self._set_status(", ".join(parts) + ".")

        # -- settings --------------------------------------------------------

        def _collect_settings(self) -> Settings:
            return Settings(
                api_key=str(dpg.get_value(self.TAG_API_KEY)).strip(),
                model=str(dpg.get_value(self.TAG_MODEL)).strip() or "deepseek-chat",
                auto_rename=bool(dpg.get_value(self.TAG_AUTO_RENAME)),
                append_mode=bool(dpg.get_value(self.TAG_APPEND_MODE)),
                inputs_text=str(dpg.get_value(self.TAG_INPUTS)),
            )

        def _on_save_settings(self) -> None:
            self.settings = self._collect_settings()
            config_save_settings(self.settings)
            self._set_status("Settings saved.")

        def _on_toggle_setting(self) -> None:
            self.settings = self._collect_settings()
            config_save_settings(self.settings)
            if self._entries:
                self._render_entries()

        def _on_forget_all(self) -> None:
            if not _confirm("Translate Filename", "Delete all stored translation history? This cannot be undone."):
                return
            forget_all_history()
            self._set_status("All translation history cleared.")

        # -- drag & drop -----------------------------------------------------

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

        def _on_os_drop(self, data, keys) -> None:
            if data is None:
                return
            if isinstance(data, str):
                paths = [ln.strip() for ln in data.splitlines() if ln.strip()] or (
                    [data.strip()] if data.strip() else []
                )
            elif isinstance(data, list):
                paths = [str(p).strip() for p in data if str(p).strip()]
            else:
                s = str(data).strip()
                paths = [s] if s else []
            if paths:
                self._merge_inputs(paths)
            dpg.bind_item_theme(self.TAG_INPUTS, None)

        def _on_os_drag_over(self, keys) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd
            if dpg.is_item_hovered(self.TAG_INPUTS):
                dpg.bind_item_theme(self.TAG_INPUTS, "theme_drop_hover")
                dpg_dnd.set_drop_effect(dpg_dnd.DROPEFFECT.MOVE)
            else:
                dpg.bind_item_theme(self.TAG_INPUTS, None)
                dpg_dnd.set_drop_effect()

        def _on_os_drag_leave(self) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd
            dpg.bind_item_theme(self.TAG_INPUTS, None)
            dpg_dnd.set_drop_effect()

        # -- lifecycle ---------------------------------------------------

        def _poll_input_blur(self) -> None:
            if dpg.does_item_exist(self.TAG_INPUTS) and dpg.is_item_deactivated_after_edit(self.TAG_INPUTS):
                self._sync_entries_from_input()

        def _poll_autosize(self) -> None:
            if not dpg.does_item_exist(self.TAG_STATUS):
                return
            pos = dpg.get_item_pos(self.TAG_STATUS)
            size = dpg.get_item_rect_size(self.TAG_STATUS)
            content_bottom = pos[1] + size[1]
            if content_bottom <= 0:
                return
            target = int(content_bottom) + 30
            target = max(self.WINDOW_MIN_HEIGHT, min(target, self.WINDOW_MAX_HEIGHT))
            if abs(target - dpg.get_viewport_height()) > 2:
                dpg.set_viewport_height(target)

        def _poll_job(self) -> None:
            thread = self._job_thread
            if not thread or thread.is_alive() or self._job_reported:
                return
            self._job_reported = True
            dpg.configure_item(self.TAG_TRANSLATE_BTN, enabled=True)
            ok = sum(1 for e in self._entries if e.status == "ok")
            errs = sum(1 for e in self._entries if e.status == "error")
            renamed = sum(1 for e in self._entries if e.applied)
            self._render_entries()
            parts = [f"Translated {ok}"]
            if renamed:
                parts.append(f"{renamed} renamed")
            if errs:
                parts.append(f"{errs} failed")
            self._set_status(", ".join(parts) + ".")
            self._job_thread = None

        def on_close(self) -> None:
            try:
                config_save_settings(self._collect_settings())
            except OSError:
                pass

        def run(self) -> None:
            while dpg.is_dearpygui_running():
                self._poll_input_blur()
                self._poll_job()
                self._poll_autosize()
                dpg.render_dearpygui_frame()
                dpg.run_callbacks(dpg.get_callback_queue())
            dpg.destroy_context()

    App().run()
