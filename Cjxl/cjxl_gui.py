"""Dear PyGui front-end for JPEG XL (cjxl) tool."""

from __future__ import annotations

import os
import queue
import sys
import textwrap
import threading
from pathlib import Path
from typing import Optional

from cjxl_logic import *

OUTPUT_WRAP_WIDTH = 100


def _wrap_output_line(line: str) -> str:
    if not line:
        return line
    return textwrap.fill(
        line,
        width=OUTPUT_WRAP_WIDTH,
        break_long_words=True,
        replace_whitespace=False,
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


def _bounded_int(raw: str, default: int, lo: int, hi: int) -> int:
    s = raw.strip()
    if not s:
        return default
    try:
        n = int(float(s))
    except ValueError:
        return default
    return max(lo, min(hi, n))


def _bounded_float(raw: str, default: float, lo: float, hi: float) -> float:
    s = raw.strip()
    if not s:
        return default
    try:
        n = float(s)
    except ValueError:
        return default
    return max(lo, min(hi, n))


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


def _pick_native_folder(title: str, initial: str = "") -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    kwargs: dict = {"title": title, "parent": root}
    init = _browse_initial_dir(initial)
    if init:
        kwargs["initialdir"] = init
    path = filedialog.askdirectory(**kwargs)
    root.destroy()
    return path if path else ""


def _pick_native_files(title: str, initial: str = "") -> list[str]:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    kwargs: dict = {"title": title, "parent": root}
    init = _browse_initial_dir(initial)
    if init:
        kwargs["initialdir"] = init
    paths = filedialog.askopenfilenames(**kwargs)
    root.destroy()
    return list(paths) if paths else []


def _show_error(title: str, message: str) -> None:
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showerror(title, message, parent=root)
    root.destroy()


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
    initial_tab_folder: Optional[str] = None,
) -> None:
    import dearpygui.dearpygui as dpg

    class App:
        ENCODE_MODE_ITEMS = ("Quality (1–100)", "Distance (0–3)")
        QUALITY_MIN = 1
        QUALITY_MAX = 100
        DISTANCE_MIN = 0
        DISTANCE_MAX = 3
        TAG_FILES = "files_input"
        TAG_ENCODE_MODE = "encode_mode_radio"
        TAG_QUALITY_GROUP = "quality_group"
        TAG_QUALITY = "quality_input"
        TAG_DISTANCE_GROUP = "distance_group"
        TAG_DISTANCE = "distance_input"
        TAG_EFFORT = "effort_input"
        TAG_PROGRESSIVE = "progressive_check"
        TAG_REPLACE_SOURCE = "replace_source_check"
        TAG_BIN_DIR = "bin_dir_input"
        TAG_OUTPUT = "output_text"

        def __init__(self) -> None:
            self._tab_folder = (initial_tab_folder or "").strip()
            self._job_thread: threading.Thread | None = None
            self._job_result_box: list[ConvertResult] = []
            self._job_reported = False
            self._log_queue: queue.Queue[tuple[str, bool]] = queue.Queue()
            self._output_lines: list[str] = []
            self.settings = config_load_settings()
            files_default = build_initial_files_text(
                self.settings.files_text,
                initial_only_list,
                initial_only_files,
                initial_tab_folder,
            )
            self._theme_apply = "theme_btn_apply"
            self._section_tags: dict[str, int | str] = {}

            dpg.create_context()
            self._init_os_drag_drop()
            self._build_themes()

            with dpg.window(tag="primary_window", label="JPEG XL (cjxl)", no_title_bar=True):
                self._build_layout(files_default)

            self._build_fonts()
            dpg.create_viewport(
                title="Image → JPEG XL (cjxl)",
                width=920,
                height=640,
                min_width=680,
                min_height=480,
            )
            self._register_os_drag_drop_handlers()
            dpg.setup_dearpygui()
            dpg.show_viewport()
            dpg.set_primary_window("primary_window", True)
            dpg.set_exit_callback(self.on_close)
            self._sync_encode_fields()

        def _build_themes(self) -> None:
            accent = (72, 168, 190)
            accent_h = (92, 198, 220)
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
                    dpg.add_theme_color(dpg.mvThemeCol_CheckMark, accent_h)
                    dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 10)
                    dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 8)
                    dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6)
                    dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 14, 14)
                    dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 8)

            with dpg.theme(tag=self._theme_apply):
                with dpg.theme_component(dpg.mvButton):
                    dpg.add_theme_color(dpg.mvThemeCol_Button, apply_bg)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, apply_h)
                    dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (80, 178, 150))

            with dpg.theme(tag="theme_output_panel"):
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
                    dpg.add_theme_color(dpg.mvThemeCol_Border, (72, 168, 190))

            dpg.bind_theme("app_theme")

        def _build_fonts(self) -> None:
            ui_font_path = _pick_unicode_ui_font()
            with dpg.font_registry():
                if ui_font_path:
                    dpg.bind_font(dpg.add_font(str(ui_font_path), 14))
                    title_font = dpg.add_font(str(ui_font_path), 18)
                    dpg.bind_item_font("title_main", title_font)
                else:
                    segoe = _windows_fonts_dir() / "segoeui.ttf"
                    if segoe.is_file():
                        dpg.bind_font(dpg.add_font(str(segoe), 14))

        def _hover_tip(self, parent: int | str, text: str) -> None:
            with dpg.tooltip(parent, delay=0.4):
                dpg.add_text(text, wrap=400, color=(200, 208, 220))

        def _section(self, title: str, tip: str, key: str):
            hdr = dpg.add_collapsing_header(
                label=title,
                default_open=self.settings.gui_sections.get(key, GUI_SECTION_DEFAULTS[key]),
                tag=dpg.generate_uuid(),
            )
            self._section_tags[key] = hdr
            if tip:
                self._hover_tip(hdr, tip)
            return hdr

        def _build_layout(self, files_default: str) -> None:
            dpg.add_text("Image → JPEG XL", tag="title_main", color=(120, 200, 220))
            dpg.add_text("Bulk encode via libjxl cjxl", color=(130, 138, 155))
            dpg.add_spacer(height=6)

            with dpg.group(horizontal=True):
                with dpg.child_window(width=400, height=-1, border=True, tag="panel_actions"):
                    dpg.bind_item_theme("panel_actions", "theme_actions_panel")

                    hdr_files = self._section(
                        "Input",
                        "Selected images, images in selected folders, or all images in the "
                        "current tab folder if nothing is selected.\n"
                        "PNG/JPEG/GIF/etc. go straight to cjxl; TIFF/WebP/BMP and similar formats "
                        "are converted to PNG first (Pillow).\n"
                        "Launch from Directory Opus to fill from your selection. "
                        "Drag files or folders from Explorer onto the path box.",
                        "files",
                    )
                    with dpg.group(parent=hdr_files):
                        files_input = dpg.add_input_text(
                            tag=self.TAG_FILES,
                            default_value=files_default,
                            multiline=True,
                            width=-1,
                            height=110,
                            tab_input=False,
                        )
                        self._hover_tip(files_input, "One path per line — file or folder.")
                        with dpg.group(horizontal=True):
                            add_btn = dpg.add_button(label="Add files…", callback=self._browse_add_files)
                            add_dir_btn = dpg.add_button(label="Add folder…", callback=self._browse_add_folder)
                            clear_btn = dpg.add_button(label="Clear", callback=self._clear_files)
                        self._hover_tip(add_btn, "Append image files.")
                        self._hover_tip(add_dir_btn, "Append a folder (all images inside, recursively).")
                        self._hover_tip(clear_btn, "Clear all paths.")

                    hdr_encode = self._section("Encode settings", "", "encode")
                    with dpg.group(parent=hdr_encode):
                        dpg.add_radio_button(
                            tag=self.TAG_ENCODE_MODE,
                            items=list(self.ENCODE_MODE_ITEMS),
                            default_value=self.ENCODE_MODE_ITEMS[self.settings.encode_mode],
                            horizontal=True,
                            callback=self._on_encode_mode_change,
                        )
                        tip_quality = (
                            "Maps to an internal target bitrate/quality level, similar to JPEG.\n"
                            "1–100. Typical photos: 85–95. 100 = lossless."
                        )
                        with dpg.group(tag=self.TAG_QUALITY_GROUP):
                            qual_label = dpg.add_text("Quality", color=(150, 158, 175))
                            qual = dpg.add_input_int(
                                tag=self.TAG_QUALITY,
                                default_value=_bounded_int(
                                    self.settings.quality, 90, self.QUALITY_MIN, self.QUALITY_MAX
                                ),
                                min_value=self.QUALITY_MIN,
                                max_value=self.QUALITY_MAX,
                                min_clamped=True,
                                max_clamped=True,
                                step=1,
                                width=80,
                            )
                            self._hover_tip(qual_label, tip_quality)
                            self._hover_tip(qual, tip_quality)
                        tip_distance = (
                            "Distance 0 — mathematically lossless (visually identical, bit-exact reconstruction)\n"
                            "Distance 1 — visually lossless for most content (default for \"good enough\")\n"
                            "Distance 2-3 — slight artifacts visible on close inspection, significant size savings\n"
                            "Higher distances — progressively more aggressive compression, more visible artifacts"
                        )
                        with dpg.group(tag=self.TAG_DISTANCE_GROUP):
                            dist_label = dpg.add_text("Distance", color=(150, 158, 175))
                            dist = dpg.add_input_float(
                                tag=self.TAG_DISTANCE,
                                default_value=_bounded_float(
                                    self.settings.distance, 1.0, self.DISTANCE_MIN, self.DISTANCE_MAX
                                ),
                                min_value=float(self.DISTANCE_MIN),
                                max_value=float(self.DISTANCE_MAX),
                                min_clamped=True,
                                max_clamped=True,
                                step=0.1,
                                format="%.1f",
                                width=80,
                            )
                            self._hover_tip(dist_label, tip_distance)
                            self._hover_tip(dist, tip_distance)
                        dpg.add_text("Effort (1–10, blank = default)", color=(150, 158, 175))
                        effort = dpg.add_input_text(
                            tag=self.TAG_EFFORT,
                            default_value=self.settings.effort,
                            width=80,
                            hint="7",
                        )
                        self._hover_tip(effort, "Higher = smaller files, slower encode. Blank uses cjxl default (7).")
                        prog = dpg.add_checkbox(
                            tag=self.TAG_PROGRESSIVE,
                            label="Progressive decode",
                            default_value=self.settings.progressive,
                        )
                        self._hover_tip(prog, "Decode a rough preview before the full image is ready.")
                        replace_src = dpg.add_checkbox(
                            tag=self.TAG_REPLACE_SOURCE,
                            label="Replace source file",
                            default_value=self.settings.replace_source,
                        )
                        self._hover_tip(
                            replace_src,
                            "After a successful encode, send the original image to the Recycle Bin.",
                        )
                        dpg.add_text("cjxl bin folder", color=(150, 158, 175))
                        dpg.add_input_text(
                            tag=self.TAG_BIN_DIR,
                            default_value=self.settings.cjxl_bin_dir,
                            width=-1,
                        )
                        self._hover_tip(
                            self.TAG_BIN_DIR,
                            "Folder containing cjxl.exe (jxl-x64-windows-static).",
                        )

                    dpg.add_spacer(height=6)
                    btn = dpg.add_button(label="Convert", callback=self._on_convert, width=-1)
                    dpg.bind_item_theme(btn, self._theme_apply)
                    self._hover_tip(btn, "Encode all listed images to JPEG XL.")

                with dpg.child_window(width=-1, height=-1, border=True, tag="panel_output"):
                    dpg.bind_item_theme("panel_output", "theme_output_panel")
                    dpg.add_text("Output", color=(120, 200, 220))
                    dpg.add_spacer(height=4)
                    dpg.add_input_text(
                        tag=self.TAG_OUTPUT,
                        multiline=True,
                        readonly=True,
                        width=-1,
                        height=-1,
                        tab_input=False,
                        default_value=(
                            "Set encode options on the left, then Convert.\n\n"
                            "Progress and cjxl output appear here."
                        ),
                    )

        def _encode_mode_index(self) -> int:
            mode = dpg.get_value(self.TAG_ENCODE_MODE)
            if mode in (1, "1", self.ENCODE_MODE_ITEMS[1]):
                return 1
            return 0

        def _on_encode_mode_change(self) -> None:
            self._sync_encode_fields()

        def _sync_encode_fields(self) -> None:
            quality_active = self._encode_mode_index() == 0
            dpg.configure_item(self.TAG_QUALITY_GROUP, show=quality_active)
            dpg.configure_item(self.TAG_DISTANCE_GROUP, show=not quality_active)

        def _collect_settings(self) -> Settings:
            sections: dict[str, bool] = {}
            for key, tag in self._section_tags.items():
                if dpg.does_item_exist(tag):
                    sections[key] = bool(dpg.get_value(tag))
            return Settings(
                encode_mode=self._encode_mode_index(),
                quality=str(dpg.get_value(self.TAG_QUALITY)),
                distance=str(dpg.get_value(self.TAG_DISTANCE)),
                effort=str(dpg.get_value(self.TAG_EFFORT)).strip(),
                progressive=bool(dpg.get_value(self.TAG_PROGRESSIVE)),
                replace_source=bool(dpg.get_value(self.TAG_REPLACE_SOURCE)),
                cjxl_bin_dir=str(dpg.get_value(self.TAG_BIN_DIR)).strip() or DEFAULT_CJXL_BIN_DIR,
                files_text=str(dpg.get_value(self.TAG_FILES)),
                gui_sections=sections,
            )

        def _sync_output_display(self) -> None:
            dpg.set_value(
                self.TAG_OUTPUT,
                "\n".join(_wrap_output_line(line) for line in self._output_lines),
            )

        def _set_output(self, text: str) -> None:
            self._output_lines = text.splitlines()
            self._sync_output_display()
            if dpg.is_dearpygui_running():
                dpg.render_dearpygui_frame()
                dpg.run_callbacks(dpg.get_callback_queue())

        def _append_stream_line(self, text: str, replace_last: bool) -> None:
            if replace_last and self._output_lines:
                self._output_lines[-1] = text
            else:
                self._output_lines.append(text)
            self._sync_output_display()

        def _drain_log_queue(self) -> None:
            while True:
                try:
                    text, replace_last = self._log_queue.get_nowait()
                except queue.Empty:
                    break
                self._append_stream_line(text, replace_last)

        def _on_convert(self) -> None:
            if self._job_thread and self._job_thread.is_alive():
                self._set_output("A conversion is already running — wait for it to finish.")
                return

            settings = self._collect_settings()
            err = validate_settings(settings)
            if err:
                _show_error("cjxl", err)
                return

            paths_text = settings.files_text.strip()
            if not paths_text and not self._tab_folder:
                self._set_output(
                    "No paths listed.\n\n"
                    "Add image files or folders, or launch from Directory Opus."
                )
                return

            self.settings = settings
            config_save_settings(settings)
            self._log_queue = queue.Queue()
            self._set_output("Converting…\n")
            self._job_result_box = []
            self._job_reported = False

            def worker() -> None:
                def on_output(text: str, replace_last: bool) -> None:
                    self._log_queue.put((text, replace_last))

                self._job_result_box.append(
                    run_convert(
                        paths_text,
                        settings,
                        tab_folder=self._tab_folder or None,
                        on_output=on_output,
                    )
                )

            self._job_thread = threading.Thread(target=worker, daemon=True)
            self._job_thread.start()

        def _poll_job(self) -> None:
            thread = self._job_thread
            if not thread or thread.is_alive() or self._job_reported:
                return
            self._job_reported = True
            self._drain_log_queue()
            if self._job_result_box:
                result = self._job_result_box[0]
                self._append_stream_line("", False)
                self._append_stream_line(result.summary, False)
                if not result.ok:
                    _show_error("cjxl", result.summary)
            self._job_thread = None

        def _browse_hint(self) -> str:
            for ln in str(dpg.get_value(self.TAG_FILES)).splitlines():
                s = ln.strip()
                if s:
                    return s
            return self._tab_folder

        def _browse_add_files(self) -> None:
            if sys.platform != "win32":
                self._set_output("File picker is only supported on Windows.")
                return
            paths = _pick_native_files("Select image files", self._browse_hint())
            if paths:
                self._merge_files(paths)

        def _browse_add_folder(self) -> None:
            if sys.platform != "win32":
                self._set_output("Folder picker is only supported on Windows.")
                return
            folder = _pick_native_folder("Select folder", self._browse_hint())
            if folder:
                self._merge_files([folder])

        def _merge_files(self, new_paths: list[str]) -> None:
            existing = [ln.strip() for ln in str(dpg.get_value(self.TAG_FILES)).splitlines() if ln.strip()]
            for p in new_paths:
                s = str(p).strip()
                if s:
                    existing.append(s)
            dpg.set_value(self.TAG_FILES, "\n".join(dedupe_path_lines(existing)))

        def _clear_files(self) -> None:
            dpg.set_value(self.TAG_FILES, "")

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
                self._merge_files(paths)
            dpg.bind_item_theme(self.TAG_FILES, None)

        def _on_os_drag_over(self, keys) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd
            if dpg.is_item_hovered(self.TAG_FILES):
                dpg.bind_item_theme(self.TAG_FILES, "theme_drop_hover")
                dpg_dnd.set_drop_effect(dpg_dnd.DROPEFFECT.MOVE)
            else:
                dpg.bind_item_theme(self.TAG_FILES, None)
                dpg_dnd.set_drop_effect()

        def _on_os_drag_leave(self) -> None:
            import DearPyGui_DragAndDrop as dpg_dnd
            dpg.bind_item_theme(self.TAG_FILES, None)
            dpg_dnd.set_drop_effect()

        def on_close(self) -> None:
            try:
                config_save_settings(self._collect_settings())
            except OSError:
                pass

        def run(self) -> None:
            while dpg.is_dearpygui_running():
                self._drain_log_queue()
                self._poll_job()
                dpg.render_dearpygui_frame()
                dpg.run_callbacks(dpg.get_callback_queue())
            dpg.destroy_context()

    App().run()
