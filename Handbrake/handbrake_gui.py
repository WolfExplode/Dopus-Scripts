"""Dear PyGui front-end for HandBrake tool."""

from __future__ import annotations

import os
import queue
import sys
import textwrap
import threading
from pathlib import Path
from typing import Optional

from handbrake_logic import (
    Settings,
    build_encode_options,
    build_initial_files_text,
    cancel_running_jobs,
    config_load_settings,
    config_save_settings,
    dedupe_path_lines,
    list_preset_json_files,
    parse_input_paths,
    resolve_preset_path,
    run_encode,
    ValidationError,
    _delete_only_list_file,
    CONFIG_DIR,
)
from dpg_splitter import PanelSplitter

OUTPUT_WRAP_WIDTH = 100


def _wrap_output_line(line: str) -> str:
    if not line:
        return line
    return textwrap.fill(line, width=OUTPUT_WRAP_WIDTH, break_long_words=True, replace_whitespace=False)


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


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    import dearpygui.dearpygui as dpg

    class App:
        TAG_FILES = "files_input"
        TAG_PRESET = "preset_combo"
        TAG_MAXSIDE = "maxside_input"
        TAG_QUALITY = "quality_input"
        TAG_FRAMERATE = "framerate_input"
        TAG_SMALL_CUTOFF = "small_cutoff_input"
        TAG_SMALL_QUALITY = "small_quality_input"
        TAG_SMALL_FRAMERATE = "small_framerate_input"
        TAG_FRAME_START = "frame_start_input"
        TAG_FRAME_END = "frame_end_input"
        TAG_FORMAT = "format_combo"
        TAG_REPLACE = "replace_check"
        TAG_OUTPUT = "output_text"

        FORMAT_DEFAULT_LABEL = "Preset default"
        FORMAT_LABELS = [FORMAT_DEFAULT_LABEL, "mp4", "mkv"]

        @classmethod
        def _format_to_label(cls, output_format: str) -> str:
            return output_format if output_format in ("mp4", "mkv") else cls.FORMAT_DEFAULT_LABEL

        @classmethod
        def _label_to_format(cls, label: str) -> str:
            return label if label in ("mp4", "mkv") else ""

        def __init__(self) -> None:
            self._only_list_path = initial_only_list
            self._shutdown_called = False
            self._job_thread: threading.Thread | None = None
            self._job_result_box: list = []
            self._job_reported = False
            self._log_queue: queue.Queue[tuple[str, bool]] = queue.Queue()
            self._output_lines: list[str] = []
            self._auto_scroll_output = False
            self._splitter = PanelSplitter("panel_actions", left_width=380, min_left=320, min_right=240, config_dir=CONFIG_DIR)

            self.settings = config_load_settings()
            self._preset_rows = list_preset_json_files()
            files_default = build_initial_files_text(
                self.settings.files_text, initial_only_list, initial_only_files
            )

            dpg.create_context()
            self._init_os_drag_drop()
            self._build_themes()

            with dpg.window(tag="primary_window", label="HandBrake Tool", no_title_bar=True):
                self._build_layout(files_default)

            self._build_fonts()
            dpg.create_viewport(title="HandBrake Tool", width=940, height=680, min_width=700, min_height=480)
            self._register_os_drag_drop_handlers()
            dpg.setup_dearpygui()
            dpg.show_viewport()
            dpg.set_primary_window("primary_window", True)
            dpg.set_exit_callback(self.on_close)

        def _build_themes(self) -> None:
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
                    dpg.add_theme_color(dpg.mvThemeCol_CheckMark, (92, 198, 220))
                    dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 10)
                    dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 8)
                    dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6)
                    dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 14, 14)
                    dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 8)

            with dpg.theme(tag="theme_encode_btn"):
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

            self._splitter.build_theme()

            dpg.bind_theme("app_theme")

        def _build_fonts(self) -> None:
            ui_font_path = _pick_unicode_ui_font()
            with dpg.font_registry():
                if ui_font_path:
                    dpg.bind_font(dpg.add_font(str(ui_font_path), 14))
                    title_font = dpg.add_font(str(ui_font_path), 18)
                    dpg.bind_item_font("title_main", title_font)

        def _hover_tip(self, parent, text: str) -> None:
            with dpg.tooltip(parent, delay=0.4):
                dpg.add_text(text, wrap=400, color=(200, 208, 220))

        def _build_layout(self, files_default: str) -> None:
            dpg.add_text("HandBrake Tool", tag="title_main", color=(120, 200, 220))
            dpg.add_text("Batch-encode with a HandBrake preset", color=(130, 138, 155))
            dpg.add_spacer(height=6)

            with dpg.group(horizontal=True):
                with dpg.child_window(width=self._splitter.left_width, height=-1, border=True, tag="panel_actions"):
                    dpg.bind_item_theme("panel_actions", "theme_actions_panel")
                    with dpg.child_window(width=-1, height=140, border=True, tag="panel_files"):
                        dpg.add_text("Selected files", color=(150, 158, 175))
                        files_input = dpg.add_input_text(
                            tag=self.TAG_FILES, default_value=files_default,
                            multiline=True, width=-1, height=70, tab_input=False,
                        )
                        self._hover_tip(
                            files_input,
                            "One path per line (file or folder). Folders are scanned recursively for video files.\n"
                            "Drag files or folders from Explorer onto this box to append paths.",
                        )
                        with dpg.group(horizontal=True):
                            dpg.add_button(label="Add files...", callback=self._browse_add_files)
                            dpg.add_button(label="Add folder...", callback=self._browse_add_folder)
                            dpg.add_button(label="Clear", callback=self._clear_files)

                    with dpg.child_window(width=-1, height=-1, border=True, tag="panel_options"):
                        dpg.add_text("Preset (from Handbrake folder)", color=(150, 158, 175))
                        preset_names = [p.name for p in self._preset_rows]
                        default_preset = self.settings.preset_file if self.settings.preset_file in preset_names else (
                            preset_names[0] if preset_names else ""
                        )
                        combo = dpg.add_combo(tag=self.TAG_PRESET, items=preset_names, default_value=default_preset, width=-1)
                        self._hover_tip(combo, "HandBrake preset JSON from the Handbrake folder.")

                        with dpg.group(horizontal=True):
                            with dpg.group():
                                dpg.add_text("Max picture side (px)", color=(150, 158, 175))
                                max_side = dpg.add_input_text(
                                    tag=self.TAG_MAXSIDE, default_value=self.settings.max_side, width=100
                                )
                                self._hover_tip(
                                    max_side,
                                    "Maximum width or height of the encoded video, in pixels. "
                                    "The longer side is scaled down to this size; blank uses 1920.",
                                )
                            with dpg.group():
                                dpg.add_text("Video quality (-q)", color=(150, 158, 175))
                                quality = dpg.add_input_text(
                                    tag=self.TAG_QUALITY, default_value=self.settings.video_quality, width=100
                                )
                                self._hover_tip(
                                    quality,
                                    "Optional HandBrake quality value passed as -q. Lower values generally "
                                    "mean higher quality and larger files. Blank keeps the preset value.",
                                )

                        dpg.add_text("Frame rate (-r, optional)", color=(150, 158, 175))
                        framerate = dpg.add_input_text(
                            tag=self.TAG_FRAMERATE, default_value=self.settings.video_framerate, width=100
                        )
                        self._hover_tip(
                            framerate,
                            "Optional frame rate passed as -r. Blank keeps the frame rate from the preset/source.",
                        )

                        dpg.add_spacer(height=4)
                        dpg.add_text("If smaller than (MB): -q / -r overrides", color=(150, 158, 175))
                        with dpg.group(horizontal=True):
                            small_cutoff = dpg.add_input_text(
                                tag=self.TAG_SMALL_CUTOFF,
                                default_value=self.settings.small_file_cutoff_mb,
                                width=90,
                                hint="MB",
                            )
                            self._hover_tip(
                                small_cutoff,
                                "Size threshold in MB. For each input smaller than this value, use the "
                                "small-file -q and -r overrides in the next fields.",
                            )
                            small_quality = dpg.add_input_text(
                                tag=self.TAG_SMALL_QUALITY,
                                default_value=self.settings.small_file_quality,
                                width=70,
                                hint="-q",
                            )
                            self._hover_tip(
                                small_quality,
                                "The -q quality value used for files below the MB cutoff. "
                                "This replaces the normal Video quality value.",
                            )
                            small_framerate = dpg.add_input_text(
                                tag=self.TAG_SMALL_FRAMERATE,
                                default_value=self.settings.small_file_framerate,
                                width=70,
                                hint="-r",
                            )
                            self._hover_tip(
                                small_framerate,
                                "Optional -r frame rate used for files below the MB cutoff. "
                                "Blank falls back to the normal Frame rate value.",
                            )

                        dpg.add_spacer(height=4)
                        dpg.add_text("Frame range (optional)", color=(150, 158, 175))
                        with dpg.group(horizontal=True):
                            frame_start = dpg.add_input_text(
                                tag=self.TAG_FRAME_START,
                                default_value=self.settings.frame_range_start,
                                width=90,
                                hint="start",
                            )
                            self._hover_tip(
                                frame_start,
                                "Optional first frame to encode (zero-based). Blank starts at the beginning.",
                            )
                            frame_end = dpg.add_input_text(
                                tag=self.TAG_FRAME_END,
                                default_value=self.settings.frame_range_end,
                                width=90,
                                hint="end",
                            )
                            self._hover_tip(
                                frame_end,
                                "Optional last frame to encode (inclusive). Blank continues to the end.",
                            )

                        dpg.add_spacer(height=4)
                        dpg.add_text("Output container", color=(150, 158, 175))
                        format_combo = dpg.add_combo(
                            tag=self.TAG_FORMAT,
                            items=self.FORMAT_LABELS,
                            default_value=self._format_to_label(self.settings.output_format),
                            width=160,
                        )
                        self._hover_tip(format_combo, "Override the container from the preset. \"Preset default\" keeps whatever the preset specifies.")

                        dpg.add_spacer(height=6)
                        dpg.add_checkbox(
                            tag=self.TAG_REPLACE, label="Replace original (Recycle Bin original after encode)",
                            default_value=self.settings.replace_original,
                        )

                        dpg.add_spacer(height=8)
                        encode_btn = dpg.add_button(label="Encode", callback=self._run_encode, width=-1, height=28)
                        dpg.bind_item_theme(encode_btn, "theme_encode_btn")

                self._splitter.add_handle()

                with dpg.child_window(width=-1, height=-1, border=True, tag="panel_output"):
                    dpg.bind_item_theme("panel_output", "theme_output_panel")
                    dpg.add_text("Output", color=(120, 200, 220))
                    dpg.add_spacer(height=4)
                    dpg.add_input_text(
                        tag=self.TAG_OUTPUT, multiline=True, readonly=True, width=-1, height=-1,
                        tab_input=False, default_value="Pick a preset and files, then Encode.",
                    )

        def _file_paths(self) -> list[Path]:
            return parse_input_paths(str(dpg.get_value(self.TAG_FILES)))

        def _collect_settings(self) -> Settings:
            return Settings(
                preset_file=str(dpg.get_value(self.TAG_PRESET) or ""),
                max_side=str(dpg.get_value(self.TAG_MAXSIDE)).strip() or "1920",
                video_quality=str(dpg.get_value(self.TAG_QUALITY)).strip(),
                video_framerate=str(dpg.get_value(self.TAG_FRAMERATE)).strip(),
                small_file_cutoff_mb=str(dpg.get_value(self.TAG_SMALL_CUTOFF)).strip(),
                small_file_quality=str(dpg.get_value(self.TAG_SMALL_QUALITY)).strip(),
                small_file_framerate=str(dpg.get_value(self.TAG_SMALL_FRAMERATE)).strip(),
                frame_range_start=str(dpg.get_value(self.TAG_FRAME_START)).strip(),
                frame_range_end=str(dpg.get_value(self.TAG_FRAME_END)).strip(),
                output_format=self._label_to_format(str(dpg.get_value(self.TAG_FORMAT) or "")),
                replace_original=bool(dpg.get_value(self.TAG_REPLACE)),
                files_text=str(dpg.get_value(self.TAG_FILES)),
            )

        def _sync_output_display(self) -> None:
            dpg.set_value(self.TAG_OUTPUT, "\n".join(_wrap_output_line(line) for line in self._output_lines))

        def _set_output(self, text: str) -> None:
            self._output_lines = text.splitlines()
            self._sync_output_display()

        def _append_stream_line(self, text: str, replace_last: bool) -> None:
            if replace_last and self._output_lines:
                self._output_lines[-1] = text
            else:
                self._output_lines.append(text)
            if len(self._output_lines) > 300:
                self._output_lines = self._output_lines[-300:]
            self._sync_output_display()

        def _drain_log_queue(self) -> None:
            while True:
                try:
                    text, replace_last = self._log_queue.get_nowait()
                except queue.Empty:
                    break
                self._append_stream_line(text, replace_last)

        def _run_encode(self) -> None:
            if self._job_thread and self._job_thread.is_alive():
                self._set_output("A job is already running -- wait for it to finish or close the window to cancel.")
                return
            paths = self._file_paths()
            if not paths:
                self._set_output("No files listed.\n\nAdd files or launch from Directory Opus with a selection.")
                return
            if not self._preset_rows:
                self._set_output(f"No preset JSON files found in:\n{self._preset_dir_hint()}")
                return

            settings = self._collect_settings()
            preset_path = resolve_preset_path(settings.preset_file)
            if not preset_path:
                self._set_output("No preset selected.")
                return
            try:
                options = build_encode_options(settings, preset_path)
            except ValidationError as ex:
                self._set_output(str(ex))
                return

            self.settings = settings
            config_save_settings(settings)
            self._log_queue = queue.Queue()
            self._set_output("Running...\n")
            self._job_result_box = []
            self._job_reported = False
            self._auto_scroll_output = True

            def worker() -> None:
                def on_output(text: str, replace_last: bool) -> None:
                    self._log_queue.put((text, replace_last))

                self._job_result_box.append(run_encode(paths, options, on_output=on_output))

            self._job_thread = threading.Thread(target=worker, daemon=True)
            self._job_thread.start()

        def _preset_dir_hint(self) -> str:
            from handbrake_logic import PRESET_DIR
            return str(PRESET_DIR)

        def _poll_job(self) -> None:
            thread = self._job_thread
            if not thread or thread.is_alive() or self._job_reported:
                return
            self._job_reported = True
            if self._shutdown_called:
                self._job_thread = None
                return
            self._drain_log_queue()
            if self._job_result_box:
                result = self._job_result_box[0]
                self._append_stream_line("", False)
                self._append_stream_line(result.summary, False)
            self._job_thread = None
            self._auto_scroll_output = False

        def _browse_add_files(self) -> None:
            if sys.platform != "win32":
                self._set_output("File picker is only supported on Windows.")
                return
            hint = ""
            for ln in str(dpg.get_value(self.TAG_FILES)).splitlines():
                if ln.strip():
                    hint = ln.strip()
                    break
            paths = _pick_native_files("Select media files", hint)
            if paths:
                self._merge_files(paths)

        def _browse_add_folder(self) -> None:
            if sys.platform != "win32":
                self._set_output("Folder picker is only supported on Windows.")
                return
            hint = ""
            for ln in str(dpg.get_value(self.TAG_FILES)).splitlines():
                if ln.strip():
                    hint = ln.strip()
                    break
            folder = _pick_native_folder("Select folder", hint)
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
                paths = [ln.strip() for ln in data.splitlines() if ln.strip()] or ([data.strip()] if data.strip() else [])
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
            self._shutdown()

        def _shutdown(self) -> None:
            if self._shutdown_called:
                return
            self._shutdown_called = True
            if self._job_thread and self._job_thread.is_alive():
                cancel_running_jobs()
                self._job_thread.join(timeout=3.0)
            try:
                config_save_settings(self._collect_settings())
            except OSError:
                pass
            _delete_only_list_file(self._only_list_path)
            if sys.platform == "win32":
                try:
                    import DearPyGui_DragAndDrop as dpg_dnd
                    dpg_dnd.destroy()
                except Exception:
                    pass

        def run(self) -> None:
            try:
                while dpg.is_dearpygui_running():
                    self._drain_log_queue()
                    self._poll_job()
                    self._splitter.update()
                    dpg.render_dearpygui_frame()
                    dpg.run_callbacks(dpg.get_callback_queue())
                    if self._auto_scroll_output:
                        try:
                            dpg.set_y_scroll(self.TAG_OUTPUT, 1_000_000)
                        except Exception:
                            pass
            finally:
                self._shutdown()
            dpg.destroy_context()

    App().run()
