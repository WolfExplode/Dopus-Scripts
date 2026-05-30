"""Dear PyGui front-end for FFmpeg tool."""

from __future__ import annotations

import os
import queue
import sys
import textwrap
import threading
from pathlib import Path
from typing import Optional

from ffmpeg_logic import *

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


def _shutdown_gui_app(app) -> None:
    shutdown_ffmpeg_tool(paths=app._file_paths(), only_list=app._only_list_path)


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


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    import dearpygui.dearpygui as dpg

    class App:
        TAG_FILES = "files_input"
        TAG_MODE = "mode_combo"
        TAG_FORMAT = "format_combo"
        TAG_QUALITY = "quality_input"
        TAG_TRIM = "trim_frames_input"
        TAG_REPLACE = "replace_video_check"
        TAG_OUTPUT = "output_text"

        def __init__(self) -> None:
            self._only_list_path = initial_only_list
            self._shutdown_called = False
            self._job_thread: threading.Thread | None = None
            self._job_result_box: list = []
            self._job_reported = False
            self._log_queue: queue.Queue[tuple[str, bool]] = queue.Queue()
            self._output_lines: list[str] = []
            self.settings = config_load_settings()
            files_default = build_initial_files_text(
                self.settings.files_text, initial_only_list, initial_only_files
            )
            self._theme_apply = "theme_btn_apply"
            self._section_tags: dict[str, int | str] = {}

            dpg.create_context()
            self._init_os_drag_drop()
            self._build_themes()

            with dpg.window(tag="primary_window", label="FFmpeg Tool", no_title_bar=True):
                self._build_layout(files_default)

            self._build_fonts()
            dpg.create_viewport(title="FFmpeg Tool", width=980, height=720, min_width=720, min_height=520)
            self._register_os_drag_drop_handlers()
            dpg.setup_dearpygui()
            dpg.show_viewport()
            dpg.set_primary_window("primary_window", True)
            dpg.set_exit_callback(self.on_close)
            self._sync_format_combo()
            self._sync_quality_enabled()

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
            self._hover_tip(hdr, tip)
            return hdr

        def _action_button(self, parent, label: str, action: str, tip: str) -> None:
            btn = dpg.add_button(label=label, callback=lambda: self._run(action), parent=parent, width=-1)
            dpg.bind_item_theme(btn, self._theme_apply)
            self._hover_tip(btn, tip)

        def _build_layout(self, files_default: str) -> None:
            dpg.add_text("FFmpeg Tool", tag="title_main", color=(120, 200, 220))
            dpg.add_text("Video / audio conversion and utilities", color=(130, 138, 155))
            dpg.add_spacer(height=6)

            with dpg.group(horizontal=True):
                with dpg.child_window(width=400, height=-1, border=True, tag="panel_actions"):
                    dpg.bind_item_theme("panel_actions", "theme_actions_panel")

                    hdr_files = self._section(
                        "Selected files",
                        "One path per line. Launch from Directory Opus to fill from your selection.\n"
                        "Drag files from Explorer onto this box to append paths.",
                        "files",
                    )
                    with dpg.group(parent=hdr_files):
                        files_input = dpg.add_input_text(
                            tag=self.TAG_FILES,
                            default_value=files_default,
                            multiline=True,
                            width=-1,
                            height=120,
                            tab_input=False,
                        )
                        self._hover_tip(files_input, "Media files to process.")
                        with dpg.group(horizontal=True):
                            add_btn = dpg.add_button(label="Add files…", callback=self._browse_add_files)
                            clear_btn = dpg.add_button(label="Clear", callback=self._clear_files)
                        self._hover_tip(add_btn, "Append files via file picker.")
                        self._hover_tip(clear_btn, "Clear all paths.")

                    hdr_convert = self._section(
                        "Convert",
                        "Convert selected files to a new format (output beside source, never overwrites).",
                        "convert",
                    )
                    with dpg.group(parent=hdr_convert):
                        dpg.add_text("Mode", color=(150, 158, 175))
                        mode = dpg.add_combo(
                            tag=self.TAG_MODE,
                            items=["Video", "Audio"],
                            default_value="Video" if self.settings.mode == 0 else "Audio",
                            width=-1,
                            callback=self._on_mode_change,
                        )
                        self._hover_tip(mode, "Video or audio conversion presets.")
                        dpg.add_text("Format", color=(150, 158, 175))
                        dpg.add_combo(tag=self.TAG_FORMAT, items=[], width=-1, callback=self._on_format_change)
                        dpg.add_text("Quality (CRF)", color=(150, 158, 175))
                        qual = dpg.add_input_text(
                            tag=self.TAG_QUALITY,
                            default_value=self.settings.quality,
                            width=80,
                        )
                        self._hover_tip(qual, "18–28 typical for video CRF presets (lower = better quality).")
                        dpg.add_spacer(height=4)
                        self._action_button(
                            hdr_convert, "Convert", "convert",
                            "Run conversion on all listed files.",
                        )

                    hdr_rotate = self._section(
                        "Rotate / flip (in place)",
                        "Re-encodes video; audio/subtitles copied. Video files only.",
                        "rotate",
                    )
                    with dpg.group(parent=hdr_rotate):
                        with dpg.group(horizontal=True):
                            for label, act in (
                                ("90° CW", "rotatecw"), ("90° CCW", "rotateccw"),
                                ("Flip H", "fliph"), ("Flip V", "flipv"),
                            ):
                                btn = dpg.add_button(label=label, callback=lambda s, a=act: self._run(a), width=88)
                                dpg.bind_item_theme(btn, self._theme_apply)

                    hdr_trim = self._section(
                        "Trim start (in place)",
                        "Re-encode after skipping the first N frames. Replaces the original file.",
                        "trim",
                    )
                    with dpg.group(parent=hdr_trim):
                        with dpg.group(horizontal=True):
                            dpg.add_text("Skip frames:", color=(150, 158, 175))
                            dpg.add_input_text(
                                tag=self.TAG_TRIM,
                                default_value=self.settings.trim_frames,
                                width=60,
                            )
                        self._action_button(
                            hdr_trim, "Trim & replace", "trimstart",
                            "Skip leading frames and replace original.",
                        )

                    hdr_cover = self._section(
                        "Cover (split/combine)",
                        "1 image + 1 media: embed cover or replace video with still.\n"
                        "Media only: extract .jpg and strip cover.",
                        "cover",
                    )
                    with dpg.group(parent=hdr_cover):
                        cb = dpg.add_checkbox(
                            tag=self.TAG_REPLACE,
                            label="Replace video with image",
                            default_value=self.settings.replace_video_with_image,
                        )
                        self._hover_tip(cb, "Slideshow still + copied audio instead of embedding cover.")
                        self._action_button(
                            hdr_cover, "Split/combine cover", "cover",
                            "Embed, extract, or replace cover.",
                        )

                    hdr_audio = self._section(
                        "Audio tools",
                        "Mono remux, split/combine A/V, or extract each channel to WAV.",
                        "audio",
                    )
                    with dpg.group(parent=hdr_audio):
                        self._action_button(hdr_audio, "Audio → mono", "mono", "Re-encode audio to mono, copy video.")
                        self._action_button(
                            hdr_audio, "Split/combine Audio/Video", "splitav",
                            "Split video to video-only + .audio.mka, or combine 1 video + 1 audio.",
                        )
                        self._action_button(
                            hdr_audio, "All audio ch → WAV", "splitch",
                            "First audio stream: mono WAV per channel (stem.ch01.wav …).",
                        )

                    hdr_merge = self._section(
                        "Merge videos",
                        "2+ videos → output.ext in the first file's folder (sorted by name) with chapter markers. Uses lossless copy when streams match.",
                        "merge",
                    )
                    with dpg.group(parent=hdr_merge):
                        self._action_button(
                            hdr_merge, "Merge with chapters", "mergevid",
                            "Merge sorted by filename; lossless copy when possible, otherwise re-encode with CRF from Convert.",
                        )

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
                        default_value="Pick an action on the left.\n\nFFmpeg output streams here while a job runs.",
                    )

        def _current_formats(self) -> tuple[FormatPreset, ...]:
            return VIDEO_FORMATS if dpg.get_value(self.TAG_MODE) == "Video" else AUDIO_FORMATS

        def _sync_format_combo(self) -> None:
            formats = self._current_formats()
            names = [f.name for f in formats]
            dpg.configure_item(self.TAG_FORMAT, items=names)
            saved = self.settings.format_name
            if saved in names:
                dpg.set_value(self.TAG_FORMAT, saved)
            elif names:
                dpg.set_value(self.TAG_FORMAT, names[0])

        def _sync_quality_enabled(self) -> None:
            formats = self._current_formats()
            fmt_name = dpg.get_value(self.TAG_FORMAT)
            fmt = formats[format_preset_by_name(formats, fmt_name)]
            enabled = quality_applicable(dpg.get_value(self.TAG_MODE) == "Video", fmt)
            dpg.configure_item(self.TAG_QUALITY, enabled=enabled)

        def _on_mode_change(self) -> None:
            self._sync_format_combo()
            self._sync_quality_enabled()

        def _on_format_change(self) -> None:
            self._sync_quality_enabled()

        def _file_paths(self) -> list[Path]:
            return parse_file_paths(str(dpg.get_value(self.TAG_FILES)))

        def _collect_settings(self) -> Settings:
            mode = 0 if dpg.get_value(self.TAG_MODE) == "Video" else 1
            fmt_name = str(dpg.get_value(self.TAG_FORMAT))
            sections: dict[str, bool] = {}
            for key, tag in self._section_tags.items():
                if dpg.does_item_exist(tag):
                    sections[key] = bool(dpg.get_value(tag))
            return Settings(
                mode=mode,
                format_name=fmt_name,
                quality=str(dpg.get_value(self.TAG_QUALITY)).strip() or "23",
                last_action=self.settings.last_action,
                replace_video_with_image=bool(dpg.get_value(self.TAG_REPLACE)),
                trim_frames=str(dpg.get_value(self.TAG_TRIM)).strip() or "1",
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

        def _run(self, action: str) -> None:
            if self._job_thread and self._job_thread.is_alive():
                self._set_output("A job is already running — wait for it to finish or close the window to cancel.")
                return
            paths = self._file_paths()
            if not paths:
                self._set_output("No files listed.\n\nAdd files or launch from Directory Opus with a selection.")
                return
            settings = self._collect_settings()
            settings.last_action = action
            self.settings = settings
            config_save_settings(settings, last_action=action)
            self._log_queue = queue.Queue()
            self._set_output("Running…\n")
            self._job_result_box = []
            self._job_reported = False

            def worker() -> None:
                def on_output(text: str, replace_last: bool) -> None:
                    self._log_queue.put((text, replace_last))

                self._job_result_box.append(
                    run_action(action, paths, settings, on_output=on_output)
                )

            self._job_thread = threading.Thread(target=worker, daemon=True)
            self._job_thread.start()

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

        def _browse_add_files(self) -> None:
            if sys.platform != "win32":
                self._set_output("File picker is only supported on Windows.")
                return
            hint = ""
            for p in self._file_paths():
                hint = os.fspath(p)
                break
            paths = _pick_native_files("Select media files", hint)
            if paths:
                self._merge_files(paths)

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
            _shutdown_gui_app(self)
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
                    dpg.render_dearpygui_frame()
                    dpg.run_callbacks(dpg.get_callback_queue())
            finally:
                self._shutdown()
            dpg.destroy_context()

    App().run()
