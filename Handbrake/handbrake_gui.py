"""PySide6 front-end for the HandBrake tool."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from handbrake_logic import (
    CONFIG_DIR,
    PRESET_DIR,
    RANGE_UNIT_FRAMES,
    RANGE_UNIT_SECONDS,
    Settings,
    ValidationError,
    _delete_only_list_file,
    build_encode_options,
    build_initial_files_text,
    cancel_running_jobs,
    config_load_settings,
    config_save_settings,
    list_preset_json_files,
    parse_input_paths,
    parse_range_unit,
    resolve_preset_path,
    run_encode,
)

import uikit as ui
from uikit import icons

ACCENT = "#F59E42"


class HandBrakeWindow(ui.ToolWindow):
    def __init__(self, settings: Settings, files_text: str, only_list: Optional[str]):
        self.settings = settings
        self._only_list = only_list
        self._presets = list_preset_json_files()

        self.files = ui.PathList(
            empty_text="Drop videos or folders here\nFolders are scanned recursively",
            files_title="Select video files",
            noun="path",
        )
        self.files.set_text(files_text)
        files_card = ui.Card("Videos", "Launch from Directory Opus, drop from Explorer, or add below.", icons.OPEN_FILE)
        files_card.add(self.files, 1)

        console = ui.ConsolePanel(placeholder="Pick a preset and files, then Encode. HandBrakeCLI progress streams here.")
        super().__init__(
            title="HandBrake Tool",
            tagline="Batch encode with presets",
            glyph=icons.VIDEO,
            config_dir=CONFIG_DIR,
            inputs=files_card,
            console=console,
        )
        self.jobs = ui.JobHost(self, console)

        self.encode_btn = ui.button("Encode", "primary", icons.PLAY, "Encode every listed video (Ctrl+Enter).",
                                    self.encode, min_width=120)
        self.jobs.lock_while_running(self.encode_btn)
        self.header_actions.addWidget(self.encode_btn)

        self.add_page("encode", "Encode", icons.MOVIE, self._build_encode(), overline="Preset & output",
                      on_run=self.encode)
        self.add_page("tuning", "Quality", icons.SPEED, self._build_tuning(), overline="Picture & rate",
                      on_run=self.encode)
        self.add_page("range", "Range", icons.CUT, self._build_range(), overline="Partial encode",
                      on_run=self.encode)

        note = ui.RailNote("Presets", f"{len(self._presets)} preset JSON file(s) next to the tool.")
        open_btn = ui.button("Open presets folder", "ghost", icons.FOLDER, str(PRESET_DIR),
                             lambda: os.startfile(PRESET_DIR))  # type: ignore[attr-defined]
        note.layout().addWidget(open_btn)
        self.add_rail_note(note)
        self.on_close = self._shutdown

    # ── pages ──────────────────────────────────────────────────────────────
    def _build_encode(self):
        s = self.settings
        names = [p.name for p in self._presets]
        current = s.preset_file if s.preset_file in names else (names[0] if names else "")
        self.preset = ui.combo(names, current, "HandBrake preset JSON from the tool folder.")
        self.container = ui.Segmented(
            [("", "Preset default"), ("mp4", "MP4"), ("mkv", "MKV")],
            s.output_format if s.output_format in ("mp4", "mkv") else "",
        )
        self.replace = ui.OptionRow(
            "Replace original",
            "After a successful encode the original goes to the Recycle Bin, but only when the new file "
            "is meaningfully smaller.",
            s.replace_original,
        )
        card = ui.Card("Preset", "Every file is encoded with the chosen preset.", icons.MOVIE)
        if names:
            card.add(ui.field("Preset", self.preset))
        else:
            card.add(ui.hint(f"No preset JSON files found in {PRESET_DIR}. Export one from HandBrake into that folder."))
        card.add(ui.field("Container", self.container,
                          "Override the preset's container, or keep whatever it specifies."))
        card.add(self.replace)
        return ui.page(card)

    def _build_tuning(self):
        s = self.settings
        self.max_side = ui.line_edit(s.max_side, "1920", "Longest side in pixels. Blank uses 1920.")
        self.quality = ui.line_edit(s.video_quality, "preset", "Passed as -q. Lower is higher quality. Blank keeps the preset.")
        self.framerate = ui.line_edit(s.video_framerate, "source", "Passed as -r. Blank keeps the preset/source rate.")
        pic = ui.Card("Picture & quality", "Overrides applied on top of the preset.", icons.SPEED)
        pic.add(ui.grid_fields(
            ui.field("Max picture side (px)", self.max_side),
            ui.field("Video quality (-q)", self.quality),
            ui.field("Frame rate (-r)", self.framerate),
            columns=3,
        ))
        self.small_cutoff = ui.line_edit(s.small_file_cutoff_mb, "MB", "Inputs smaller than this use the values on the right.")
        self.small_quality = ui.line_edit(s.small_file_quality, "-q", "Replaces Video quality for small inputs.")
        self.small_framerate = ui.line_edit(s.small_file_framerate, "-r", "Blank falls back to the normal frame rate.")
        small = ui.Card("Small files", "Different -q / -r for inputs below a size cutoff.", icons.FILTER)
        small.add(ui.grid_fields(
            ui.field("Smaller than (MB)", self.small_cutoff),
            ui.field("Quality (-q)", self.small_quality),
            ui.field("Frame rate (-r)", self.small_framerate),
            columns=3,
        ))
        return ui.page(pic, small)

    def _build_range(self):
        s = self.settings
        self.range_unit = ui.Segmented([(RANGE_UNIT_FRAMES, "Frames"), (RANGE_UNIT_SECONDS, "Seconds")],
                                       parse_range_unit(s.frame_range_unit))
        self.range_start = ui.line_edit(s.frame_range_start, "start",
                                        "Frames: first frame (zero-based). Seconds: 12.5 or 1:30. Blank = beginning.")
        self.range_end = ui.line_edit(s.frame_range_end, "end",
                                      "Frames: last frame (inclusive). Seconds: 42 or 0:01:05.5. Blank = to the end.")
        card = ui.Card("Encode part of each file", "Leave both blank to encode everything.", icons.CUT)
        card.add(ui.field("Unit", self.range_unit))
        card.add(ui.grid_fields(ui.field("Start", self.range_start), ui.field("End", self.range_end)))
        return ui.page(card)

    # ── running ────────────────────────────────────────────────────────────
    def _collect(self) -> Settings:
        return Settings(
            preset_file=self.preset.currentText(),
            max_side=self.max_side.text().strip() or "1920",
            video_quality=self.quality.text().strip(),
            video_framerate=self.framerate.text().strip(),
            small_file_cutoff_mb=self.small_cutoff.text().strip(),
            small_file_quality=self.small_quality.text().strip(),
            small_file_framerate=self.small_framerate.text().strip(),
            frame_range_start=self.range_start.text().strip(),
            frame_range_end=self.range_end.text().strip(),
            frame_range_unit=self.range_unit.value(),
            output_format=self.container.value(),
            replace_original=self.replace.isChecked(),
            files_text=self.files.text(),
        )

    def _file_paths(self) -> list[Path]:
        return parse_input_paths(self.files.text())

    def encode(self) -> None:
        if self.jobs.is_running():
            self.toast("A job is already running")
            return
        paths = self._file_paths()
        if not paths:
            self.toast("Add videos first")
            return
        if not self._presets:
            ui.dialogs.error(self, "No presets", f"No preset JSON files found in:\n{PRESET_DIR}")
            return
        settings = self._collect()
        preset_path = resolve_preset_path(settings.preset_file)
        if not preset_path:
            self.toast("Choose a preset")
            return
        try:
            options = build_encode_options(settings, preset_path)
        except ValidationError as ex:
            ui.dialogs.error(self, "Check the settings", str(ex))
            return
        self.settings = settings
        config_save_settings(settings)
        n = len(paths)
        self.jobs.start(
            f"Encode · {preset_path.stem} · {n} file{'s' if n != 1 else ''}",
            lambda emit: run_encode(paths, options, on_output=emit),
            cancel=cancel_running_jobs,
        )

    def _shutdown(self) -> None:
        if self.jobs.is_running():
            cancel_running_jobs()
            self.jobs.runner.wait(3.0)
        try:
            config_save_settings(self._collect())
        except OSError:
            pass
        _delete_only_list_file(self._only_list)


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    ui.create_app("HandBrakeTool", ACCENT)
    settings = config_load_settings()
    files_text = build_initial_files_text(settings.files_text, initial_only_list, initial_only_files)
    win = HandBrakeWindow(settings, files_text, initial_only_list)
    win.restore_ui()
    win.run_app()
