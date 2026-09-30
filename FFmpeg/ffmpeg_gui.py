"""PySide6 front-end for the FFmpeg tool."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from ffmpeg_logic import (
    AUDIO_FORMATS,
    CONFIG_DIR,
    MERGE_CONTAINERS,
    MONO_CHANNELS,
    VIDEO_FORMATS,
    Settings,
    build_initial_files_text,
    cancel_running_jobs,
    chapters_to_youtube_text,
    config_load_settings,
    config_save_settings,
    copy_text_to_clipboard,
    cover_combine_preflight,
    format_preset_by_name,
    is_thumb_audio,
    is_thumb_media,
    is_thumb_video,
    merge_preflight_mismatch,
    parse_cut_frame,
    parse_cut_seconds,
    parse_file_paths,
    probe_media_duration_sec,
    probe_video_avg_frame_rate,
    quality_applicable,
    run_action,
    shutdown_ffmpeg_tool,
)

import uikit as ui
from uikit import icons

ACCENT = "#2EC5EA"

ACTION_TITLES = {
    "convert": "Convert",
    "rotatecw": "Rotate 90° clockwise",
    "rotateccw": "Rotate 90° counter-clockwise",
    "fliph": "Flip horizontal",
    "flipv": "Flip vertical",
    "trimstart": "Cut & replace",
    "cover": "Split / combine cover",
    "discardvid": "Discard video",
    "discardaud": "Discard audio",
    "mono": "Audio to mono",
    "splitav": "Split / combine audio & video",
    "splitch": "Audio channels to WAV",
    "mergevid": "Merge with chapters",
    "framesample": "Timelapse sample",
}

PAGE_FOR_ACTION = {
    "convert": "convert", "rotatecw": "rotate", "rotateccw": "rotate", "fliph": "rotate",
    "flipv": "rotate", "trimstart": "cut", "cover": "cover", "discardvid": "cover",
    "discardaud": "audio", "mono": "audio", "splitav": "audio", "splitch": "audio",
    "mergevid": "merge", "framesample": "timelapse",
}


def _format_duration(seconds: float) -> str:
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = seconds % 60
    if abs(secs - round(secs)) < 0.000001:
        return f"{hours:02d}:{minutes:02d}:{int(round(secs)):02d}"
    return f"{hours:02d}:{minutes:02d}:{secs:06.3f}"


class FFmpegWindow(ui.ToolWindow):
    def __init__(self, settings: Settings, files_text: str, only_list: Optional[str]):
        self.settings = settings
        self._only_list = only_list
        self._cut_fps: Optional[float] = None

        self.files = ui.PathList(
            empty_text="Drop media files or folders here\nFolders include every media file inside",
            files_title="Select media files",
            noun="path",
        )
        self.files.set_text(files_text)
        files_card = ui.Card("Files", "Launch from Directory Opus, drop from Explorer, or add below.", icons.OPEN_FILE)
        files_card.add(self.files, 1)

        console = ui.ConsolePanel(placeholder="Pick an action. FFmpeg output streams here while a job runs.")
        super().__init__(
            title="FFmpeg Tool",
            tagline="Convert · cut · merge",
            glyph=icons.MOVIE,
            config_dir=CONFIG_DIR,
            inputs=files_card,
            console=console,
        )
        self.jobs = ui.JobHost(self, console)
        self.last_note = ui.RailNote("Ctrl+click in Opus", "")
        self.add_rail_note(self.last_note)
        self._update_last_note()

        self.add_page("convert", "Convert", icons.VIDEO, self._build_convert(),
                      overline="Transcode", on_run=lambda: self.run("convert"))
        self.add_page("cut", "Cut", icons.CUT, self._build_cut(),
                      overline="Trim in place", on_run=lambda: self.run("trimstart"))
        self.add_page("rotate", "Rotate & flip", icons.ROTATE, self._build_rotate(), overline="Transform in place")
        self.add_page("cover", "Cover art", icons.IMAGE, self._build_cover(),
                      overline="Embed · extract", on_run=lambda: self.run("cover"))
        self.add_page("audio", "Audio", icons.MUSIC, self._build_audio(), overline="Streams & channels")
        self.add_page("merge", "Merge", icons.MERGE, self._build_merge(),
                      overline="Join videos", on_run=lambda: self.run("mergevid"))
        self.add_page("timelapse", "Timelapse", icons.TIMELAPSE, self._build_timelapse(),
                      overline="Frame sampling", on_run=lambda: self.run("framesample"))

        self.files.changed.connect(self._on_files_changed)
        self.on_close = self._shutdown
        self._sync_formats()
        self._autofill_cut_end()
        self._update_cut_duration()

    # ── pages ──────────────────────────────────────────────────────────────
    def _action(self, action: str, text: str = "", kind: str = "primary", glyph: str = icons.PLAY, tip: str = ""):
        btn = ui.button(text or ACTION_TITLES[action], kind, glyph, tip, lambda: self.run(action))
        self.jobs.lock_while_running(btn)
        return btn

    def _build_convert(self):
        s = self.settings
        self.mode = ui.Segmented([("video", "Video"), ("audio", "Audio")], "video" if s.mode == 0 else "audio")
        self.mode.changed.connect(lambda _: self._sync_formats())
        self.format = ui.combo([], tip="Encoder preset for the output file.")
        self.format.currentTextChanged.connect(lambda _: self._sync_quality())
        self.quality = ui.line_edit(s.quality, "23", "Constant rate factor. 18–28 is typical; lower is better quality.")
        self.quality.setMaximumWidth(120)
        card = ui.Card("Convert to a new format",
                       "Output is written next to each source. Existing files are never overwritten.", icons.VIDEO)
        card.add(ui.field("Media type", self.mode))
        card.add(ui.grid_fields(
            ui.field("Format", self.format),
            ui.field("Quality (CRF)", self.quality, "Only for presets that use CRF."),
        ))
        card.add_actions(self._action("convert", tip="Convert every listed file."))
        return ui.page(card)

    def _build_cut(self):
        s = self.settings
        self.cut_unit = ui.Segmented([("Seconds", "Timestamps"), ("Frames", "Frames")], s.cut_unit)
        self.cut_unit.changed.connect(self._on_cut_unit)
        hint_text = "0" if s.cut_unit == "Frames" else "00:00:00"
        self.cut_start = ui.line_edit(s.cut_start, hint_text, "Where the kept range starts.")
        self.cut_end = ui.line_edit(s.cut_end, hint_text, "Where the kept range ends. Blank keeps everything to the end.")
        self.cut_start.textChanged.connect(lambda _: self._update_cut_duration())
        self.cut_end.textChanged.connect(lambda _: self._update_cut_duration())
        self.cut_duration = ui.Pill("Duration: to end", "idle")
        card = ui.Card("Keep a range",
                       "Replaces each original with the selected range. Frames work for video only. "
                       "End is filled from the first media file.", icons.CUT)
        card.add(ui.field("Range unit", self.cut_unit))
        card.add(ui.grid_fields(ui.field("Start", self.cut_start), ui.field("End", self.cut_end)))
        card.add(ui.row(self.cut_duration, None))
        card.add_actions(self._action("trimstart", glyph=icons.CUT, tip="Keep the range and replace each original."))
        return ui.page(card)

    def _build_rotate(self):
        card = ui.Card("Rotate or flip video",
                       "Re-encodes the video stream in place; audio and subtitles are copied. Video files only.",
                       icons.ROTATE)
        tiles = []
        for action, text, glyph in (
            ("rotatecw", "90° clockwise", icons.ROTATE),
            ("rotateccw", "90° counter-clockwise", icons.UNDO),
            ("fliph", "Flip horizontal", icons.FLIP_H),
            ("flipv", "Flip vertical", icons.FLIP_V),
        ):
            b = ui.button(text, "tile", glyph, ACTION_TITLES[action], lambda a=action: self.run(a))
            self.jobs.lock_while_running(b)
            tiles.append(b)
        card.add(ui.row(tiles[0], tiles[1]))
        card.add(ui.row(tiles[2], tiles[3]))
        return ui.page(card)

    def _build_cover(self):
        s = self.settings
        self.replace_video = ui.OptionRow(
            "Replace video with the image",
            "Builds a still-image video with the audio copied, instead of embedding the cover.",
            s.replace_video_with_image,
        )
        pair = ui.Card("Split or combine cover",
                       "Image + media pair up when one file name contains the other (song.wav + song_cover.jpg). "
                       "Media alone: the cover is extracted to .jpg and stripped.", icons.IMAGE)
        pair.add(self.replace_video)
        pair.add_actions(self._action("cover", glyph=icons.IMAGE))
        still = ui.Card("Discard motion video",
                        "Grabs one frame and replaces the video with a still slideshow; audio is copied.",
                        icons.SLIDESHOW)
        still.add_actions(self._action("discardvid", kind="secondary", glyph=icons.SLIDESHOW))
        return ui.page(pair, still)

    def _build_audio(self):
        s = self.settings
        self.mono_channel = ui.combo(MONO_CHANNELS, s.mono_channel,
                                     "auto downmixes every channel. 1–8 keeps only that source channel "
                                     "(1 = L, 2 = R); falls back to downmix if the file has fewer channels.")
        self.mono_channel.setMinimumWidth(96)
        card = ui.Card("Audio tools", "Stream-level edits. Video is copied without re-encoding.", icons.MUSIC)
        card.add(ui.ActionRow("Discard audio", "Remove every audio stream; video copied in place (lossless).",
                              self._action("discardaud", "Discard", "secondary", icons.DELETE)))
        card.add(ui.Divider())
        mono_btn = self._action("mono", "To mono", "secondary", icons.AUDIO)
        mono = ui.ActionRow("Audio to mono", "Re-encode audio to one channel, copy video.", mono_btn)
        mono.layout().insertWidget(1, ui.label("Channel", "FieldLabel"))
        mono.layout().insertWidget(2, self.mono_channel)
        card.add(mono)
        card.add(ui.Divider())
        card.add(ui.ActionRow("Split / combine audio & video",
                              "Split a video into video-only + .audio.mka, or combine 1 video + 1 audio.",
                              self._action("splitav", "Split / combine", "secondary", icons.SHARE)))
        card.add(ui.Divider())
        card.add(ui.ActionRow("Every channel to WAV", "First audio stream: one mono WAV per channel (stem.ch01.wav …).",
                              self._action("splitch", "Extract", "secondary", icons.DOWNLOAD)))
        return ui.page(card)

    def _build_merge(self):
        s = self.settings
        self.merge_container = ui.combo(MERGE_CONTAINERS, s.merge_container,
                                        "auto matches the first file. .mp4 / .mkv force that container.")
        card = ui.Card("Merge videos",
                       "2+ videos, sorted by name, into one file with a chapter per input. Lossless when the "
                       "streams match; otherwise you choose to fix the outliers or re-encode everything.",
                       icons.MERGE)
        card.add(ui.field("Output container", self.merge_container))
        yt = ui.button("Copy YouTube chapters", "secondary", icons.COPY,
                       "2+ files: chapters from file names + durations. 1 merged file: its embedded chapters.",
                       self._copy_chapters)
        self.jobs.lock_while_running(yt)
        card.add_actions(yt, self._action("mergevid", glyph=icons.MERGE))
        return ui.page(card)

    def _build_timelapse(self):
        s = self.settings
        self.sample_start = ui.line_edit(s.sample_start, "3", "First sampled frame, in seconds from the start.")
        self.sample_interval = ui.line_edit(s.sample_interval, "2", "Seconds between samples.")
        self.sample_count = ui.line_edit(s.sample_count, "48", "How many frames the output holds.")
        card = ui.Card("Timelapse sample",
                       "Takes one frame every interval and writes <name>_frames.mp4 (H.264, no audio) "
                       "next to each video.", icons.TIMELAPSE)
        card.add(ui.grid_fields(
            ui.field("Start offset (s)", self.sample_start),
            ui.field("Interval (s)", self.sample_interval),
            ui.field("Output frames", self.sample_count),
            columns=3,
        ))
        card.add_actions(self._action("framesample", "Extract frames", glyph=icons.TIMELAPSE))
        return ui.page(card)

    # ── convert helpers ────────────────────────────────────────────────────
    def _formats(self):
        return VIDEO_FORMATS if self.mode.value() == "video" else AUDIO_FORMATS

    def _sync_formats(self) -> None:
        names = [f.name for f in self._formats()]
        self.format.blockSignals(True)
        self.format.clear()
        self.format.addItems(names)
        if self.settings.format_name in names:
            self.format.setCurrentText(self.settings.format_name)
        self.format.blockSignals(False)
        self._sync_quality()

    def _sync_quality(self) -> None:
        formats = self._formats()
        fmt = formats[format_preset_by_name(formats, self.format.currentText())]
        self.quality.setEnabled(quality_applicable(self.mode.value() == "video", fmt))

    # ── cut helpers ────────────────────────────────────────────────────────
    def _on_files_changed(self) -> None:
        self._cut_fps = None
        self._autofill_cut_end()
        self._update_cut_duration()

    def _on_cut_unit(self, unit: str) -> None:
        zero = "0" if unit == "Frames" else "00:00:00"
        for e in (self.cut_start, self.cut_end):
            e.setPlaceholderText(zero)
        self.cut_start.setText(zero)
        self.cut_end.setText("")
        self._cut_fps = None
        self._autofill_cut_end()
        self._update_cut_duration()

    def _autofill_cut_end(self) -> None:
        media = [p for p in self._file_paths() if is_thumb_media(p.name)]
        if not media:
            return
        unit = self.cut_unit.value()
        if unit == "Frames" and any(is_thumb_audio(p.name) for p in media):
            self.cut_unit.set_value("Seconds")
            unit = "Seconds"
            self.cut_start.setText("00:00:00")
        duration = probe_media_duration_sec(media[0])
        if duration <= 0:
            return
        if unit == "Frames":
            fps = probe_video_avg_frame_rate(media[0])
            self._cut_fps = fps
            self.cut_end.setText(str(max(1, int(round(duration * fps)))) if fps > 0 else "")
        else:
            self.cut_end.setText(_format_duration(duration))

    def _update_cut_duration(self) -> None:
        start_text = self.cut_start.text().strip()
        end_text = self.cut_end.text().strip()
        tone = "busy"
        if not end_text:
            text = "to end"
        elif self.cut_unit.value() == "Frames":
            start, end = parse_cut_frame(start_text), parse_cut_frame(end_text)
            if start >= 0 and end > start:
                frames = end - start
                if self._cut_fps is None:
                    videos = [p for p in self._file_paths() if is_thumb_video(p.name)]
                    self._cut_fps = probe_video_avg_frame_rate(videos[0]) if videos else -1.0
                if self._cut_fps > 0:
                    text = f"{_format_duration(frames / self._cut_fps)}  ·  {frames} frames @ {self._cut_fps:g} fps"
                else:
                    text = f"{frames} frames"
            else:
                text, tone = "invalid range", "error"
        else:
            start, end = parse_cut_seconds(start_text), parse_cut_seconds(end_text)
            if start >= 0 and end > start:
                text = _format_duration(end - start)
            else:
                text, tone = "invalid range", "error"
        self.cut_duration.set_tone(tone, f"Duration  {text}")

    # ── running ────────────────────────────────────────────────────────────
    def _file_paths(self) -> list[Path]:
        return parse_file_paths(self.files.text())

    def _collect(self) -> Settings:
        s = self.settings
        return Settings(
            mode=0 if self.mode.value() == "video" else 1,
            format_name=self.format.currentText(),
            quality=self.quality.text().strip() or "23",
            last_action=s.last_action,
            replace_video_with_image=self.replace_video.isChecked(),
            merge_container=self.merge_container.currentText() or "auto",
            mono_channel=self.mono_channel.currentText() or "auto",
            trim_frames=s.trim_frames,
            cut_unit=self.cut_unit.value(),
            cut_start=self.cut_start.text().strip(),
            cut_end=self.cut_end.text().strip(),
            sample_start=self.sample_start.text().strip() or "3",
            sample_interval=self.sample_interval.text().strip() or "2",
            sample_count=self.sample_count.text().strip() or "48",
            files_text=self.files.text(),
        )

    def _update_last_note(self) -> None:
        title = ACTION_TITLES.get(self.settings.last_action, self.settings.last_action)
        self.last_note.set_text(f"Repeats the last action on the selection:\n{title}")

    def run(self, action: str) -> None:
        if self.jobs.is_running():
            self.toast("A job is already running")
            return
        paths = self._file_paths()
        if not paths:
            if self.files.paths():
                ui.dialogs.error(self, "No media found",
                                 "None of the listed paths are media or image files.\n\n"
                                 "Folders must contain media or image files.")
            else:
                self.toast("Add files first")
            return
        settings = self._collect()
        settings.last_action = action

        if action == "mergevid":
            mismatch = merge_preflight_mismatch(paths)
            if mismatch:
                message, eligible = mismatch
                choice = ui.dialogs.ask_yes_no_cancel(self, "Streams do not match", message)
                if choice is None:
                    return
                if eligible:
                    settings.merge_fix_outliers = choice
                elif not choice:
                    return
                else:
                    settings.merge_fix_outliers = False
        if action == "cover":
            err = cover_combine_preflight(paths)
            if err:
                ui.dialogs.error(self, "Cover match error", err)
                return

        self.settings = settings
        config_save_settings(settings, last_action=action)
        self._update_last_note()
        n = len(paths)
        self.jobs.start(
            f"{ACTION_TITLES.get(action, action)} · {n} file{'s' if n != 1 else ''}",
            lambda emit: run_action(action, paths, settings, on_output=emit),
            cancel=cancel_running_jobs,
        )

    def _copy_chapters(self) -> None:
        paths = self._file_paths()
        if not paths:
            self.toast("Add files first")
            return
        text, err = chapters_to_youtube_text(paths)
        if err:
            self.console.set_text(err, tone="error")
            return
        if not copy_text_to_clipboard(text or ""):
            self.toast("Could not copy to the clipboard")
            return
        n = len([ln for ln in (text or "").splitlines() if ln.strip()])
        self.console.set_text(f"Copied {n} chapter(s) to the clipboard (YouTube format):\n\n{text}")
        self.toast(f"Copied {n} chapters")

    def _shutdown(self) -> None:
        if self.jobs.is_running():
            cancel_running_jobs()
            self.jobs.runner.wait(3.0)
        try:
            config_save_settings(self._collect())
        except OSError:
            pass
        shutdown_ffmpeg_tool(paths=self._file_paths(), only_list=self._only_list)


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
    initial_section: Optional[str] = None,
) -> None:
    ui.create_app("FFmpegTool", ACCENT)
    settings = config_load_settings()
    files_text = build_initial_files_text(settings.files_text, initial_only_list, initial_only_files)
    win = FFmpegWindow(settings, files_text, initial_only_list)
    win.restore_ui(initial_section or "", PAGE_FOR_ACTION.get(settings.last_action, ""))
    win.run_app()
