"""PySide6 front-end for the yt-dlp Downloader."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import QApplication

from ytdlp_logic import (
    CONFIG_DIR,
    Runner,
    Settings,
    config_load_settings,
    config_save_settings,
    run_download,
    split_clipboard,
)

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))

import uikit as ui  # noqa: E402
from uikit import icons  # noqa: E402

ACCENT = "#FF7A59"


class YtDlpWindow(ui.ToolWindow):
    def __init__(self, settings: Settings, dest: str, url: str, extra: str):
        self.settings = settings
        self._runner: Optional[Runner] = None
        console = ui.ConsolePanel(placeholder="Paste a URL (or a whole yt-dlp command) and Download. "
                                              "yt-dlp output streams here.")
        super().__init__(
            title="yt-dlp",
            tagline="Video & audio downloader",
            glyph=icons.DOWNLOAD,
            config_dir=CONFIG_DIR,
            console=console,
            size=(1180, 760),
        )
        self.jobs = ui.JobHost(self, console)
        self.download_btn = ui.button("Download", "primary", icons.DOWNLOAD, "Start the download (Ctrl+Enter).",
                                      self.download, min_width=130)
        self.jobs.lock_while_running(self.download_btn)
        self.header_actions.addWidget(self.download_btn)

        self.add_page("download", "Download", icons.DOWNLOAD, self._build_download(url, extra, dest),
                      overline="From the clipboard", on_run=self.download)
        self.add_page("network", "Network", icons.GLOBE, self._build_network(), overline="Cookies & updates",
                      on_run=self.download)
        self.add_rail_note(ui.RailNote("Ctrl+click in Opus",
                                       "Downloads the clipboard URL straight away with these settings, "
                                       "into the current folder."))
        self.on_close = self._save

    def _build_download(self, url: str, extra: str, dest: str):
        s = self.settings
        self.url = ui.line_edit(url, "https://…", "Video, playlist or channel URL.")
        paste = ui.icon_button(icons.PASTE, "Paste from clipboard (a yt-dlp command fills the arguments too)",
                               self.paste)
        self.dest = ui.PathField(dest or s.dest, placeholder="Folder to save into", title="Download folder")
        src = ui.Card("Source", "Filled from the clipboard when the window opens.", icons.LINK)
        src.add(ui.field("URL", ui.row(self.url, paste)))
        src.add(ui.field("Save to", self.dest))

        self.mode = ui.Segmented([("audio", "Audio"), ("video", "Video")], "audio" if s.audio else "video")
        self.mode.changed.connect(lambda _: self._sync_mode())
        self.prefix = ui.line_edit(s.file_prefix, "optional", "Text put in front of every file name.")
        self.extra = ui.line_edit(extra, "e.g. --download-sections \"*1:00-2:00\"",
                                  "Passed to yt-dlp as typed, after the options below, so they win.")
        self.mp4 = ui.OptionRow("Save video as MP4", "Merge and remux into an .mp4 container.", s.mp4_container)
        self.metadata = ui.OptionRow("Embed metadata", "Thumbnail, subtitles, tags and chapters.", s.metadata)
        self.date_prefix = ui.OptionRow("Upload date in the name", "[mm-dd-yyyy] before the title.", s.date_prefix)
        fmt = ui.Card("Output", "Section downloads keep the time range in the name and cap video at 1080p.",
                      icons.MOVIE)
        fmt.add(ui.field("Keep", self.mode))
        fmt.add(ui.grid_fields(ui.field("File name prefix", self.prefix),
                               ui.field("Extra yt-dlp arguments", self.extra)))
        fmt.add(ui.grid_fields(self.mp4, self.metadata, self.date_prefix, columns=3))
        self._sync_mode()
        return ui.page(src, fmt)

    def _build_network(self):
        s = self.settings
        self.cookies = ui.OptionRow("Use Firefox cookies", "--cookies-from-browser firefox. If yt-dlp fails, the "
                                                           "download is retried once without cookies.",
                                    s.firefox_cookies)
        self.no_cookies = ui.OptionRow("Disable cookies", "--no-cookies; overrides cookies in yt-dlp's own config.",
                                       s.no_cookies)
        self.impersonate = ui.OptionRow("Impersonate Chrome", "--impersonate chrome; helps with bot detection "
                                                              "(needs curl_cffi).", s.impersonate)
        net = ui.Card("Requests", "", icons.GLOBE)
        for w in (self.cookies, self.no_cookies, self.impersonate):
            net.add(w)
        self.overwrite = ui.OptionRow("Overwrite existing files", "--force-overwrites instead of --no-overwrites.",
                                      s.overwrite)
        self.update = ui.OptionRow("Update yt-dlp first", "Compares with the latest GitHub release and runs "
                                                          "pip install --upgrade when behind.", s.update)
        self.keep_console = ui.OptionRow("Keep the console open after Ctrl+click",
                                         "The headless download window waits for Enter when it finishes.",
                                         s.keep_console)
        misc = ui.Card("Behaviour", "", icons.SETTINGS)
        for w in (self.overwrite, self.update, self.keep_console):
            misc.add(w)
        return ui.page(net, misc)

    def _sync_mode(self) -> None:
        self.mp4.setEnabled(self.mode.value() == "video")

    def paste(self) -> None:
        url, extra = split_clipboard(QApplication.clipboard().text())
        if url:
            self.url.setText(url)
        if extra:
            self.extra.setText(extra)
        if not url and not extra:
            self.toast("The clipboard has no text")

    def _collect(self) -> Settings:
        return Settings(
            audio=self.mode.value() == "audio",
            mp4_container=self.mp4.isChecked(),
            metadata=self.metadata.isChecked(),
            date_prefix=self.date_prefix.isChecked(),
            file_prefix=self.prefix.text().strip(),
            firefox_cookies=self.cookies.isChecked(),
            no_cookies=self.no_cookies.isChecked(),
            overwrite=self.overwrite.isChecked(),
            update=self.update.isChecked(),
            impersonate=self.impersonate.isChecked(),
            keep_console=self.keep_console.isChecked(),
            dest=self.dest.value(),
        )

    def download(self) -> None:
        if self.jobs.is_running():
            self.toast("A download is already running")
            return
        url = self.url.text().strip()
        if not url:
            self.toast("Paste a URL first")
            return
        dest = self.dest.value()
        if not dest or not Path(dest).is_dir():
            ui.dialogs.error(self, "yt-dlp", f"Choose an existing folder to save into.\n\n{dest}")
            return
        settings = self._collect()
        self.settings = settings
        config_save_settings(settings)
        extra = self.extra.text()

        def job(emit):
            self._runner = Runner(emit)
            return run_download(url, dest, settings, extra, self._runner)

        self.jobs.start(f"{'Audio' if settings.audio else 'Video'} · {url}", job,
                        cancel=lambda: self._runner.cancel() if self._runner else None)

    def _save(self) -> None:
        if self.jobs.is_running() and self._runner is not None:
            self._runner.cancel()
            self.jobs.runner.wait(3.0)
        config_save_settings(self._collect())


def run_gui(dest: str = "") -> None:
    app = ui.create_app("YtDlpTool", ACCENT)
    settings = config_load_settings()
    url, extra = split_clipboard(app.clipboard().text())
    if not url.lower().startswith(("http://", "https://")):
        url = ""  # clipboard held something else; don't prefill it
    win = YtDlpWindow(settings, os.path.normpath(dest) if dest else "", url, extra)
    win.restore_ui("download")
    win.run_app()
