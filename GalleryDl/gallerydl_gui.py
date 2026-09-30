"""PySide6 front-end for the gallery-dl Pinterest scraper."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional

from gallerydl_logic import (
    CACHE_DIR,
    OUTPUT_ROOT,
    PROFILES,
    Settings,
    board_rows,
    config_load_settings,
    config_save_settings,
    refresh_boards,
    run_scrape,
)

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))

import uikit as ui  # noqa: E402
from process_runner import ProcessRunner  # noqa: E402
from uikit import icons  # noqa: E402

ACCENT = "#F43F5E"


class GalleryDlWindow(ui.ToolWindow):
    def __init__(self, settings: Settings):
        self.settings = settings
        self._runner: Optional[ProcessRunner] = None
        console = ui.ConsolePanel(placeholder="Pick a profile and board, then Scrape. gallery-dl output streams here.")
        super().__init__(
            title="Pinterest",
            tagline="gallery-dl board scraper",
            glyph=icons.IMAGE,
            config_dir=CACHE_DIR,
            console=console,
            size=(1120, 720),
        )
        self.jobs = ui.JobHost(self, console)
        self.scrape_btn = ui.button("Scrape", "primary", icons.DOWNLOAD, "Download the chosen board (Ctrl+Enter).",
                                    self.scrape, min_width=120)
        self.jobs.lock_while_running(self.scrape_btn)
        self.header_actions.addWidget(self.scrape_btn)
        self.add_page("scrape", "Scrape", icons.IMAGE, self._build(), overline="Profile & board", on_run=self.scrape)

        root = Path(OUTPUT_ROOT)
        note = ui.RailNote("Saved to", "…" + os.sep + os.sep.join(root.parts[-2:]))
        note.setToolTip(OUTPUT_ROOT)
        note.layout().addWidget(ui.button("Open folder", "ghost", icons.FOLDER, OUTPUT_ROOT, self._open_root))
        self.add_rail_note(note)
        self.on_close = self._save
        self._fill_boards(settings.board_url)

    def _build(self):
        s = self.settings
        self.profile = ui.Segmented([(str(i), label) for i, (label, _) in enumerate(PROFILES)], str(s.profile))
        self.profile.changed.connect(lambda _: self._fill_boards())
        self.board = ui.combo([], tip="Boards come from a cached list; Refresh fetches it again.")
        refresh = ui.button("Refresh boards", "secondary", icons.REFRESH,
                            "Ask gallery-dl for this profile's boards and update the cache.", self.refresh)
        self.jobs.lock_while_running(refresh)
        self.cookies = ui.OptionRow("Use Firefox cookies", "--cookies-from-browser firefox (needed for private "
                                                           "boards).", s.firefox_cookies)
        card = ui.Card("What to scrape", "gallery-dl adds pinterest\\<user>\\<board> under the output folder.",
                       icons.IMAGE)
        card.add(ui.field("Profile", self.profile))
        board_row = ui.row(self.board, refresh)
        board_row.setStretch(0, 1)
        card.add(ui.field("Board", board_row))
        card.add(self.cookies)
        return ui.page(card)

    def _profile_index(self) -> int:
        return int(self.profile.value())

    def _fill_boards(self, select_url: str = "") -> None:
        rows = board_rows(self._profile_index())
        self.board.clear()
        for label, url in rows:
            self.board.addItem(label, url)
        idx = max(0, self.board.findData(select_url)) if select_url else 0
        self.board.setCurrentIndex(idx)

    def _collect(self) -> Settings:
        return Settings(
            profile=self._profile_index(),
            board_url=self.board.currentData() or "",
            firefox_cookies=self.cookies.isChecked(),
        )

    def _start(self, title: str, fn, on_done=None) -> None:
        settings = self._collect()
        self.settings = settings
        config_save_settings(settings)

        def job(emit):
            self._runner = ProcessRunner(emit)
            return fn(settings, self._runner)

        self.jobs.start(title, job, cancel=lambda: self._runner.cancel() if self._runner else None, on_done=on_done)

    def refresh(self) -> None:
        idx = self._profile_index()
        current = self.board.currentData() or ""
        self._start(f"Refresh boards · {PROFILES[idx][0]}",
                    lambda s, r: refresh_boards(idx, s.firefox_cookies, r),
                    on_done=lambda _: self._fill_boards(current))

    def scrape(self) -> None:
        url = self.board.currentData() or PROFILES[self._profile_index()][1]
        self._start(f"Scrape · {self.board.currentText()}", lambda s, r: run_scrape(url, s.firefox_cookies, r))

    def _open_root(self) -> None:
        Path(OUTPUT_ROOT).mkdir(parents=True, exist_ok=True)
        os.startfile(OUTPUT_ROOT)  # type: ignore[attr-defined]

    def _save(self) -> None:
        if self.jobs.is_running() and self._runner is not None:
            self._runner.cancel()
            self.jobs.runner.wait(3.0)
        config_save_settings(self._collect())


def run_gui() -> None:
    ui.create_app("GalleryDlPinterest", ACCENT)
    win = GalleryDlWindow(config_load_settings())
    win.restore_ui()
    win.run_app()
