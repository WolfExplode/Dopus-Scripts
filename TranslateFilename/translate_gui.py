"""PySide6 front-end for Translate Filename: batch translate, review, rename.

Each input line becomes an entry card: original name, editable translation,
editable text-only phrase (used by append mode) and any notes. Qt falls back
across fonts per character, so mixed Hangul/Kana/Han names render correctly.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from translate_logic import (
    CONFIG_DIR,
    DEFAULT_PROVIDER,
    PROVIDERS,
    Settings,
    build_initial_inputs_text,
    build_output_filename,
    config_load_settings,
    config_save_settings,
    dedupe_lines,
    default_model_for,
    forget_all_history,
    rename_file_apply,
    translate_name,
    untranslate_file,
)

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))

import uikit as ui  # noqa: E402
from uikit import icons  # noqa: E402

ACCENT = "#F472B6"
NO_HISTORY = "No translation history for this file."


@dataclass
class Entry:
    raw_line: str
    path: Optional[Path]
    original: str
    status: str = "pending"  # pending | ok | error
    translation: str = ""
    text_only: str = ""
    error: str = ""
    warning: str = ""
    apply_error: str = ""
    applied: bool = False


def _entry_for(line: str) -> Entry:
    p = Path(line)
    if p.is_file() or p.is_dir():
        return Entry(line, p, p.name)
    return Entry(line, None, line)


class EntryCard(QFrame):
    """One input: original name, translation fields, status and notes."""

    def __init__(self, entry: Entry, window: "TranslateWindow"):
        super().__init__()
        self.entry = entry
        self.window = window
        self.setObjectName("Inset")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(8)

        head = QHBoxLayout()
        head.setSpacing(10)
        glyph = icons.CHARACTERS if entry.path is None else (icons.FOLDER if entry.path.is_dir() else icons.DOCUMENT)
        head.addWidget(ui.Glyph(glyph, 14, ui.palette().muted))
        self.original = QLabel()
        self.original.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.original.setWordWrap(True)
        self.original.setStyleSheet("font-weight: 600;")
        head.addWidget(self.original, 1)
        self.pill = ui.Pill()
        head.addWidget(self.pill, 0, Qt.AlignTop)
        lay.addLayout(head)

        p = ui.palette()
        self.translation = ui.line_edit(entry.translation, "Translation appears here", "Full translated name (edit freely).")
        self.translation.setStyleSheet(f"QLineEdit {{ color: {p.text}; }}")
        self.text_only = ui.line_edit(entry.text_only, "Text-only phrase", "Used by append mode: original + this phrase.")
        self.text_only.setStyleSheet(f"QLineEdit {{ color: {p.warning}; }}")
        self.translation.textEdited.connect(self._edited)
        self.text_only.textEdited.connect(self._edited)
        lay.addWidget(self.translation)
        lay.addWidget(self.text_only)

        self.notes = QVBoxLayout()
        self.notes.setSpacing(2)
        lay.addLayout(self.notes)
        self.refresh()

    def _edited(self, _text: str) -> None:
        self.entry.translation = self.translation.text()
        self.entry.text_only = self.text_only.text()
        self._render_notes()

    def refresh(self) -> None:
        e = self.entry
        self.original.setText(e.path.name if e.path is not None else e.original)
        if self.translation.text() != e.translation:
            self.translation.setText(e.translation)
        if self.text_only.text() != e.text_only:
            self.text_only.setText(e.text_only)
        if e.applied:
            self.pill.set_tone("ok", "Renamed")
        elif e.status == "ok":
            self.pill.set_tone("busy", "Translated")
        elif e.status == "error":
            self.pill.set_tone("error", "Error")
        elif e.status == "working":
            self.pill.set_tone("warn", "Translating")
        else:
            self.pill.set_tone("idle", "Pending")
        self._render_notes()

    def _render_notes(self) -> None:
        while self.notes.count():
            item = self.notes.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        p = ui.palette()
        e = self.entry
        notes: list[tuple[str, str]] = []
        if e.status == "error":
            notes.append((p.danger, f"Translate error: {e.error}"))
        if e.apply_error:
            notes.append((p.danger, e.apply_error))
        if e.status == "ok":
            is_dir = e.path.is_dir() if e.path is not None else False
            new_name, warn = build_output_filename(e.original, e.translation, e.text_only,
                                                   self.window.append_mode(), is_dir=is_dir)
            if warn:
                notes.append((p.warning, warn))
            if not e.applied and e.path is not None:
                notes.append((p.muted, f"Will rename to  {new_name}"))
        if e.path is None:
            notes.append((p.muted, "Text only: no file behind this line, so Rename and Revert skip it."))
        for colour, text in notes:
            lbl = QLabel(text)
            lbl.setWordWrap(True)
            lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            lbl.setStyleSheet(f"color: {colour}; font-size: 9pt;")
            self.notes.addWidget(lbl)


class TranslateWindow(ui.ToolWindow):
    def __init__(self, settings: Settings, inputs_text: str):
        self.settings = settings
        self.entries: list[Entry] = []
        self.cards: dict[int, EntryCard] = {}

        self.inputs = ui.PathList(
            empty_text="Drop files or folders here\nor paste lines of text (Ctrl+V)",
            files_title="Select files",
            noun="entry",
            plural="entries",
            text_items=True,
        )
        self.inputs.set_text(inputs_text)
        inputs_card = ui.Card("Inputs", "Each path or text line becomes an entry below.", icons.OPEN_FILE)
        inputs_card.add(self.inputs, 1)

        console = ui.ConsolePanel(placeholder="Translate results and rename outcomes are logged here.")
        super().__init__(
            title="Translate Filename",
            tagline="File names → English",
            glyph=icons.TRANSLATE,
            config_dir=CONFIG_DIR,
            inputs=inputs_card,
            console=console,
        )
        self.jobs = ui.JobHost(self, console)

        self.translate_btn = ui.button("Translate", "primary", icons.TRANSLATE,
                                       "Translate every new or failed entry (Ctrl+Enter).", self.translate)
        self.rename_btn = ui.button("Rename", "secondary", icons.RENAME,
                                    "Rename each translated file on disk.", self.rename)
        self.revert_btn = ui.button("Revert", "ghost", icons.UNDO,
                                    "Rename files back to their original names from history.", self.revert)
        self.jobs.lock_while_running(self.translate_btn, self.rename_btn, self.revert_btn)
        for b in (self.revert_btn, self.rename_btn, self.translate_btn):
            self.header_actions.addWidget(b)

        self.entry_host = QWidget()
        self.entry_layout = QVBoxLayout(self.entry_host)
        self.entry_layout.setContentsMargins(0, 0, 0, 0)
        self.entry_layout.setSpacing(10)
        self.add_page("translate", "Translate", icons.TRANSLATE, self.entry_host, overline="Review & rename",
                      on_run=self.translate)
        self.add_page("settings", "Settings", icons.SETTINGS, self._build_settings(), overline="Provider & behaviour",
                      show_inputs=False)
        self.provider_note = ui.RailNote("Provider", "")
        self.add_rail_note(self.provider_note)
        self._update_provider_note()

        self.inputs.changed.connect(self.sync_entries)
        self.jobs.runner.output.connect(lambda *_: self._refresh_cards())
        self.on_close = self._save
        self.sync_entries()

    # ── settings page ──────────────────────────────────────────────────────
    def _build_settings(self):
        s = self.settings
        self.provider = ui.Segmented([(k, info["label"]) for k, info in PROVIDERS.items()],
                                     s.provider if s.provider in PROVIDERS else DEFAULT_PROVIDER)
        self.provider.changed.connect(self._on_provider)
        self.deepseek_key = ui.line_edit(s.api_key, "sk-…", password=True)
        self.kimi_key = ui.line_edit(s.kimi_api_key, "sk-…", password=True)
        self.model = ui.line_edit(s.model, default_model_for(s.provider))
        api = ui.Card("Provider", "Keys are stored in %APPDATA%\\TranslateFilename\\settings.json.", icons.GLOBE)
        api.add(ui.field("Service", self.provider))
        api.add(ui.grid_fields(ui.field("DeepSeek API key", self.deepseek_key),
                               ui.field("Kimi API key", self.kimi_key)))
        api.add(ui.field("Model", self.model, "Blank uses the provider's default model."))
        save = ui.button("Save settings", "primary", icons.SAVE, "", self._save_settings)
        api.add_actions(save)

        self.auto_rename = ui.OptionRow(
            "Rename right after translating",
            "Translate also renames each file. Ctrl+click in Opus always does this.", s.auto_rename)
        self.append = ui.OptionRow(
            "Append mode",
            "Keep the original name and append the translated phrase: original + ' ' + text-only + extension.",
            s.append_mode)
        self.auto_rename.box.toggled.connect(lambda _: self._save_settings(quiet=True))
        self.append.box.toggled.connect(lambda _: self._on_append_toggled())
        behaviour = ui.Card("Behaviour", "", icons.SETTINGS)
        behaviour.add(self.auto_rename)
        behaviour.add(self.append)

        forget = ui.button("Forget all history", "danger", icons.DELETE,
                           "Deletes the original-name history that Revert uses.", self._forget)
        history = ui.Card("History", "Every rename records the original name so Revert can undo it.", icons.CLOCK)
        history.add_actions(forget)
        return ui.page(api, behaviour, history)

    def _on_provider(self, provider: str) -> None:
        current = self.model.text().strip()
        known = {info["default_model"] for info in PROVIDERS.values()}
        if not current or current in known:
            self.model.setText(default_model_for(provider))
        self._save_settings(quiet=True)

    def _update_provider_note(self) -> None:
        s = self.settings
        label = PROVIDERS.get(s.provider, PROVIDERS[DEFAULT_PROVIDER])["label"]
        key = "key set" if s.active_api_key() else "no API key yet"
        self.provider_note.set_text(f"{label} · {s.model}\n{key}")

    def _collect(self) -> Settings:
        provider = self.provider.value()
        return Settings(
            provider=provider,
            api_key=self.deepseek_key.text().strip(),
            kimi_api_key=self.kimi_key.text().strip(),
            model=self.model.text().strip() or default_model_for(provider),
            auto_rename=self.auto_rename.isChecked(),
            append_mode=self.append.isChecked(),
            inputs_text=self.inputs.text(),
        )

    def _save_settings(self, quiet: bool = False) -> None:
        self.settings = self._collect()
        config_save_settings(self.settings)
        self._update_provider_note()
        if not quiet:
            self.toast("Settings saved")

    def _on_append_toggled(self) -> None:
        self._save_settings(quiet=True)
        self._refresh_cards()

    def _forget(self) -> None:
        if ui.dialogs.confirm(self, "Translate Filename",
                              "Delete all stored translation history? Revert will no longer work for files "
                              "renamed so far.", "Delete history"):
            forget_all_history()
            self.toast("Translation history cleared")

    def append_mode(self) -> bool:
        return self.append.isChecked() if hasattr(self, "append") else self.settings.append_mode

    # ── entries ────────────────────────────────────────────────────────────
    def sync_entries(self) -> None:
        lines = dedupe_lines([ln.strip() for ln in self.inputs.paths() if ln.strip()])
        old = {e.raw_line: e for e in self.entries}
        self.entries = [old.get(ln) or _entry_for(ln) for ln in lines]
        self._render_cards()

    def _render_cards(self) -> None:
        while self.entry_layout.count():
            item = self.entry_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.cards = {}
        if not self.entries:
            empty = ui.Card("Nothing to translate yet",
                            "Add files above, or launch from Directory Opus with a selection.", icons.TRANSLATE)
            self.entry_layout.addWidget(empty)
        for e in self.entries:
            card = EntryCard(e, self)
            self.cards[id(e)] = card
            self.entry_layout.addWidget(card)
        self.entry_layout.addStretch(1)

    def _refresh_cards(self) -> None:
        for card in self.cards.values():
            card.refresh()

    # ── actions ────────────────────────────────────────────────────────────
    def _apply_entry(self, e: Entry, append_mode: bool) -> None:
        if e.path is None or not e.path.exists():
            return
        new_name, warning = build_output_filename(e.path.name, e.translation, e.text_only, append_mode,
                                                  is_dir=e.path.is_dir())
        new_path, err = rename_file_apply(e.path, new_name)
        if err:
            e.apply_error = err
            return
        e.path = new_path
        e.applied = True
        e.warning = warning or ""
        e.apply_error = ""

    def translate(self) -> None:
        if self.jobs.is_running():
            return
        pending = [e for e in self.entries if e.status in ("pending", "error")]
        if not pending:
            self.toast("Every entry is already translated")
            return
        settings = self._collect()
        if not settings.active_api_key():
            self.select_page("settings")
            self.toast("Add an API key first")
            return
        self.settings = settings
        config_save_settings(settings)
        auto_rename = settings.auto_rename
        append_mode = settings.append_mode
        for e in pending:
            e.status = "working"
        self._refresh_cards()

        def job(emit):
            ok = failed = renamed = 0
            for e in pending:
                result = translate_name(e.original, settings)
                if result.ok:
                    e.translation, e.text_only = result.translation, result.text_only
                    e.status, e.error = "ok", ""
                    ok += 1
                    emit(f"✓ {e.original}  →  {e.translation}", False)
                    if auto_rename and e.path is not None:
                        self._apply_entry(e, append_mode)
                        if e.applied:
                            renamed += 1
                            emit(f"  renamed to {e.path.name}", False)
                        elif e.apply_error:
                            emit(f"  rename failed: {e.apply_error}", False)
                else:
                    e.status, e.error = "error", result.error
                    failed += 1
                    emit(f"Error: {e.original}: {result.error}", False)
            parts = [f"Translated {ok}"]
            if renamed:
                parts.append(f"{renamed} renamed")
            if failed:
                parts.append(f"{failed} failed")
            return _Result(failed == 0, ", ".join(parts) + ".")

        n = len(pending)
        self.jobs.start(f"Translating {n} entr{'y' if n == 1 else 'ies'} with {settings.model}", job,
                        on_done=lambda _: self._after_translate())

    def _after_translate(self) -> None:
        self._sync_inputs_to_paths()
        self._refresh_cards()

    def rename(self) -> None:
        append_mode = self.append_mode()
        candidates = [e for e in self.entries if e.status == "ok" and e.path is not None and e.path.exists()]
        skipped = sum(1 for e in self.entries if e.status == "ok" and e.path is None)
        if not candidates:
            self.toast("Translate entries with a file behind them first")
            return
        self.console.begin("Rename")
        renamed = failed = 0
        for e in candidates:
            before = e.applied
            old_name = e.path.name
            self._apply_entry(e, append_mode)
            if e.apply_error:
                failed += 1
                self.console.append(f"Error: {old_name}: {e.apply_error}")
            elif not before:
                renamed += 1
                self.console.append(f"{old_name}  →  {e.path.name}")
        parts = [f"Renamed {renamed}"]
        if failed:
            parts.append(f"{failed} failed")
        if skipped:
            parts.append(f"{skipped} skipped (no file)")
        self.console.finish(failed == 0, ", ".join(parts) + ".")
        self._sync_inputs_to_paths()
        self._refresh_cards()

    def revert(self) -> None:
        candidates = [e for e in self.entries if e.path is not None and e.path.exists()]
        if not candidates:
            self.toast("No entries have a file behind them")
            return
        self.console.begin("Revert to original names")
        reverted = no_history = failed = 0
        for e in candidates:
            old_name = e.path.name
            new_path, err = untranslate_file(e.path)
            if err == NO_HISTORY:
                no_history += 1
                continue
            if err:
                failed += 1
                e.apply_error = err
                self.console.append(f"Error: {old_name}: {err}")
                continue
            e.path, e.applied, e.warning, e.apply_error = new_path, False, "", ""
            e.original = new_path.name
            reverted += 1
            self.console.append(f"{old_name}  →  {new_path.name}")
        parts = [f"Reverted {reverted}"]
        if no_history:
            parts.append(f"{no_history} had no history")
        if failed:
            parts.append(f"{failed} failed")
        self.console.finish(failed == 0, ", ".join(parts) + ".")
        self._sync_inputs_to_paths()
        self._refresh_cards()

    def _sync_inputs_to_paths(self) -> None:
        """Keep the input list pointing at renamed files without losing entry state."""
        lines = []
        for e in self.entries:
            new_line = str(e.path) if e.path is not None else e.raw_line
            e.raw_line = new_line
            lines.append(new_line)
        self.inputs.blockSignals(True)
        self.inputs.set_paths(lines)
        self.inputs.blockSignals(False)

    def _save(self) -> None:
        try:
            config_save_settings(self._collect())
        except OSError:
            pass


@dataclass
class _Result:
    ok: bool
    summary: str


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
) -> None:
    ui.create_app("TranslateFilename", ACCENT)
    settings = config_load_settings()
    inputs_text = build_initial_inputs_text(settings.inputs_text, initial_only_list, initial_only_files)
    win = TranslateWindow(settings, inputs_text)
    launched_with_files = bool(initial_only_list or initial_only_files)
    win.restore_ui("translate" if launched_with_files else "",
                   "translate" if settings.active_api_key() else "settings")
    if launched_with_files and settings.active_api_key():
        QTimer.singleShot(250, win.translate)
    win.run_app()
