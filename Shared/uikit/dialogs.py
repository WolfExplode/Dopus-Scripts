"""Themed message boxes (QMessageBox picks up the app stylesheet)."""

from __future__ import annotations

from typing import Optional

from PySide6.QtWidgets import QMessageBox, QWidget


def error(parent: Optional[QWidget], title: str, text: str) -> None:
    QMessageBox.critical(parent, title, text)


def info(parent: Optional[QWidget], title: str, text: str) -> None:
    QMessageBox.information(parent, title, text)


def confirm(parent: Optional[QWidget], title: str, text: str, yes: str = "Continue") -> bool:
    box = QMessageBox(QMessageBox.Question, title, text, parent=parent)
    ok = box.addButton(yes, QMessageBox.AcceptRole)
    ok.setProperty("kind", "primary")
    box.addButton("Cancel", QMessageBox.RejectRole)
    box.exec()
    return box.clickedButton() is ok


def ask_yes_no_cancel(
    parent: Optional[QWidget], title: str, text: str, yes: str = "Yes", no: str = "No"
) -> Optional[bool]:
    """True for yes, False for no, None when cancelled."""
    box = QMessageBox(QMessageBox.Question, title, text, parent=parent)
    y = box.addButton(yes, QMessageBox.YesRole)
    y.setProperty("kind", "primary")
    n = box.addButton(no, QMessageBox.NoRole)
    box.addButton("Cancel", QMessageBox.RejectRole)
    box.exec()
    clicked = box.clickedButton()
    if clicked is y:
        return True
    if clicked is n:
        return False
    return None
