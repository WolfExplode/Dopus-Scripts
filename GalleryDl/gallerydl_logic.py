"""gallery-dl Pinterest scraper: profiles, cached board lists and the download run.

Board lists are cached per profile (boards_<user>.txt) under
%APPDATA%\\DOpus_gallery_dl_pinterest, the same place the DOpus script used,
so existing caches keep working.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
import urllib.parse
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable, Optional

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))
from process_runner import ProcessRunner  # noqa: E402

OutputSink = Callable[[str, bool], None]

GALLERY_DL_EXE = r"C:\Users\WXP\AppData\Local\Programs\Python\Python310\Scripts\gallery-dl.exe"
# gallery-dl appends its own pinterest\<user>\<board> under this root.
OUTPUT_ROOT = r"C:\Users\WXP\Documents\Pureref\gallery-dl"

PROFILES: list[tuple[str, str]] = [
    ("allyfire1281", "https://www.pinterest.com/allyfire1281/"),
    ("FedTheBeast", "https://www.pinterest.com/FedTheBeast/"),
    ("fireally31", "https://www.pinterest.com/fireally31/"),
]

CACHE_DIR = Path(os.environ.get("APPDATA", "")) / "DOpus_gallery_dl_pinterest"
CONFIG_PATH = CACHE_DIR / "settings.json"
LEGACY_INI_PATH = CACHE_DIR / "settings.ini"


@dataclass
class Settings:
    profile: int = 0
    board_url: str = ""
    firefox_cookies: bool = True


@dataclass
class JobResult:
    ok: bool
    summary: str


# ── settings ──────────────────────────────────────────────────────────────


def config_load_settings() -> Settings:
    s = Settings()
    if CONFIG_PATH.is_file():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            s.profile = int(data.get("profile", 0))
            s.board_url = str(data.get("board_url", ""))
            s.firefox_cookies = bool(data.get("firefox_cookies", True))
        except (OSError, ValueError, TypeError):
            pass
    elif LEGACY_INI_PATH.is_file():
        try:
            for line in LEGACY_INI_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
                k, _, v = line.partition("=")
                if k == "profile":
                    s.profile = int(v or 0)
                elif k == "boardUrl":
                    s.board_url = v.strip()
                elif k == "cookies":
                    s.firefox_cookies = v.strip() != "0"
        except (OSError, ValueError):
            pass
    if not 0 <= s.profile < len(PROFILES):
        s.profile = 0
    return s


def config_save_settings(settings: Settings) -> None:
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(asdict(settings), indent=2) + "\n", encoding="utf-8")
    except OSError:
        pass


# ── boards ────────────────────────────────────────────────────────────────


def slug_from_url(url: str) -> str:
    last = url.strip().rstrip("/").split("/")[-1] or "user"
    return re.sub(r"[^\w\-]", "_", last)


def board_cache_path(profile_url: str) -> Path:
    return CACHE_DIR / f"boards_{slug_from_url(profile_url)}.txt"


def is_board_url(line: str) -> bool:
    line = line.strip().lstrip("\ufeff")
    if not line or line.startswith("#") or "pinterest." not in line.lower() or "/pin/" in line:
        return False
    m = re.search(r"pinterest\.[^/]+/[^/]+/([^/?#]+)/?$", line)
    if not m:
        return False
    return m.group(1).lower() not in ("pins", "_created", "_saved", "search", "ideas")


def board_label(url: str) -> str:
    m = re.search(r"/([^/?#]+)/?$", url.strip())
    if not m:
        return url
    return urllib.parse.unquote(m.group(1).replace("+", " "))


def read_board_cache(profile_url: str) -> list[str]:
    path = board_cache_path(profile_url)
    if not path.is_file():
        return []
    out: list[str] = []
    seen: set[str] = set()
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip().lstrip("\ufeff")
        if is_board_url(line) and line not in seen:
            seen.add(line)
            out.append(line)
    return out


def board_rows(profile_index: int) -> list[tuple[str, str]]:
    """(label, url) rows for the board picker: whole profile first, then cached boards."""
    profile_url = PROFILES[profile_index][1]
    rows = [("All boards (entire profile)", profile_url)]
    rows += [(board_label(u), u) for u in read_board_cache(profile_url)]
    return rows


# ── gallery-dl ────────────────────────────────────────────────────────────


def resolve_gallery_dl() -> Optional[str]:
    if Path(GALLERY_DL_EXE).is_file():
        return GALLERY_DL_EXE
    return shutil.which("gallery-dl")


def _cookie_args(cookies: bool) -> list[str]:
    return ["--cookies-from-browser", "firefox"] if cookies else []


def refresh_boards(profile_index: int, cookies: bool, runner: ProcessRunner) -> JobResult:
    exe = resolve_gallery_dl()
    if not exe:
        return JobResult(False, f"gallery-dl.exe not found at:\n{GALLERY_DL_EXE}")
    label, url = PROFILES[profile_index]
    found: list[str] = []

    def collect(line: str) -> None:
        if is_board_url(line) and line.strip() not in found:
            found.append(line.strip())

    runner.emit(f"Listing boards for {label}…", False)
    runner.run([exe, "-g", *_cookie_args(cookies), url], on_line=collect)
    if runner.cancelled:
        return JobResult(False, "Cancelled.")
    if not found:
        return JobResult(False, f"No board URLs found for {url}.\nCheck Firefox cookies, the network and "
                                "gallery-dl. The previous cache was kept.")
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    board_cache_path(url).write_text("\n".join(found) + "\n", encoding="utf-8")
    return JobResult(True, f"Saved {len(found)} board(s) for {label}.")


def run_scrape(url: str, cookies: bool, runner: ProcessRunner) -> JobResult:
    exe = resolve_gallery_dl()
    if not exe:
        return JobResult(False, f"gallery-dl.exe not found at:\n{GALLERY_DL_EXE}")
    runner.emit(f"Saving under {OUTPUT_ROOT}", False)
    rc = runner.run([exe, "-d", OUTPUT_ROOT, *_cookie_args(cookies), url])
    if runner.cancelled:
        return JobResult(False, "Scrape cancelled.")
    return JobResult(rc == 0, "Scrape finished." if rc == 0 else f"gallery-dl exited with code {rc}.")
