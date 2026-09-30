"""yt-dlp Downloader: settings, argument building, yt-dlp discovery and the download run.

yt-dlp is called with an argv list (no shell), so %(...)s output templates and
URLs with ? or & reach it untouched. The URL still goes through --batch-file:
when only a .bat/.cmd shim is available, cmd.exe would otherwise mangle it.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Callable, Optional

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))
from process_runner import ProcessRunner  # noqa: E402

OutputSink = Callable[[str, bool], None]

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "YtDlpTool"
CONFIG_PATH = CONFIG_DIR / "settings.json"
LEGACY_INI_PATH = Path(os.environ.get("APPDATA", "")) / "DOpus_ytdlp_settings.ini"
GITHUB_LATEST = "https://api.github.com/repos/yt-dlp/yt-dlp/releases/latest"


@dataclass
class Settings:
    audio: bool = True
    mp4_container: bool = False
    metadata: bool = False
    date_prefix: bool = True
    file_prefix: str = ""
    firefox_cookies: bool = False
    no_cookies: bool = False
    overwrite: bool = False
    update: bool = False
    impersonate: bool = False
    keep_console: bool = False
    dest: str = ""


@dataclass
class DownloadResult:
    ok: bool
    summary: str


# ── settings ──────────────────────────────────────────────────────────────


def _migrate_legacy_ini() -> dict:
    """Read the DOpus-script era key=value file (mode=0 audio / 1 video …)."""
    if not LEGACY_INI_PATH.is_file():
        return {}
    raw: dict[str, str] = {}
    try:
        for line in LEGACY_INI_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                raw[k.strip()] = v.strip()
    except OSError:
        return {}

    def flag(key: str, default: bool) -> bool:
        return raw.get(key, "1" if default else "0").strip() != "0"

    return {
        "audio": raw.get("mode", "0") != "1",
        "mp4_container": flag("mp4container", False),
        "metadata": flag("metadata", False),
        "date_prefix": flag("dateprefix", True),
        "file_prefix": raw.get("fileprefix", ""),
        "firefox_cookies": flag("cookies", False),
        "no_cookies": flag("nocookies", False),
        "overwrite": flag("overwrite", False),
        "update": flag("update", False),
        "impersonate": flag("impersonate", False),
        "keep_console": flag("keepps", False),
    }


def config_load_settings() -> Settings:
    data: dict = {}
    if CONFIG_PATH.is_file():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            data = {}
    else:
        data = _migrate_legacy_ini()
    s = Settings()
    for f in fields(Settings):
        if f.name in data:
            value = data[f.name]
            setattr(s, f.name, str(value) if f.type in ("str", str) else bool(value))
    return s


def config_save_settings(settings: Settings) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(asdict(settings), indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    except OSError:
        pass


# ── clipboard / command parsing ──────────────────────────────────────────


def tokenize_command_line(s: str) -> list[str]:
    """Split on whitespace, honouring "double" or 'single' quoted runs."""
    out: list[str] = []
    cur: list[str] = []
    quote = ""
    started = False
    for c in s:
        if c in "\"'" and (not quote or c == quote):
            quote = "" if quote else c
            started = True
        elif not quote and c in " \t\r\n":
            if started:
                out.append("".join(cur))
                cur, started = [], False
        else:
            cur.append(c)
            started = True
    if started:
        out.append("".join(cur))
    return out


def parse_ytdlp_command(text: str) -> Optional[tuple[str, str]]:
    """A full "yt-dlp … URL" command (e.g. from the YouTube Clipper extension) →
    (url, remaining args). None when the text is not such a command.
    -o/--output is dropped: this tool builds its own template."""
    t = (text or "").strip()
    if not re.match(r"^yt-dlp(\.exe)?[ \t]", t, re.I):
        return None
    toks = tokenize_command_line(t)
    url = ""
    extras: list[str] = []
    i = 1
    while i < len(toks):
        tok = toks[i]
        if tok in ("-o", "--output"):
            i += 2
            continue
        if re.match(r"^https?://", tok, re.I):
            url = tok
        else:
            extras.append(f'"{tok}"' if (" " in tok or not tok) else tok)
        i += 1
    return url, " ".join(extras)


def split_clipboard(text: str) -> tuple[str, str]:
    """Clipboard text → (url, extra args)."""
    parsed = parse_ytdlp_command(text)
    if parsed is not None:
        return parsed
    return (text or "").strip(), ""


def read_clipboard_text() -> str:
    """Clipboard text via Win32 (usable without a Qt application)."""
    if sys.platform != "win32":
        return ""
    import ctypes

    CF_UNICODETEXT = 13
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.GetClipboardData.restype = ctypes.c_void_p
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    if not user32.OpenClipboard(None):
        return ""
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return ""
        try:
            return ctypes.wstring_at(ptr)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


# ── argument building ─────────────────────────────────────────────────────


def escape_output_prefix(s: str) -> str:
    """Literal text placed before the %(…) template: % doubled, newlines flattened."""
    t = (s or "").strip()
    return t.replace("%", "%%").replace("\r", " ").replace("\n", " ")


def build_args(settings: Settings, extra_args: str, *, with_cookies: bool) -> list[str]:
    extra = tokenize_command_line(extra_args or "")
    sections = any(a.lower().startswith("--download-sections") for a in extra)
    section_suffix = " %(section_start)s-%(section_end)s" if sections else ""
    core = ("[%(upload_date>%m-%d-%Y)s] " if settings.date_prefix else "") + "%(title)s" + section_suffix + ".%(ext)s"
    args = ["-o", escape_output_prefix(settings.file_prefix) + core]
    if settings.audio:
        args += ["-f", "bestaudio"]
    args.append("--force-overwrites" if settings.overwrite else "--no-overwrites")
    # Lets yt-dlp fetch the EJS challenge solver scripts YouTube now needs.
    args += ["--remote-components", "ejs:github"]
    if settings.impersonate:
        args += ["--impersonate", "chrome"]
    if settings.no_cookies:
        args.append("--no-cookies")
    if not settings.audio:
        if settings.mp4_container:
            args += ["--merge-output-format", "mp4", "--remux-video", "mp4"]
        if sections:
            # Section downloads go through ffmpeg's single-stream path; cap at 1080p.
            args += ["-S", "res:1080"]
    if settings.metadata:
        # Subtitles always cover the whole video, so skip them for clips.
        if settings.audio:
            args += ["--extract-audio", "--audio-format", "best", "--add-metadata", "--embed-thumbnail"]
            if not sections:
                args.append("--embed-subs")
            args += ["--parse-metadata", ":(?P<chapters>)"]
        else:
            args += ["--add-metadata", "--embed-thumbnail"]
            if not sections:
                args += ["--write-auto-subs", "--embed-subs"]
    args += extra
    if with_cookies:
        args += ["--cookies-from-browser", "firefox"]
    return args


# ── yt-dlp discovery / update ─────────────────────────────────────────────


def _win_flags() -> int:
    return subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


def _python_beside(exe: str) -> Optional[str]:
    p = Path(exe).resolve().parent.parent / "python.exe"
    return str(p) if p.is_file() else None


def resolve_ytdlp() -> tuple[list[str], Optional[str]]:
    """(command prefix, python.exe that owns it). Prefers a real .exe so no cmd.exe
    shim touches the arguments; pyenv's shim is resolved through `pyenv which`."""
    try:
        r = subprocess.run(["pyenv", "which", "yt-dlp"], capture_output=True, text=True,
                           creationflags=_win_flags(), timeout=15, shell=(sys.platform == "win32"))
        exe = r.stdout.strip().splitlines()[0].strip() if r.returncode == 0 and r.stdout.strip() else ""
        if exe and Path(exe).is_file():
            return [exe], _python_beside(exe)
    except (OSError, subprocess.SubprocessError, IndexError):
        pass
    found = shutil.which("yt-dlp")
    if found and found.lower().endswith(".exe"):
        return [found], _python_beside(found)
    import importlib.util

    if importlib.util.find_spec("yt_dlp") is not None:
        py = sys.executable
        if py.lower().endswith("pythonw.exe"):
            alt = py[:-len("pythonw.exe")] + "python.exe"
            py = alt if Path(alt).is_file() else py
        return [py, "-m", "yt_dlp"], py
    return [found or "yt-dlp"], None


def _utf8_env() -> dict:
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


def local_version(cmd: list[str]) -> str:
    try:
        r = subprocess.run([*cmd, "--version"], capture_output=True, text=True, encoding="utf-8",
                           errors="replace", creationflags=_win_flags(), env=_utf8_env(), timeout=60)
        return r.stdout.strip().splitlines()[0].strip() if r.returncode == 0 and r.stdout.strip() else ""
    except (OSError, subprocess.SubprocessError, IndexError):
        return ""


def latest_version() -> str:
    try:
        req = urllib.request.Request(GITHUB_LATEST, headers={"User-Agent": "YtDlpTool"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            tag = json.loads(resp.read().decode("utf-8")).get("tag_name", "")
        return str(tag).lstrip("v").strip()
    except (OSError, ValueError):
        return ""


# ── running ───────────────────────────────────────────────────────────────

# yt-dlp redraws download progress on one line; keep it on one log line too.
_PROGRESS_RE = re.compile(r"^\[download\]\s+\d+(\.\d+)?%")


class Runner(ProcessRunner):
    def __init__(self, emit: OutputSink):
        super().__init__(emit, progress_re=_PROGRESS_RE, env=_utf8_env())


def update_ytdlp(runner: Runner, cmd: list[str], python: Optional[str]) -> None:
    emit = runner.emit
    local = local_version(cmd)
    latest = latest_version()
    if not latest:
        emit("Could not check the latest yt-dlp release on GitHub; skipping update.", False)
        return
    if local == latest:
        emit(f"yt-dlp is up to date ({local}).", False)
        return
    emit(f"yt-dlp update: {local or 'unknown'} → {latest}", False)
    py = python or "python"
    rc = runner.run([py, "-m", "pip", "install", "--upgrade", "yt-dlp"])
    if rc != 0:
        emit("Warning: pip upgrade failed; update yt-dlp manually (e.g. yt-dlp -U).", False)
    else:
        emit(f"yt-dlp now at {local_version(cmd) or latest}.", False)


def run_download(
    url: str,
    dest: str,
    settings: Settings,
    extra_args: str,
    runner: Runner,
) -> DownloadResult:
    emit = runner.emit
    url = (url or "").strip().replace("\r", " ").replace("\n", " ")
    if not url:
        return DownloadResult(False, "No URL.")
    if not dest or not Path(dest).is_dir():
        return DownloadResult(False, f"Destination folder does not exist:\n{dest}")
    cmd, python = resolve_ytdlp()
    emit(f"yt-dlp: {' '.join(cmd)}", False)
    emit(f"Saving to {dest}", False)
    if settings.update:
        update_ytdlp(runner, cmd, python)

    fd, url_file = tempfile.mkstemp(prefix="yt-dlp-url-", suffix=".txt")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(url + "\n")
        use_cookies = settings.firefox_cookies and not settings.no_cookies
        args = build_args(settings, extra_args, with_cookies=use_cookies)
        rc = runner.run([*cmd, *args, "--batch-file", url_file], cwd=dest)
        if rc != 0 and use_cookies and not runner.cancelled:
            emit("", False)
            emit(f"Warning: yt-dlp failed (exit {rc}); retrying without browser cookies…", False)
            rc = runner.run([*cmd, *build_args(settings, extra_args, with_cookies=False), "--batch-file", url_file],
                            cwd=dest)
    finally:
        try:
            os.remove(url_file)
        except OSError:
            pass
    if runner.cancelled:
        return DownloadResult(False, "Download cancelled.")
    if rc == 0:
        return DownloadResult(True, "Download finished.")
    return DownloadResult(False, f"yt-dlp exited with code {rc}.")


# ── headless (Ctrl+click in Opus) ─────────────────────────────────────────


def _move_console_bottom_right() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes

        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if not hwnd:
            return
        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
        work = wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(work), 0)  # SPI_GETWORKAREA
        w, h = rect.right - rect.left, rect.bottom - rect.top
        x, y = max(work.left, work.right - w), max(work.top, work.bottom - h)
        ctypes.windll.user32.SetWindowPos(hwnd, None, x, y, 0, 0, 0x0001 | 0x0004 | 0x0040)
    except (AttributeError, OSError):
        pass


def run_repeat(dest: str) -> int:
    """Download the clipboard URL with the saved settings, printing to this console."""
    _move_console_bottom_right()
    settings = config_load_settings()
    url, extra = split_clipboard(read_clipboard_text())
    if not url:
        print("No URL in the clipboard. Copy a URL or a yt-dlp command first.")
        input("Press Enter to close")
        return 2

    def emit(text: str, replace_last: bool) -> None:
        print(("\r" + text) if replace_last else ("\n" + text), end="", flush=True)

    result = run_download(url, dest or settings.dest, settings, extra, Runner(emit))
    print("\n\n" + result.summary)
    if settings.keep_console or not result.ok:
        input("Press Enter to close")
    return 0 if result.ok else 1
