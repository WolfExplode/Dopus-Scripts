"""Video/audio conversion and ffmpeg utilities (ported from DOpus_ffmpeg.js)."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

OutputSink = Callable[[str, bool], None]

# --- extensions ---

THUMB_IMAGE_EXT = {
    ".jpg", ".jpeg", ".jfif", ".pjpeg", ".pjp", ".png", ".apng", ".webp", ".bmp",
    ".dib", ".gif", ".tif", ".tiff", ".heic", ".heif", ".avif", ".jxl",
}
THUMB_VIDEO_EXT = {
    ".mp4", ".m4v", ".mov", ".qt", ".mkv", ".webm", ".avi", ".wmv", ".asf",
    ".mpg", ".mpeg", ".mpe", ".m1v", ".m2v", ".mpv", ".vob", ".ts", ".mts",
    ".m2t", ".m2ts", ".3gp", ".3g2", ".flv", ".f4v", ".ogv", ".ogm", ".dv", ".mxf",
}
THUMB_AUDIO_EXT = {
    ".mp3", ".mp2", ".mpa", ".m4a", ".m4b", ".m4p", ".aac", ".adts", ".flac",
    ".wav", ".aiff", ".aif", ".aifc", ".caf", ".ogg", ".oga", ".opus", ".mka",
    ".wma", ".ac3", ".eac3", ".dts", ".amr", ".awb", ".au", ".snd", ".ape",
    ".tta", ".wv", ".weba",
}
STILL_REPLACE_FPS = "1"
STILL_REPLACE_CRF_X264 = "22"
STILL_REPLACE_CRF_VP9 = "22"
STILL_REPLACE_MAXRATE = "1M"
STILL_REPLACE_BUFSIZE = "2M"
STILL_REPLACE_KEYFRAME_X264 = "-g 999999 -keyint_min 999999 -sc_threshold 0 -x264-params scenecut=0"
STILL_REPLACE_KEYFRAME_VP9 = "-g 999999 -keyint_min 999999"
STILL_REPLACE_MAX_LONG_SIDE = 1920

LAST_ACTIONS = (
    "convert", "cover", "mono", "splitav", "splitch", "mergevid",
    "rotatecw", "rotateccw", "fliph", "flipv", "trimstart",
)
DEFAULT_LAST_ACTION = "convert"

GUI_SECTION_DEFAULTS: dict[str, bool] = {
    "files": True,
    "convert": True,
    "rotate": False,
    "trim": False,
    "cover": False,
    "audio": False,
    "merge": False,
}

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "FFmpegTool"
CONFIG_PATH = CONFIG_DIR / "settings.json"
LEGACY_INI_PATH = Path(os.environ.get("APPDATA", "")) / "DOpus_ffmpeg_settings.ini"


@dataclass(frozen=True)
class FormatPreset:
    name: str
    ext: str
    codec: str
    crf: bool = False


VIDEO_FORMATS: tuple[FormatPreset, ...] = (
    FormatPreset("MP4 H.264 (Fast)", ".mp4", "libx264 -crf 23 -preset fast -c:a aac -b:a 192k -pix_fmt yuv420p", True),
    FormatPreset("MP4 H.265/HEVC", ".mp4", "libx265 -crf 28 -preset fast -c:a aac -b:a 192k -pix_fmt yuv420p", True),
    FormatPreset("MP4 YouTube Ready", ".mp4", "libx264 -crf 23 -preset slow -c:a aac -b:a 256k -pix_fmt yuv420p -movflags +faststart", True),
    FormatPreset("MOV ProRes 422", ".mov", "prores -profile:v 2 -c:a pcm_s16le"),
    FormatPreset("MOV ProRes 4444", ".mov", "prores -profile:v 3 -alpha_bits 0 -c:a pcm_s16le"),
    FormatPreset("MOV H.264", ".mov", "libx264 -crf 23 -preset fast -c:a aac -b:a 192k -pix_fmt yuv420p", True),
    FormatPreset("WebM VP9", ".webm", "libvpx-vp9 -crf 30 -b:v 0 -c:a libopus -b:a 128k", True),
    FormatPreset("AVI Uncompressed", ".avi", "rawvideo -c:a pcm_s16le"),
)

AUDIO_FORMATS: tuple[FormatPreset, ...] = (
    FormatPreset("MP3 High Quality (320k)", ".mp3", "libmp3lame -q:a 0 -b:a 320k"),
    FormatPreset("MP3 Standard (192k)", ".mp3", "libmp3lame -q:a 2 -b:a 192k"),
    FormatPreset("MP3 Voice (64k)", ".mp3", "libmp3lame -q:a 6 -b:a 64k"),
    FormatPreset("FLAC Lossless", ".flac", "flac"),
    FormatPreset("WAV PCM 16-bit", ".wav", "pcm_s16le"),
    FormatPreset("WAV PCM 24-bit", ".wav", "pcm_s24le"),
    FormatPreset("AAC M4A", ".m4a", "aac -b:a 256k"),
    FormatPreset("OGG Vorbis", ".ogg", "libvorbis -q:a 6"),
    FormatPreset("OGG Opus", ".ogg", "libopus -b:a 128k"),
)


@dataclass
class ActionResult:
    ok: bool
    summary: str
    log: list[str] = field(default_factory=list)


@dataclass
class Settings:
    mode: int = 0
    format_name: str = ""
    quality: str = "23"
    last_action: str = DEFAULT_LAST_ACTION
    replace_video_with_image: bool = False
    trim_frames: str = "1"
    files_text: str = ""
    merge_fix_outliers: bool = True
    gui_sections: dict[str, bool] = field(default_factory=lambda: dict(GUI_SECTION_DEFAULTS))


def file_ext_lower(name: str) -> str:
    p = name.rfind(".")
    if p < 0:
        return ""
    return name[p:].lower()


def is_thumb_image(name: str) -> bool:
    return file_ext_lower(name) in THUMB_IMAGE_EXT


def is_thumb_video(name: str) -> bool:
    return file_ext_lower(name) in THUMB_VIDEO_EXT


def is_thumb_audio(name: str) -> bool:
    return file_ext_lower(name) in THUMB_AUDIO_EXT


def is_thumb_media(name: str) -> bool:
    return is_thumb_video(name) or is_thumb_audio(name)


def dedupe_path_lines(lines: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for ln in lines:
        key = ln.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(ln)
    return out


def parse_file_paths(text: str) -> list[Path]:
    paths: list[Path] = []
    for ln in text.splitlines():
        s = ln.strip()
        if not s:
            continue
        p = Path(s)
        if p.is_file():
            paths.append(p.resolve())
    return dedupe_paths(paths)


def dedupe_paths(paths: list[Path]) -> list[Path]:
    seen: set[str] = set()
    out: list[Path] = []
    for p in paths:
        key = os.fspath(p).casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def paths_from_only_list(list_path: Optional[str], only_files: Optional[list[str]]) -> list[Path]:
    lines: list[str] = []
    if list_path:
        try:
            lines.extend(Path(list_path).read_text(encoding="utf-8").splitlines())
        except OSError:
            pass
    if only_files:
        lines.extend(only_files)
    return parse_file_paths("\n".join(lines))


def build_initial_files_text(
    saved: str,
    only_list: Optional[str],
    only_files: Optional[list[str]],
) -> str:
    from_dopus = paths_from_only_list(only_list, only_files)
    if from_dopus:
        return "\n".join(os.fspath(p) for p in from_dopus)
    return saved.strip()


def _migrate_legacy_ini() -> Optional[dict]:
    if not LEGACY_INI_PATH.is_file():
        return None
    out: dict = {}
    try:
        for line in LEGACY_INI_PATH.read_text(encoding="utf-8").splitlines():
            line = line.rstrip("\r")
            if "=" not in line:
                continue
            key, val = line.split("=", 1)
            out[key.strip()] = val.strip()
    except OSError:
        return None
    return out


def config_read() -> dict:
    if CONFIG_PATH.is_file():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            pass
    legacy = _migrate_legacy_ini()
    if legacy:
        return {
            "mode": int(legacy.get("mode", 0) or 0),
            "format_name": legacy.get("formatName", ""),
            "quality": legacy.get("quality", "23"),
            "last_action": legacy.get("lastAction", DEFAULT_LAST_ACTION),
            "replace_video_with_image": legacy.get("replaceVideoWithImage") == "1",
            "trim_frames": legacy.get("trimFrames", "1"),
        }
    return {}


def config_write(data: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def config_load_settings() -> Settings:
    data = config_read()
    sections = data.get("gui_sections") or {}
    gui_sections = dict(GUI_SECTION_DEFAULTS)
    if isinstance(sections, dict):
        for k in GUI_SECTION_DEFAULTS:
            if k in sections:
                gui_sections[k] = bool(sections[k])
    action = str(data.get("last_action") or DEFAULT_LAST_ACTION)
    if action not in LAST_ACTIONS:
        action = DEFAULT_LAST_ACTION
    return Settings(
        mode=int(data.get("mode", 0) or 0),
        format_name=str(data.get("format_name") or ""),
        quality=str(data.get("quality") or "23") or "23",
        last_action=action,
        replace_video_with_image=bool(data.get("replace_video_with_image")),
        trim_frames=str(data.get("trim_frames") or "1") or "1",
        files_text=str(data.get("files_text") or ""),
        merge_fix_outliers=bool(data.get("merge_fix_outliers", True)),
        gui_sections=gui_sections,
    )


def config_save_settings(
    settings: Settings,
    *,
    last_action: Optional[str] = None,
) -> None:
    data = config_read()
    data["mode"] = settings.mode
    data["format_name"] = settings.format_name
    data["quality"] = settings.quality
    data["replace_video_with_image"] = settings.replace_video_with_image
    data["trim_frames"] = settings.trim_frames
    data["files_text"] = settings.files_text
    data["merge_fix_outliers"] = settings.merge_fix_outliers
    data["gui_sections"] = settings.gui_sections
    if last_action:
        if last_action in LAST_ACTIONS:
            data["last_action"] = last_action
    else:
        data["last_action"] = settings.last_action
    try:
        config_write(data)
    except OSError:
        pass


def config_load_last_action() -> str:
    action = config_load_settings().last_action
    return action if action in LAST_ACTIONS else DEFAULT_LAST_ACTION


def _quote(path: Path | str) -> str:
    return f'"{path}"'


_shutting_down = False
_cleanup_done = False
_active_procs: list[subprocess.Popen] = []
_registered_temps: list[Path] = []
_job_outputs: dict[Path, bool] = {}
_output_sink: OutputSink | None = None


class JobLog(list):
    def append(self, item) -> None:  # type: ignore[override]
        super().append(item)
        if _output_sink is not None and isinstance(item, str):
            _output_sink(item, False)


def _job_log() -> JobLog:
    return JobLog()


def _emit_output(text: str, *, replace_last: bool = False) -> None:
    if _output_sink is not None:
        _output_sink(text, replace_last)


def _cancelled() -> bool:
    return _shutting_down


def register_temp_path(path: Path | str) -> None:
    p = Path(path)
    if p not in _registered_temps:
        _registered_temps.append(p)


def begin_job(*, on_output: OutputSink | None = None) -> None:
    global _shutting_down, _output_sink
    _shutting_down = False
    _output_sink = on_output
    _job_outputs.clear()


def _win_subprocess_flags() -> int:
    if sys.platform != "win32":
        return 0
    return subprocess.CREATE_NO_WINDOW


def register_job_output(path: Path | str) -> None:
    _job_outputs[Path(path)] = False


def complete_job_output(path: Path | str) -> None:
    _job_outputs[Path(path)] = True


def _cleanup_incomplete_job_outputs() -> None:
    for path, done in list(_job_outputs.items()):
        if not done:
            _safe_delete(path)
    _job_outputs.clear()


def _finish_job_outputs(outputs: list[Path], rc: int) -> None:
    if not outputs:
        return
    if rc == 0 and not _cancelled():
        for p in outputs:
            complete_job_output(p)
    else:
        for p in outputs:
            _safe_delete(p)
            _job_outputs.pop(p, None)


def _kill_process_tree(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    if sys.platform == "win32":
        subprocess.run(
            f"taskkill /F /T /PID {proc.pid}",
            shell=True,
            capture_output=True,
            check=False,
            creationflags=_win_subprocess_flags(),
        )
    else:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


def _stop_all_processes() -> None:
    for proc in list(_active_procs):
        _kill_process_tree(proc)
    for proc in list(_active_procs):
        try:
            proc.wait(timeout=2)
        except subprocess.TimeoutExpired:
            pass
    _active_procs.clear()


def _original_path_from_bak(bak: Path) -> Optional[Path]:
    name = bak.name
    marker = ".__opus_"
    i = name.find(marker)
    orig_tag = "_orig"
    j = name.rfind(orig_tag)
    if i < 0 or j < 0:
        return None
    stem = name[:i]
    ext = name[j + len(orig_tag) :]
    if not ext.startswith("."):
        return None
    return bak.parent / f"{stem}{ext}"


def _restore_interrupted_in_place_files(search_dirs: set[Path]) -> None:
    for folder in search_dirs:
        if not folder.is_dir():
            continue
        for bak in folder.glob("*.__opus_*_orig*"):
            original = _original_path_from_bak(bak)
            if not original or not bak.is_file():
                continue
            if not original.is_file():
                try:
                    shutil.move(bak, original)
                except OSError:
                    pass
            else:
                _safe_delete(bak)


def _cleanup_opus_temp_files(search_dirs: set[Path]) -> None:
    for folder in search_dirs:
        if not folder.is_dir():
            continue
        for p in folder.glob("*.__opus_*"):
            _safe_delete(p)


def _cleanup_appdata_temp_files() -> None:
    temp = Path(os.environ.get("TEMP", "."))
    if not temp.is_dir():
        return
    for pattern in (
        "DOpus_ffmpeg_merge_*.ffmeta",
        "DOpus_ffmpeg_merge_*.concat.txt",
        "DOpus_ffmpeg_merge_*_fixed.*",
        "DOpus_ffmpeg_chfilt_*.txt",
        "FFmpegTool_only_*.txt",
    ):
        for p in temp.glob(pattern):
            _safe_delete(p)


def _delete_only_list_file(only_list: Optional[str]) -> None:
    if not only_list:
        return
    p = Path(only_list)
    name = p.name
    if name.startswith("FFmpegTool_only_") and name.endswith(".txt"):
        _safe_delete(p)


def cancel_running_jobs() -> None:
    global _shutting_down
    _shutting_down = True
    _stop_all_processes()
    _cleanup_incomplete_job_outputs()


def shutdown_ffmpeg_tool(
    *,
    paths: Optional[list[Path]] = None,
    only_list: Optional[str] = None,
) -> None:
    global _cleanup_done
    if _cleanup_done:
        return
    _cleanup_done = True
    cancel_running_jobs()

    search_dirs: set[Path] = set()
    for p in paths or []:
        search_dirs.add(p.parent)
    for p in _registered_temps:
        search_dirs.add(p.parent)

    _restore_interrupted_in_place_files(search_dirs)
    _cleanup_opus_temp_files(search_dirs)
    for p in list(_registered_temps):
        _safe_delete(p)
    _registered_temps.clear()
    _cleanup_incomplete_job_outputs()
    _cleanup_appdata_temp_files()
    _delete_only_list_file(only_list)


def _stream_process_stderr(proc: subprocess.Popen) -> None:
    if proc.stderr is None:
        return
    carry = ""
    try:
        while True:
            chunk = proc.stderr.read(4096)
            if not chunk:
                break
            carry += chunk.decode("utf-8", errors="replace")
            while carry:
                r = carry.find("\r")
                n = carry.find("\n")
                if r == -1 and n == -1:
                    break
                if n != -1 and (r == -1 or n < r):
                    line, carry = carry[:n], carry[n + 1 :]
                    line = line.strip("\r")
                    if line:
                        _emit_output(line, replace_last=False)
                else:
                    line, carry = carry[:r], carry[r + 1 :]
                    if line:
                        _emit_output(line, replace_last=True)
        tail = carry.strip("\r\n")
        if tail:
            _emit_output(tail, replace_last=False)
    except OSError:
        pass


def _run_cmd(
    cmd: str,
    log: list[str],
    *,
    job_output: Optional[Path | str] = None,
    job_outputs: Optional[list[Path | str]] = None,
) -> int:
    if _cancelled():
        log.append("Cancelled.")
        return -1

    outputs: list[Path] = []
    if job_output is not None:
        outputs.append(Path(job_output))
    if job_outputs:
        outputs.extend(Path(p) for p in job_outputs)
    for p in outputs:
        register_job_output(p)

    log.append(cmd)
    try:
        popen_kwargs: dict = {
            "shell": True,
            "stderr": subprocess.PIPE,
            "stdout": subprocess.DEVNULL,
            "creationflags": _win_subprocess_flags(),
        }
        proc = subprocess.Popen(cmd, **popen_kwargs)
        _active_procs.append(proc)
        reader = threading.Thread(target=_stream_process_stderr, args=(proc,), daemon=True)
        reader.start()
        try:
            rc = proc.wait()
        finally:
            if proc in _active_procs:
                _active_procs.remove(proc)
            if proc.stderr is not None:
                try:
                    proc.stderr.close()
                except OSError:
                    pass
            reader.join(timeout=2.0)
    except OSError as ex:
        log.append(f"Error: {ex}")
        _finish_job_outputs(outputs, -1)
        return -1

    _finish_job_outputs(outputs, rc)
    return rc


def _ffprobe_line(media_path: Path, args: list[str]) -> str:
    cmd = ["ffprobe.exe", "-v", "error", *args, os.fspath(media_path)]
    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            creationflags=_win_subprocess_flags(),
        )
        if r.returncode != 0:
            return ""
        return r.stdout.strip().replace("\r", "").replace("\n", "")
    except OSError:
        return ""


def probe_first_audio_codec(media_path: Path) -> str:
    return _ffprobe_line(
        media_path,
        ["-select_streams", "a:0", "-show_entries", "stream=codec_name", "-of", "csv=p=0"],
    )


def probe_audio_channel_count(media_path: Path) -> int:
    line = _ffprobe_line(
        media_path,
        ["-select_streams", "a:0", "-show_entries", "stream=channels", "-of", "csv=p=0"],
    )
    try:
        n = int(line)
        return n if n >= 1 else -1
    except ValueError:
        return -1


def probe_media_duration_sec(media_path: Path) -> float:
    line = _ffprobe_line(
        media_path,
        ["-show_entries", "format=duration", "-of", "csv=p=0"],
    )
    try:
        f = float(line)
        return f if f > 0 else -1.0
    except ValueError:
        return -1.0


def _parse_video_dimensions(line: str) -> Optional[tuple[int, int]]:
    parts = line.split("x")
    if len(parts) < 2:
        return None
    try:
        w, h = int(parts[0]), int(parts[1])
    except ValueError:
        return None
    if w <= 0 or h <= 0:
        return None
    if w % 2:
        w += 1
    if h % 2:
        h += 1
    return w, h


def clamp_dims_longest_side(w: int, h: int, max_side: int = STILL_REPLACE_MAX_LONG_SIDE) -> tuple[int, int]:
    longest = max(w, h)
    if longest <= max_side:
        return w, h
    scale = max_side / longest
    w = int(w * scale)
    h = int(h * scale)
    if w % 2:
        w -= 1
    if h % 2:
        h -= 1
    return max(w, 2), max(h, 2)


def probe_image_dimensions(img_path: Path) -> Optional[tuple[int, int]]:
    line = _ffprobe_line(
        img_path,
        ["-select_streams", "v:0", "-show_entries", "stream=width,height", "-of", "csv=p=0:s=x"],
    )
    return _parse_video_dimensions(line)


def probe_video_dimensions(media_path: Path) -> Optional[tuple[int, int]]:
    line = _ffprobe_line(
        media_path,
        ["-select_streams", "v:0", "-show_entries", "stream=width,height", "-of", "csv=p=0:s=x"],
    )
    return _parse_video_dimensions(line)


@dataclass
class MergeStreamInfo:
    duration: float
    width: int
    height: int
    video_codec: str
    pix_fmt: str
    audio_codec: str
    audio_sample_rate: int
    audio_channels: int


def probe_merge_stream_info(media_path: Path) -> Optional[MergeStreamInfo]:
    dur = probe_media_duration_sec(media_path)
    if dur <= 0:
        return None
    v_line = _ffprobe_line(
        media_path,
        [
            "-select_streams", "v:0",
            "-show_entries", "stream=codec_name,width,height,pix_fmt",
            "-of", "csv=p=0",
        ],
    )
    if not v_line:
        return None
    parts = v_line.split(",")
    if len(parts) < 4:
        return None
    try:
        width, height = int(parts[1]), int(parts[2])
    except ValueError:
        return None
    if width <= 0 or height <= 0:
        return None

    audio_codec = probe_first_audio_codec(media_path)
    audio_sample_rate = 0
    audio_channels = 0
    if audio_codec:
        rate_line = _ffprobe_line(
            media_path,
            ["-select_streams", "a:0", "-show_entries", "stream=sample_rate", "-of", "csv=p=0"],
        )
        try:
            audio_sample_rate = int(rate_line)
        except ValueError:
            return None
        audio_channels = probe_audio_channel_count(media_path)
        if audio_channels < 1:
            return None

    return MergeStreamInfo(
        duration=dur,
        width=width,
        height=height,
        video_codec=parts[0],
        pix_fmt=parts[3],
        audio_codec=audio_codec,
        audio_sample_rate=audio_sample_rate,
        audio_channels=audio_channels,
    )


def merge_lossless_compatible(infos: list[MergeStreamInfo], ext: str, paths: list[Path]) -> bool:
    if any(file_ext_lower(p.name) != ext for p in paths):
        return False
    ref = infos[0]
    for info in infos[1:]:
        if info.video_codec != ref.video_codec:
            return False
        if info.width != ref.width or info.height != ref.height:
            return False
        if info.pix_fmt != ref.pix_fmt:
            return False
        if bool(info.audio_codec) != bool(ref.audio_codec):
            return False
        if ref.audio_codec:
            if info.audio_codec != ref.audio_codec:
                return False
            if info.audio_sample_rate != ref.audio_sample_rate:
                return False
            if info.audio_channels != ref.audio_channels:
                return False
    return True


def merge_stream_signature(info: MergeStreamInfo) -> tuple:
    return (
        info.video_codec,
        info.width,
        info.height,
        info.pix_fmt,
        info.audio_codec,
        info.audio_sample_rate,
        info.audio_channels,
    )


def merge_majority_signature(infos: list[MergeStreamInfo]) -> tuple:
    return Counter(merge_stream_signature(info) for info in infos).most_common(1)[0][0]


def merge_info_from_signature(sig: tuple) -> MergeStreamInfo:
    return MergeStreamInfo(
        duration=0.0,
        width=sig[1],
        height=sig[2],
        video_codec=sig[0],
        pix_fmt=sig[3],
        audio_codec=sig[4],
        audio_sample_rate=sig[5],
        audio_channels=sig[6],
    )


def merge_stream_diffs(info: MergeStreamInfo, target: MergeStreamInfo) -> list[str]:
    diffs: list[str] = []
    if info.video_codec != target.video_codec:
        diffs.append(f"video codec {info.video_codec} (expected {target.video_codec})")
    if info.width != target.width or info.height != target.height:
        diffs.append(f"size {info.width}x{info.height} (expected {target.width}x{target.height})")
    if info.pix_fmt != target.pix_fmt:
        diffs.append(f"pixel format {info.pix_fmt} (expected {target.pix_fmt})")
    if bool(info.audio_codec) != bool(target.audio_codec):
        got = info.audio_codec or "none"
        want = target.audio_codec or "none"
        diffs.append(f"audio {got} (expected {want})")
    elif target.audio_codec:
        if info.audio_codec != target.audio_codec:
            diffs.append(f"audio codec {info.audio_codec} (expected {target.audio_codec})")
        if info.audio_sample_rate != target.audio_sample_rate:
            diffs.append(
                f"audio rate {info.audio_sample_rate} (expected {target.audio_sample_rate})"
            )
        if info.audio_channels != target.audio_channels:
            diffs.append(
                f"audio channels {info.audio_channels} (expected {target.audio_channels})"
            )
    return diffs


def merge_stream_mismatch_warnings(
    paths: list[Path], infos: list[MergeStreamInfo]
) -> tuple[list[str], MergeStreamInfo, list[int]]:
    target_sig = merge_majority_signature(infos)
    target = merge_info_from_signature(target_sig)
    outlier_idxs = [
        i for i, info in enumerate(infos) if merge_stream_signature(info) != target_sig
    ]
    if not outlier_idxs:
        return [], target, outlier_idxs

    lines = [
        "Warning: streams do not match - lossless merge is not possible as-is.",
        (
            f"Majority profile: {target.width}x{target.height} {target.video_codec}, "
            f"audio {target.audio_codec or 'none'}"
            + (
                f" {target.audio_sample_rate} Hz {target.audio_channels}ch"
                if target.audio_codec
                else ""
            )
            + f". {len(outlier_idxs)} outlier(s):"
        ),
    ]
    for i in outlier_idxs:
        diffs = merge_stream_diffs(infos[i], target)
        lines.append(f"  {paths[i].name}: {', '.join(diffs)}")
    return lines, target, outlier_idxs


def merge_preflight_mismatch(
    paths: list[Path],
) -> Optional[tuple[str, bool]]:
    """Return popup message and whether outlier-only fix is possible, or None if lossless-ready."""
    if len(paths) < 2:
        return None
    sorted_paths = sorted(paths, key=lambda p: p.name.lower())
    if any(not is_thumb_video(p.name) for p in sorted_paths):
        return None

    infos: list[MergeStreamInfo] = []
    for p in sorted_paths:
        info = probe_merge_stream_info(p)
        if not info:
            return None
        infos.append(info)

    ext = file_ext_lower(sorted_paths[0].name)
    if merge_lossless_compatible(infos, ext, sorted_paths):
        return None

    warnings, target, outlier_idxs = merge_stream_mismatch_warnings(sorted_paths, infos)
    if not outlier_idxs:
        return None

    body = "\n".join(
        line for line in warnings if not line.startswith("Warning:")
    ).strip()
    body = f"Streams do not match.\n\n{body}"
    eligible = merge_outlier_fix_eligible(infos, target)
    if eligible:
        body += (
            "\n\nRe-encode only the mismatched file(s) and quick-merge the rest?\n\n"
            "Yes = fix outlier(s) + quick merge\n"
            "No = re-encode all files\n"
            "Cancel = abort merge"
        )
    else:
        body += (
            "\n\nOutlier-only fix is not available (e.g. missing/extra audio).\n\n"
            "Yes = re-encode all files\n"
            "No or Cancel = abort merge"
        )
    return body, eligible


def merge_outlier_fix_eligible(infos: list[MergeStreamInfo], target: MergeStreamInfo) -> bool:
    for info in infos:
        if merge_stream_signature(info) == merge_stream_signature(target):
            continue
        if bool(info.audio_codec) != bool(target.audio_codec):
            return False
    return True


def build_merge_outlier_fix_cmd(
    src: Path,
    dst: Path,
    info: MergeStreamInfo,
    target: MergeStreamInfo,
    ext: str,
    crf_str: str,
) -> str:
    w, h = target.width, target.height
    vf = (
        f"scale={w}:{h}:force_original_aspect_ratio=decrease,"
        f"pad={w}:{h}:(ow-iw)/2:(oh-ih)/2,setsar=1,format={target.pix_fmt}"
    )
    q = (crf_str or "").strip() or "23"
    cmd = (
        f"ffmpeg.exe -y -i {_quote(src)} -vf \"{vf}\" -map_metadata 0 "
        f"-map 0:v:0"
    )
    if target.audio_codec:
        cmd += " -map 0:a:0"
        audio_matches = (
            info.audio_codec == target.audio_codec
            and info.audio_sample_rate == target.audio_sample_rate
            and info.audio_channels == target.audio_channels
        )
        if audio_matches:
            cmd += " -c:a copy"
        else:
            cmd += (
                f" -af aresample={target.audio_sample_rate}:async=1:first_pts=0,"
                f"aformat=sample_rates={target.audio_sample_rate}:"
                f"channel_layouts={'stereo' if target.audio_channels == 2 else 'mono'}"
            )
            if ext == ".webm":
                cmd += " -c:a libopus -b:a 128k"
            else:
                cmd += " -c:a aac -b:a 192k"
    cmd += f" -c:v libx264 -crf {q} -preset ultrafast -pix_fmt {target.pix_fmt}"
    if ext in (".mp4", ".m4v"):
        cmd += " -movflags +faststart"
    cmd += f" {_quote(dst)}"
    return cmd


def build_lossless_merge_cmd(
    list_path: Path,
    meta_path: Path,
    out_path: Path,
    ext: str,
    has_audio: bool,
) -> str:
    cmd = (
        f"ffmpeg.exe -y -f concat -safe 0 -i {_quote(list_path)} -i {_quote(meta_path)}"
        f" -map 0:v:0"
    )
    if has_audio:
        cmd += " -map 0:a:0"
    cmd += " -map_metadata 1 -map_chapters 1 -c copy"
    if ext in (".mp4", ".m4v"):
        cmd += " -movflags +faststart"
    cmd += f" {_quote(out_path)}"
    return cmd


def write_concat_list_file(list_path: Path, paths: list[Path]) -> None:
    lines: list[str] = []
    for p in paths:
        escaped = p.as_posix().replace("'", "'\\''")
        lines.append(f"file '{escaped}'")
    list_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _parse_fps_rational(line: str) -> float:
    line = line.strip()
    if not line or line in ("0/0", "N/A"):
        return -1.0
    if "/" in line:
        num_s, den_s = line.split("/", 1)
        try:
            num, den = float(num_s), float(den_s)
            return num / den if den > 0 and num > 0 else -1.0
        except ValueError:
            return -1.0
    try:
        f = float(line)
        return f if f > 0 else -1.0
    except ValueError:
        return -1.0


def probe_video_avg_frame_rate(media_path: Path) -> float:
    line = _ffprobe_line(
        media_path,
        ["-select_streams", "v:0", "-show_entries", "stream=avg_frame_rate", "-of", "csv=p=0"],
    )
    return _parse_fps_rational(line)


def mime_type_for_image_ext(ext: str) -> str:
    mapping = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".jfif": "image/jpeg",
        ".pjpeg": "image/jpeg", ".pjp": "image/jpeg",
        ".png": "image/png", ".apng": "image/png", ".webp": "image/webp",
        ".gif": "image/gif", ".bmp": "image/bmp", ".dib": "image/bmp",
        ".tif": "image/tiff", ".tiff": "image/tiff",
        ".heic": "image/heif", ".heif": "image/heif", ".avif": "image/avif", ".jxl": "image/jxl",
    }
    return mapping.get(ext, "image/jpeg")


def unique_jpg_path_next_to_media(folder: Path, stem: str) -> Path:
    out = folder / f"{stem}.jpg"
    counter = 1
    while out.exists():
        out = folder / f"{stem}_{counter}.jpg"
        counter += 1
    return out


def _safe_delete(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def _replace_in_place(original: Path, tmp: Path, bak: Path, log: list[str], label: str) -> bool:
    try:
        if bak.exists():
            _safe_delete(bak)
        if original.exists():
            shutil.move(original, bak)
    except OSError:
        log.append(f"{label}: could not rename original (in use?): {original.name} — left temp: {tmp}")
        return False
    try:
        shutil.move(tmp, original)
    except OSError:
        try:
            if bak.exists():
                shutil.move(bak, original)
        except OSError:
            pass
        log.append(f"{label}: could not replace file, restored original: {original.name}")
        return False
    _safe_delete(bak)
    return True


def extract_cover_via_video_streams(
    media_path: Path, out_path: Path, prefer_second: bool, log: list[str]
) -> bool:
    order = ["0:v:1", "0:v:0"] if prefer_second else ["0:v:0", "0:v:1"]
    for stream in order:
        _safe_delete(out_path)
        cmd = (
            f"ffmpeg.exe -y -i {_quote(media_path)} -map {stream} "
            f"-frames:v 1 -q:v 2 {_quote(out_path)}"
        )
        if _run_cmd(cmd, log) == 0 and out_path.is_file():
            return True
    return False


def extract_cover_matroska_attachment(
    media_path: Path, folder: Path, stem: str, out_path: Path, log: list[str]
) -> bool:
    raw = folder / f"{stem}.__opus_cover_raw"
    _safe_delete(raw)
    dump = f'ffmpeg.exe -y -dump_attachment:t:0 {_quote(raw)} -i {_quote(media_path)}'
    dump_exit = _run_cmd(dump, log)
    if not raw.is_file():
        log.append(f"Extract thumbnail: no Matroska attachment t:0 (exit {dump_exit})")
        return False
    _safe_delete(out_path)
    conv = f'ffmpeg.exe -y -i {_quote(raw)} -frames:v 1 -q:v 2 {_quote(out_path)}'
    conv_exit = _run_cmd(conv, log)
    _safe_delete(raw)
    return conv_exit == 0 and out_path.is_file()


def try_extract_cover_to_path(
    media_path: Path, media_name: str, folder: Path, stem: str, ext: str, out_path: Path, log: list[str]
) -> bool:
    if ext in (".mkv", ".mka"):
        if extract_cover_matroska_attachment(media_path, folder, stem, out_path, log):
            return True
        return extract_cover_via_video_streams(media_path, out_path, False, log)
    if is_thumb_video(media_name):
        return extract_cover_via_video_streams(media_path, out_path, True, log)
    return extract_cover_via_video_streams(media_path, out_path, False, log)


@dataclass(frozen=True)
class CoverEmbedPlan:
    out_ext: str
    strategy: str  # mp3 | m4a | mkv_attach | attached_pic
    remux: bool


def _is_pcm_codec(codec: str) -> bool:
    return codec.startswith("pcm_")


def cover_embed_plan(media_path: Path) -> tuple[Optional[CoverEmbedPlan], str]:
    """Pick a cover-capable container that can keep the audio stream copied (-c:a copy)."""
    ext = file_ext_lower(media_path.name)
    codec = probe_first_audio_codec(media_path)
    if not codec:
        return None, f"No audio stream — cannot embed cover:\n{media_path.name}"

    if _is_pcm_codec(codec):
        return None, (
            f"Cannot embed cover in {ext} without re-encoding audio.\n"
            f"Use Convert to FLAC/M4A first, or keep a sidecar .jpg next to the file."
        )

    if codec in ("mp3", "mp2"):
        out_ext = ".mp3"
        return CoverEmbedPlan(out_ext, "mp3", out_ext != ext), ""

    if codec in ("aac", "alac"):
        out_ext = ".m4a"
        return CoverEmbedPlan(out_ext, "m4a", out_ext != ext), ""

    if codec == "flac":
        return CoverEmbedPlan(".flac", "attached_pic", ext != ".flac"), ""

    if codec in ("vorbis", "opus"):
        if ext in (".ogg", ".oga", ".opus", ".ogv", ".ogm"):
            out_ext = ext
        elif ext == ".weba":
            out_ext = ".weba"
        else:
            out_ext = ".ogg"
        return CoverEmbedPlan(out_ext, "attached_pic", out_ext != ext), ""

    if codec in ("ac3", "eac3", "dts"):
        out_ext = ".mka" if ext not in (".mkv", ".mka") else ext
        return CoverEmbedPlan(out_ext, "mkv_attach", out_ext != ext), ""

    if codec in ("wmav1", "wmav2", "wmapro", "wmalossless"):
        out_ext = ext if ext in (".wma", ".asf") else ".wma"
        return CoverEmbedPlan(out_ext, "attached_pic", out_ext != ext), ""

    if codec in ("ape", "tta", "wavpack", "mlp", "truehd"):
        return CoverEmbedPlan(ext, "attached_pic", False), ""

    if ext in (".mkv", ".mka"):
        return CoverEmbedPlan(ext, "mkv_attach", False), ""

    return CoverEmbedPlan(ext, "attached_pic", False), ""


def ffmpeg_m4a_cover_embed(media_path: Path, img_path: Path, tmp_path: Path) -> str:
    return (
        f"ffmpeg.exe -y -i {_quote(media_path)} -i {_quote(img_path)} "
        f"-map_metadata 0 -map_chapters 0 -map 0:a? -map 1:0 -c copy "
        f"-c:v:0 mjpeg -disposition:v:0 attached_pic {_quote(tmp_path)}"
    )


def ffmpeg_mp3_cover_embed(media_path: Path, img_path: Path, tmp_path: Path) -> str:
    return (
        f"ffmpeg.exe -y -i {_quote(media_path)} -i {_quote(img_path)} "
        f"-map_metadata 0 -map 0:a? -map 1 -c:a copy -id3v2_version 3 "
        f'-metadata:s:v title="Album cover" -metadata:s:v comment="Cover (front)" '
        f"{_quote(tmp_path)}"
    )


def ffmpeg_attached_pic_copy_embed(media_path: Path, img_path: Path, tmp_path: Path) -> str:
    return (
        f"ffmpeg.exe -y -i {_quote(media_path)} -i {_quote(img_path)} "
        f"-map_metadata 0 -map_chapters 0 -map 0:a? -map 1:0 -c copy "
        f"-c:v:1 mjpeg -disposition:v:1 attached_pic {_quote(tmp_path)}"
    )


def ffmpeg_mkv_cover_attach(
    media_path: Path, img_path: Path, img_ext: str, tmp_path: Path, *, audio_only: bool
) -> str:
    mime = mime_type_for_image_ext(img_ext)
    if audio_only:
        return (
            f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map_chapters 0 "
            f"-map 0:a? -c copy -attach {_quote(img_path)} "
            f"-metadata:s:t mimetype={mime} {_quote(tmp_path)}"
        )
    return (
        f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map_chapters 0 "
        f"-map 0 -map -0:t -c copy -attach {_quote(img_path)} "
        f"-metadata:s:t mimetype={mime} {_quote(tmp_path)}"
    )


def build_cover_embed_cmd(
    media_path: Path,
    img_path: Path,
    img_ext: str,
    tmp_path: Path,
    *,
    is_video_ext: bool,
    plan: Optional[CoverEmbedPlan],
) -> str:
    if is_video_ext:
        return (
            f"ffmpeg.exe -y -i {_quote(media_path)} -i {_quote(img_path)} "
            f"-map_metadata 0 -map_chapters 0 -map 0:v:0 -map 0:a? -map 0:s? "
            f"-map 0:d? -map 0:t? -map 1 -c copy -c:v:1 mjpeg "
            f"-disposition:v:1 attached_pic {_quote(tmp_path)}"
        )
    if plan is None:
        raise ValueError("cover embed plan required for audio")

    if plan.strategy == "mp3":
        return ffmpeg_mp3_cover_embed(media_path, img_path, tmp_path)
    if plan.strategy == "m4a":
        return ffmpeg_m4a_cover_embed(media_path, img_path, tmp_path)
    if plan.strategy == "mkv_attach":
        return ffmpeg_mkv_cover_attach(media_path, img_path, img_ext, tmp_path, audio_only=True)
    return ffmpeg_attached_pic_copy_embed(media_path, img_path, tmp_path)


def still_image_video_encode_args(ext: str) -> str:
    cap = f"-maxrate {STILL_REPLACE_MAXRATE} -bufsize {STILL_REPLACE_BUFSIZE}"
    if ext == ".webm":
        return f"libvpx-vp9 -pix_fmt yuv420p -crf {STILL_REPLACE_CRF_VP9} -b:v 0 {cap} {STILL_REPLACE_KEYFRAME_VP9}"
    if ext in (".ogv", ".ogm"):
        return f"libtheora -q:v 25 {cap}"
    if ext in (".wmv", ".asf"):
        return f"wmv2 -b:v 500k {cap}"
    return (
        f"libx264 -tune stillimage -pix_fmt yuv420p -crf {STILL_REPLACE_CRF_X264} "
        f"{cap} {STILL_REPLACE_KEYFRAME_X264}"
    )


def ffmpeg_replace_video_with_image_exec(
    media_path: Path, img_path: Path, tmp_path: Path, width: int, height: int, out_ext: str
) -> str:
    v_enc = still_image_video_encode_args(out_ext)
    scale = f"scale={width}:{height}"
    return (
        f"ffmpeg.exe -y -framerate {STILL_REPLACE_FPS} -loop 1 -i {_quote(img_path)} "
        f"-i {_quote(media_path)} -map_metadata 1 -map_chapters 1 -map 0:v:0 -map 1:a "
        f"-vf {scale} -r {STILL_REPLACE_FPS} -c:v {v_enc} -c:a copy -shortest {_quote(tmp_path)}"
    )


def thumb_embed_cover(
    media_path: Path, img_path: Path, img_ext: str, log: list[str]
) -> tuple[bool, str, Path | None]:
    name = media_path.name
    folder = media_path.parent
    stem = media_path.stem
    ext = file_ext_lower(name)
    is_video_ext = is_thumb_video(name)
    if is_video_ext:
        out_ext = ext
        remux = False
        plan = None
    else:
        plan, plan_err = cover_embed_plan(media_path)
        if not plan:
            return False, plan_err, None
        out_ext = plan.out_ext
        remux = plan.remux
    out_path = folder / f"{stem}{out_ext}"
    tmp = folder / f"{stem}.__opus_thumb_tmp{out_ext}"
    bak = folder / f"{stem}.__opus_thumb_orig{ext}"
    final = out_path if remux else media_path

    if remux and out_path.exists():
        return False, f"Cannot remux to {out_ext} — file already exists:\n{out_path}", None

    _safe_delete(tmp)
    _safe_delete(bak)
    if remux:
        log.append(f"Embed cover: remux {ext} to {out_ext} (audio copy)")

    cmd = build_cover_embed_cmd(
        media_path, img_path, img_ext, tmp, is_video_ext=is_video_ext, plan=plan
    )
    if _run_cmd(cmd, log) != 0 or not tmp.is_file():
        return False, "ffmpeg failed or output missing.", None

    if remux:
        try:
            if media_path.exists():
                shutil.move(media_path, bak)
            shutil.move(tmp, final)
            _safe_delete(bak)
        except OSError as ex:
            return False, f"Could not replace media file: {ex}", None
    elif not _replace_in_place(media_path, tmp, bak, log, "Embed cover"):
        return False, "Could not replace media file.", None

    return True, "", final


def thumb_replace_video_with_image(
    media_path: Path, img_path: Path, log: list[str]
) -> tuple[bool, str, Path | None]:
    name = media_path.name
    folder = media_path.parent
    stem = media_path.stem
    ext = file_ext_lower(name)
    out_ext = ext
    remux = False
    dims = probe_image_dimensions(img_path)
    if not dims:
        return False, "Could not read image width/height (ffprobe).", None
    if not probe_first_audio_codec(media_path):
        return False, f"No audio stream in media file — cannot build still-image video.\n\n{name}", None

    out_path = folder / f"{stem}{out_ext}"
    tmp = folder / f"{stem}.__opus_still_tmp{out_ext}"
    bak = folder / f"{stem}.__opus_still_orig{ext}"
    final = out_path if remux else media_path
    w, h = clamp_dims_longest_side(*dims)

    if remux and out_path.exists():
        return False, f"Cannot remux to {out_ext} — file already exists:\n{out_path}", None

    _safe_delete(tmp)
    _safe_delete(bak)
    cmd = ffmpeg_replace_video_with_image_exec(media_path, img_path, tmp, w, h, out_ext)
    if _run_cmd(cmd, log) != 0 or not tmp.is_file():
        return False, "ffmpeg failed or output missing.", None

    if remux:
        try:
            if media_path.exists():
                shutil.move(media_path, bak)
            shutil.move(tmp, final)
            _safe_delete(bak)
        except OSError as ex:
            return False, f"Could not replace media file: {ex}", None
    elif not _replace_in_place(media_path, tmp, bak, log, "Replace video with image"):
        return False, "Could not replace media file.", None

    return True, "", final


def try_strip_cover_to_tmp(
    media_path: Path, ext: str, strip_tmp: Path, is_video_ext: bool, log: list[str]
) -> bool:
    def attempt(cmd: str) -> bool:
        _safe_delete(strip_tmp)
        return _run_cmd(cmd, log) == 0 and strip_tmp.is_file()

    cmd_video = (
        f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map_chapters 0 "
        f"-map 0:v:0 -map 0:a? -map 0:s? -map 0:d? -map 0:t? -c copy {_quote(strip_tmp)}"
    )
    cmd_mkv_audio = (
        f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map_chapters 0 "
        f"-map 0:a? -map 0:s? -c copy {_quote(strip_tmp)}"
    )
    cmd_m4a = (
        f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map_chapters 0 "
        f"-map 0:a? -c copy {_quote(strip_tmp)}"
    )
    cmd_mp3 = f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map 0:a? -c copy {_quote(strip_tmp)}"
    cmd_audio = f"ffmpeg.exe -y -i {_quote(media_path)} -map_metadata 0 -map 0:a? -c copy {_quote(strip_tmp)}"

    if ext in (".mkv", ".mka"):
        return attempt(cmd_video) or attempt(cmd_mkv_audio)
    if is_video_ext:
        return attempt(cmd_video)
    if ext in (".m4a", ".m4b", ".m4p", ".aac", ".adts"):
        return attempt(cmd_m4a)
    if ext == ".mp3":
        return attempt(cmd_mp3)
    return attempt(cmd_audio)


def split_cover_from_media(media_path: Path, log: list[str]) -> str:
    name = media_path.name
    folder = media_path.parent
    stem = media_path.stem
    ext = file_ext_lower(name)
    is_video_ext = is_thumb_video(name)
    out_path = unique_jpg_path_next_to_media(folder, stem)
    strip_tmp = folder / f"{stem}.__opus_strip_tmp{ext}"
    strip_bak = folder / f"{stem}.__opus_strip_orig{ext}"

    if not try_extract_cover_to_path(media_path, name, folder, stem, ext, out_path, log):
        log.append(f"Split cover: no embedded cover or extract failed: {name}")
        return "fail"

    _safe_delete(strip_tmp)
    _safe_delete(strip_bak)
    if not try_strip_cover_to_tmp(media_path, ext, strip_tmp, is_video_ext, log):
        log.append(f"Split cover: cover saved but strip failed: {name} -> {out_path}")
        return "partial"

    if not _replace_in_place(media_path, strip_tmp, strip_bak, log, "Split cover"):
        return "partial"

    log.append(f"Split cover OK: {media_path} + {out_path}")
    return "ok"


def mono_audio_encode_args(ext: str) -> str:
    if ext == ".webm":
        return "libopus -ac 1 -b:a 128k"
    if ext == ".avi":
        return "libmp3lame -ac 1 -b:a 192k"
    if ext == ".wmv":
        return "wmav2 -ac 1 -b:a 128k"
    if ext in (".ogv", ".ogm"):
        return "libvorbis -ac 1 -b:a 192k"
    if ext in (".flv", ".f4v"):
        return "aac -ac 1 -b:a 192k"
    if ext in (".mpg", ".mpeg", ".mpe", ".m1v", ".vob"):
        return "mp2 -ac 1 -b:a 192k"
    return "aac -ac 1 -b:a 192k"


def video_encode_for_transform(ext: str) -> str:
    if ext == ".webm":
        return "libvpx-vp9 -crf 30 -b:v 0"
    return "libx264 -crf 18 -preset fast -pix_fmt yuv420p"


def merge_output_encode_args(ext: str, crf_str: str) -> str:
    q = (crf_str or "").strip() or "23"
    if ext == ".webm":
        return f"-c:v libvpx-vp9 -crf {q} -b:v 0 -c:a libopus -b:a 128k"
    if ext in (".ogv", ".ogm"):
        return "-c:v libtheora -q:v 7 -c:a libvorbis -b:a 192k"
    if ext in (".wmv", ".asf"):
        return "-c:v wmv2 -b:v 2M -c:a wmav2 -b:a 128k"
    if ext == ".avi":
        return f"-c:v libx264 -crf {q} -preset fast -pix_fmt yuv420p -c:a libmp3lame -b:a 192k"
    mov = " -movflags +faststart" if ext in (".mp4", ".m4v") else ""
    return f"-c:v libx264 -crf {q} -preset fast -pix_fmt yuv420p -c:a aac -b:a 192k{mov}"


def build_merge_filter_complex(
    count: int,
    durations_sec: list[float],
    has_audio: list[bool],
    width: int,
    height: int,
) -> str:
    parts: list[str] = []
    concat_in = ""
    for i in range(count):
        parts.append(
            f"[{i}:v:0]scale={width}:{height}:force_original_aspect_ratio=decrease,"
            f"pad={width}:{height}:(ow-iw)/2:(oh-ih)/2,setsar=1,fps=30,format=yuv420p,"
            f"setpts=PTS-STARTPTS[v{i}]"
        )
        if has_audio[i]:
            parts.append(
                f"[{i}:a:0]aformat=sample_rates=48000:channel_layouts=stereo,"
                f"asetpts=PTS-STARTPTS[a{i}]"
            )
        else:
            parts.append(
                f"aevalsrc=0:channel_layout=stereo:sample_rate=48000:"
                f"duration={durations_sec[i]}[a{i}]"
            )
        concat_in += f"[v{i}][a{i}]"
    parts.append(f"{concat_in}concat=n={count}:v=1:a=1[outv][outa]")
    return ";".join(parts)


def escape_ffmeta_value(s: str) -> str:
    return (
        s.replace("\\", "\\\\")
        .replace("=", "\\=")
        .replace(";", "\\;")
        .replace("#", "\\#")
        .replace("\n", "\\\n")
        .replace("\r", "")
    )


def write_merge_chapter_metadata(meta_path: Path, chapters: list[dict]) -> None:
    lines = [";FFMETADATA1"]
    for ch in chapters:
        lines.extend([
            "[CHAPTER]",
            "TIMEBASE=1/1000",
            f"START={round(ch['start_ms'])}",
            f"END={round(ch['end_ms'])}",
            f"title={escape_ffmeta_value(ch['title'])}",
        ])
    meta_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def format_youtube_timestamp(start_ms: int) -> str:
    total_sec = max(0, round(start_ms / 1000))
    h = total_sec // 3600
    m = (total_sec % 3600) // 60
    s = total_sec % 60
    if h > 0:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m}:{s:02d}"


def format_youtube_chapters(chapters: list[dict]) -> str:
    lines: list[str] = []
    for ch in chapters:
        title = str(ch.get("title", "")).strip() or "Chapter"
        lines.append(f"{format_youtube_timestamp(ch['start_ms'])} {title}")
    return "\n".join(lines)


def copy_youtube_chapters(chapters: list[dict]) -> tuple[str, bool]:
    text = format_youtube_chapters(chapters)
    return text, copy_text_to_clipboard(text)


def probe_embedded_chapters(media_path: Path) -> list[dict]:
    cmd = [
        "ffprobe.exe", "-v", "quiet", "-print_format", "json",
        "-show_chapters", os.fspath(media_path),
    ]
    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=False,
            creationflags=_win_subprocess_flags(),
        )
        if r.returncode != 0:
            return []
        data = json.loads(r.stdout)
    except (OSError, json.JSONDecodeError):
        return []
    chapters: list[dict] = []
    for ch in data.get("chapters") or []:
        try:
            start_ms = round(float(ch.get("start_time", 0)) * 1000)
        except (TypeError, ValueError):
            start_ms = 0
        tags = ch.get("tags") or {}
        title = tags.get("title") or tags.get("TITLE") or ""
        chapters.append({"title": str(title), "start_ms": start_ms})
    return chapters


def build_chapters_from_video_list(paths: list[Path]) -> tuple[list[dict], Optional[str]]:
    paths = sorted(paths, key=lambda p: p.name.lower())
    for p in paths:
        if not is_thumb_video(p.name):
            return [], f"Not a video: {p.name}"
    chapters: list[dict] = []
    start_ms = 0
    for p in paths:
        dur = probe_media_duration_sec(p)
        if dur <= 0:
            return [], f"Could not read duration: {p.name}"
        chapters.append({"title": p.stem, "start_ms": start_ms})
        start_ms += round(dur * 1000)
    return chapters, None


def chapters_to_youtube_text(paths: list[Path]) -> tuple[Optional[str], Optional[str]]:
    if not paths:
        return None, "No files selected."
    videos = [p for p in paths if is_thumb_video(p.name)]
    if not videos:
        return None, "No video files in selection."

    if len(videos) == 1:
        chapters = probe_embedded_chapters(videos[0])
        if not chapters:
            return None, f"No embedded chapters in {videos[0].name}."
    else:
        chapters, err = build_chapters_from_video_list(videos)
        if err:
            return None, err

    if chapters:
        chapters[0]["start_ms"] = 0
    text = format_youtube_chapters(chapters)
    return text, None


def merge_success_summary(
    path_count: int, merge_mode: str, out_path: Path, chapters: list[dict]
) -> str:
    youtube_text, copied = copy_youtube_chapters(chapters)
    summary = (
        f"Merged {path_count} video(s) ({merge_mode}):\n{out_path}\n\n"
        f"Chapters: {', '.join(ch['title'] for ch in chapters)}"
    )
    if copied:
        summary += f"\n\nYouTube chapters copied to clipboard:\n{youtube_text}"
    else:
        summary += "\n\nCould not copy YouTube chapters to clipboard."
    return summary


def copy_text_to_clipboard(text: str) -> bool:
    if sys.platform != "win32" or not text:
        return False
    try:
        import ctypes

        CF_UNICODETEXT = 13
        GMEM_MOVEABLE = 0x0002
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p

        data = text.replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-16-le") + b"\x00\x00"
        h_global = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not h_global:
            return False
        p_global = kernel32.GlobalLock(h_global)
        if not p_global:
            kernel32.GlobalFree(h_global)
            return False
        ctypes.memmove(p_global, data, len(data))
        kernel32.GlobalUnlock(h_global)

        for _ in range(10):
            if user32.OpenClipboard(None):
                try:
                    if not user32.EmptyClipboard():
                        kernel32.GlobalFree(h_global)
                        return False
                    if user32.SetClipboardData(CF_UNICODETEXT, h_global):
                        return True
                    kernel32.GlobalFree(h_global)
                    return False
                finally:
                    user32.CloseClipboard()
            time.sleep(0.05)

        kernel32.GlobalFree(h_global)
        return False
    except Exception:
        return False


def parse_trim_frame_count(s: str) -> int:
    s = (s or "").strip()
    if not s or not s.isdigit():
        return -1
    n = int(s)
    return n if n >= 1 else -1


def ffmpeg_trim_leading_frames_exec(
    vid_path: Path, tmp_path: Path, frame_count: int, fps: float, ext: str
) -> str:
    start_sec = frame_count / fps
    ss = str(round(start_sec * 1_000_000) / 1_000_000)
    v_enc = video_encode_for_transform(ext)
    return (
        f"ffmpeg.exe -y -i {_quote(vid_path)} -ss {ss} -map_metadata 0 -map_chapters 0 "
        f'-map 0:v:0 -map "0:a?" -c:v {v_enc} -c:a copy {_quote(tmp_path)}'
    )


def format_preset_by_name(formats: tuple[FormatPreset, ...], name: str) -> int:
    for i, f in enumerate(formats):
        if f.name == name:
            return i
    return 0


def quality_applicable(is_video: bool, fmt: FormatPreset) -> bool:
    return is_video and fmt.crf


# --- cover pairing ---


def cover_stems_fuzzy_match(media_stem: str, image_stem: str) -> bool:
    """True when one stem is fully contained in the other (e.g. song + song_cover)."""
    a, b = media_stem.casefold(), image_stem.casefold()
    if a == b:
        return True
    return a in b or b in a


def match_cover_image_media_pairs(
    imgs: list[Path], media: list[Path]
) -> tuple[list[tuple[Path, Path]], Optional[str]]:
    """Pair each media file with one image by fuzzy stem match (full stem substring)."""
    problems: list[str] = []

    img_stems: dict[str, list[Path]] = {}
    for img in imgs:
        img_stems.setdefault(img.stem, []).append(img)
    med_stems: dict[str, list[Path]] = {}
    for med in media:
        med_stems.setdefault(med.stem, []).append(med)

    for stem, lst in sorted(img_stems.items()):
        if len(lst) > 1:
            names = ", ".join(p.name for p in lst)
            problems.append(f"Multiple images named \"{stem}\": {names}")
    for stem, lst in sorted(med_stems.items()):
        if len(lst) > 1:
            names = ", ".join(p.name for p in lst)
            problems.append(f"Multiple media files named \"{stem}\": {names}")

    media_matches: dict[Path, list[Path]] = {med: [] for med in media}
    image_matches: dict[Path, list[Path]] = {img: [] for img in imgs}
    for med in media:
        for img in imgs:
            if cover_stems_fuzzy_match(med.stem, img.stem):
                media_matches[med].append(img)
                image_matches[img].append(med)

    for med in sorted(media, key=lambda p: p.name.lower()):
        matches = media_matches[med]
        if not matches:
            problems.append(f"No matching image for: {med.name}")
        elif len(matches) > 1:
            names = ", ".join(p.name for p in matches)
            problems.append(f"Multiple images match {med.name}: {names}")

    for img in sorted(imgs, key=lambda p: p.name.lower()):
        matches = image_matches[img]
        if not matches:
            problems.append(f"No matching media for: {img.name}")
        elif len(matches) > 1:
            names = ", ".join(p.name for p in matches)
            problems.append(f"Multiple media match {img.name}: {names}")

    if problems:
        return [], "Could not match cover pairs:\n\n" + "\n".join(problems)

    pairs = [
        (med, media_matches[med][0])
        for med in sorted(media, key=lambda p: p.name.lower())
    ]
    return pairs, None


def cover_combine_preflight(paths: list[Path]) -> Optional[str]:
    """Return an error message when image+media selection cannot be paired, else None."""
    imgs = [p for p in paths if is_thumb_image(p.name)]
    if not imgs:
        return None
    media = [p for p in paths if is_thumb_media(p.name)]
    bad = [p.name for p in paths if p not in imgs and p not in media]
    if bad:
        return f"Unsupported file type(s): {', '.join(bad)}"
    _, err = match_cover_image_media_pairs(imgs, media)
    return err


# --- action runners ---


def run_convert(
    paths: list[Path], mode: int, format_name: str, quality: str
) -> ActionResult:
    log = _job_log()
    if not paths:
        return ActionResult(False, "No files selected.", log)

    is_video = mode == 0
    formats = VIDEO_FORMATS if is_video else AUDIO_FORMATS
    fmt = formats[format_preset_by_name(formats, format_name)]
    q = (quality or "").strip() or "23"

    processed = failed = 0
    for item in paths:
        if _cancelled():
            break
        out = item.parent / f"{item.stem}{fmt.ext}"
        counter = 1
        while out.exists():
            out = item.parent / f"{item.stem}_{counter}{fmt.ext}"
            counter += 1

        if is_video:
            vcodec = fmt.codec
            if fmt.crf:
                vcodec = re.sub(r"-crf\s+\d+", f"-crf {q}", vcodec)
            cmd = (
                f"ffmpeg.exe -i {_quote(item)} -map_metadata 0 -map_chapters 0 "
                f"-c:v {vcodec} -y {_quote(out)}"
            )
        else:
            cmd = (
                f"ffmpeg.exe -i {_quote(item)} -map_metadata 0 -map_chapters 0 -vn "
                f"-c:a {fmt.codec} -y {_quote(out)}"
            )

        log.append(f"Converting: {item.name} -> {fmt.name}")
        if _run_cmd(cmd, log, job_output=out) == 0:
            processed += 1
            log.append(f"Success: {out}")
        else:
            failed += 1
            log.append(f"Failed: {item.name}")

    summary = f"Conversion finished. Successful: {processed}"
    if failed:
        summary += f", Failed: {failed}"
    return ActionResult(failed == 0 or processed > 0, summary, log)


def run_split_or_combine_cover(
    paths: list[Path], replace_with_image: bool
) -> ActionResult:
    log = _job_log()
    title = "Replace video with image" if replace_with_image else "Split/combine cover"

    imgs = [p for p in paths if is_thumb_image(p.name)]
    media = [p for p in paths if is_thumb_media(p.name)]
    bad = [p.name for p in paths if p not in imgs and p not in media]

    if bad:
        return ActionResult(False, f"Unsupported file type(s): {', '.join(bad)}", log)

    if imgs:
        pairs, match_err = match_cover_image_media_pairs(imgs, media)
        if match_err:
            return ActionResult(False, match_err, log)

        ok_n = fail = 0
        for med, img in pairs:
            if _cancelled():
                break
            if replace_with_image:
                ok, err, out = thumb_replace_video_with_image(med, img, log)
            else:
                ok, err, out = thumb_embed_cover(med, img, file_ext_lower(img.name), log)
            if not ok:
                fail += 1
                log.append(f"Failed {med.name}: {err}")
                continue
            _safe_delete(img)
            ok_n += 1
            log.append(f"OK: {med.name} + {img.name} -> {out}")

        if ok_n == 0:
            return ActionResult(False, f"{title} failed for all {len(pairs)} pair(s).", log)
        summary = f"{title} finished. OK: {ok_n}"
        if fail:
            summary += f", Failed: {fail}"
        return ActionResult(True, summary, log)

    if not media:
        return ActionResult(False, "To split cover: select video or audio files (no image).", log)

    ok_n = partial = fail = 0
    for m in media:
        res = split_cover_from_media(m, log)
        if res == "ok":
            ok_n += 1
        elif res == "partial":
            partial += 1
        else:
            fail += 1

    summary = f"Split cover finished. Full: {ok_n}"
    if partial:
        summary += f", Cover saved but strip failed: {partial}"
    if fail:
        summary += f", Failed: {fail}"
    return ActionResult(fail == 0 or ok_n > 0 or partial > 0, summary, log)


def run_audio_to_mono(paths: list[Path]) -> ActionResult:
    log = _job_log()
    videos = [p for p in paths if is_thumb_video(p.name)]
    if not videos:
        return ActionResult(False, "Select one or more video files.", log)
    if any(not is_thumb_video(p.name) for p in paths):
        bad = [p.name for p in paths if not is_thumb_video(p.name)]
        return ActionResult(False, f"Not supported video file(s): {', '.join(bad)}", log)

    ok = fail = 0
    for vid in videos:
        ext = file_ext_lower(vid.name)
        tmp = vid.parent / f"{vid.stem}.__opus_mono_tmp{ext}"
        bak = vid.parent / f"{vid.stem}.__opus_mono_orig{ext}"
        a_enc = mono_audio_encode_args(ext)
        _safe_delete(tmp)
        _safe_delete(bak)
        cmd = (
            f"ffmpeg.exe -y -i {_quote(vid)} -map_metadata 0 -map_chapters 0 -map 0 "
            f"-c copy -c:a {a_enc} {_quote(tmp)}"
        )
        if _run_cmd(cmd, log) != 0 or not tmp.is_file():
            fail += 1
            continue
        if _replace_in_place(vid, tmp, bak, log, "Audio to mono"):
            ok += 1
        else:
            fail += 1

    summary = f"Audio to mono finished. OK: {ok}"
    if fail:
        summary += f", Failed: {fail}"
    return ActionResult(ok > 0, summary, log)


def run_extract_audio_channels(paths: list[Path]) -> ActionResult:
    log = _job_log()
    title = "Extract audio channels"
    items = [p for p in paths if is_thumb_video(p.name) or is_thumb_audio(p.name)]
    if not items:
        return ActionResult(False, "Select video or audio files.", log)
    if len(items) != len(paths):
        bad = [p.name for p in paths if p not in items]
        return ActionResult(False, f"Unsupported: {', '.join(bad)}", log)

    ok = fail = 0
    for i, item in enumerate(items):
        ch = probe_audio_channel_count(item)
        if ch < 1:
            log.append(f"{title}: no audio stream: {item.name}")
            fail += 1
            continue

        filt_path = Path(os.environ.get("TEMP", ".")) / f"DOpus_ffmpeg_chfilt_{i}_{time.time_ns()}.txt"
        register_temp_path(filt_path)
        parts = [f"[0:a:0]pan=mono|c0=c{c}[ch{c}]" for c in range(ch)]
        try:
            filt_path.write_text(";".join(parts) + "\n", encoding="utf-8")
        except OSError as ex:
            log.append(f"{title}: could not write filter file: {ex}")
            fail += 1
            continue

        cmd = f'ffmpeg.exe -y -i {_quote(item)} -filter_complex_script {_quote(filt_path)} -c:a pcm_s16le'
        wav_outs: list[Path] = []
        for c in range(ch):
            idx = c + 1
            suffix = f"0{idx}" if idx < 10 else str(idx)
            out = item.parent / f"{item.stem}.ch{suffix}.wav"
            counter = 1
            while out.exists():
                out = item.parent / f"{item.stem}.ch{suffix}_{counter}.wav"
                counter += 1
            wav_outs.append(out)
            cmd += f' -map "[ch{c}]" {_quote(out)}'

        log.append(f"{title} ({ch} ch): {cmd}")
        if _run_cmd(cmd, log, job_outputs=wav_outs) == 0:
            ok += 1
        else:
            fail += 1
        _safe_delete(filt_path)

    summary = f"{title} finished. OK: {ok}"
    if fail:
        summary += f", Failed: {fail}"
    return ActionResult(ok > 0, summary, log)


def run_split_av_copy(paths: list[Path]) -> ActionResult:
    log = _job_log()
    title = "Split/combine AV"
    videos = [p for p in paths if is_thumb_video(p.name)]
    audios = [p for p in paths if is_thumb_audio(p.name)]
    bad = [p.name for p in paths if p not in videos and p not in audios]

    if bad:
        return ActionResult(False, f"Unsupported: {', '.join(bad)}", log)

    if audios:
        if len(paths) != 2 or len(videos) != 1 or len(audios) != 1:
            return ActionResult(
                False,
                "To combine: one video + one audio. To split: video file(s) only.",
                log,
            )
        vid, aud = videos[0], audios[0]
        ext = file_ext_lower(vid.name)
        mux_tmp = vid.parent / f"{vid.stem}.__opus_mux_tmp{ext}"
        bak = vid.parent / f"{vid.stem}.__opus_mux_orig{ext}"
        _safe_delete(mux_tmp)
        _safe_delete(bak)
        cmd = (
            f"ffmpeg.exe -y -i {_quote(vid)} -i {_quote(aud)} -map_metadata 0 -map_chapters 0 "
            f"-map 0:v:0 -map 1:a:0 -c copy -shortest {_quote(mux_tmp)}"
        )
        log.append(f"{title} (combine): {cmd}")
        if _run_cmd(cmd, log) != 0 or not mux_tmp.is_file():
            return ActionResult(False, "Combine failed (ffmpeg).", log)
        _safe_delete(aud)
        if _replace_in_place(vid, mux_tmp, bak, log, title):
            return ActionResult(True, f"Remuxed to: {vid}", log)
        return ActionResult(False, "Combine: could not replace video file.", log)

    if not videos:
        return ActionResult(False, "No video file in selection.", log)

    ok = partial = fail = 0
    for vid in videos:
        ext = file_ext_lower(vid.name)
        vid_tmp = vid.parent / f"{vid.stem}.__opus_split_v_tmp{ext}"
        bak = vid.parent / f"{vid.stem}.__opus_split_orig{ext}"
        aud_out = vid.parent / f"{vid.stem}.audio.mka"
        ac = 1
        while aud_out.exists():
            aud_out = vid.parent / f"{vid.stem}.audio_{ac}.mka"
            ac += 1

        _safe_delete(vid_tmp)
        _safe_delete(bak)
        cmd_v = (
            f"ffmpeg.exe -y -i {_quote(vid)} -map_metadata 0 -map_chapters 0 "
            f"-map 0:v:0 -c copy -an {_quote(vid_tmp)}"
        )
        log.append(f"{title} (split video): {cmd_v}")
        if _run_cmd(cmd_v, log) != 0 or not vid_tmp.is_file():
            fail += 1
            continue

        cmd_a = (
            f"ffmpeg.exe -y -i {_quote(vid)} -map_metadata 0 -map_chapters 0 "
            f"-map 0:a:0 -c copy -vn {_quote(aud_out)}"
        )
        log.append(f"{title} (split audio): {cmd_a}")
        audio_ok = _run_cmd(cmd_a, log, job_output=aud_out) == 0 and aud_out.is_file()

        if _replace_in_place(vid, vid_tmp, bak, log, title):
            ok += 1 if audio_ok else 0
            partial += 0 if audio_ok else 1
        else:
            fail += 1

    summary = f"Split finished. Full: {ok}, Video-only (no separate audio): {partial}"
    if fail:
        summary += f", Failed: {fail}"
    return ActionResult(fail == 0 or ok > 0 or partial > 0, summary, log)


def run_merge_videos(paths: list[Path], crf_str: str, *, fix_outliers: bool = True) -> ActionResult:
    log = _job_log()
    title = "Merge videos"
    if len(paths) < 2:
        return ActionResult(False, "Select two or more video files.", log)
    paths = sorted(paths, key=lambda p: p.name.lower())
    if any(not is_thumb_video(p.name) for p in paths):
        bad = [p.name for p in paths if not is_thumb_video(p.name)]
        return ActionResult(False, f"Not video: {', '.join(bad)}", log)

    infos: list[MergeStreamInfo] = []
    for p in paths:
        info = probe_merge_stream_info(p)
        if not info:
            return ActionResult(False, f"Could not read media info: {p.name}", log)
        infos.append(info)

    titles = [p.stem for p in paths]
    durations = [info.duration for info in infos]
    has_audio = [bool(info.audio_codec) for info in infos]

    ext = file_ext_lower(paths[0].name)
    out_path = paths[0].parent / f"output{ext}"
    meta_path = Path(os.environ.get("TEMP", ".")) / f"DOpus_ffmpeg_merge_{time.time_ns()}.ffmeta"
    register_temp_path(meta_path)

    start_ms = 0
    chapters: list[dict] = []
    for i, dur in enumerate(durations):
        end_ms = start_ms + round(dur * 1000)
        chapters.append({"title": titles[i], "start_ms": start_ms, "end_ms": end_ms})
        start_ms = end_ms

    try:
        write_merge_chapter_metadata(meta_path, chapters)
    except OSError as ex:
        return ActionResult(False, f"Could not write chapter metadata: {ex}", log)

    lossless = merge_lossless_compatible(infos, ext, paths)
    list_path: Optional[Path] = None
    fixed_temps: list[Path] = []
    merge_mode = "lossless copy"
    merge_has_audio = has_audio[0]

    try:
        if lossless:
            merge_paths = paths
        else:
            warnings, target, outlier_idxs = merge_stream_mismatch_warnings(paths, infos)
            for line in warnings:
                log.append(line)

            if fix_outliers and outlier_idxs and merge_outlier_fix_eligible(infos, target):
                merge_paths = list(paths)
                merge_mode = "fix outliers + lossless copy"
                merge_has_audio = bool(target.audio_codec)
                log.append(
                    f"Re-encoding {len(outlier_idxs)} outlier(s) to match, then lossless merge."
                )
                for i in outlier_idxs:
                    tmp = (
                        Path(os.environ.get("TEMP", "."))
                        / f"DOpus_ffmpeg_merge_{time.time_ns()}_{i}_fixed{ext}"
                    )
                    register_temp_path(tmp)
                    fixed_temps.append(tmp)
                    cmd_fix = build_merge_outlier_fix_cmd(
                        paths[i], tmp, infos[i], target, ext, crf_str
                    )
                    log.append(f"{title} (fix outlier {paths[i].name}): {cmd_fix}")
                    if _run_cmd(cmd_fix, log, job_output=tmp) != 0 or not tmp.is_file():
                        return ActionResult(
                            False, f"Failed to fix outlier: {paths[i].name}", log
                        )
                    fixed_info = probe_merge_stream_info(tmp)
                    if not fixed_info or merge_stream_signature(fixed_info) != merge_stream_signature(target):
                        return ActionResult(
                            False,
                            f"Fixed outlier still does not match: {paths[i].name}",
                            log,
                        )
                    merge_paths[i] = tmp
            else:
                if fix_outliers and outlier_idxs and not merge_outlier_fix_eligible(infos, target):
                    log.append(
                        "Outlier fix not possible (e.g. missing/extra audio). Re-encoding all files."
                    )
                elif not fix_outliers:
                    log.append("Re-encoding all files (fix outliers option is off).")
                merge_mode = "re-encoded"
                width = max(info.width for info in infos)
                height = max(info.height for info in infos)
                if width % 2:
                    width += 1
                if height % 2:
                    height += 1

                meta_idx = len(paths)
                cmd = "ffmpeg.exe -y"
                for p in paths:
                    cmd += f" -i {_quote(p)}"
                cmd += f" -i {_quote(meta_path)}"
                cmd += (
                    f' -filter_complex "{build_merge_filter_complex(len(paths), durations, has_audio, width, height)}"'
                )
                cmd += f" -map [outv] -map [outa] -map_metadata {meta_idx} -map_chapters {meta_idx}"
                cmd += f" {merge_output_encode_args(ext, crf_str)} {_quote(out_path)}"
                log.append(f"{title} (re-encode): {cmd}")
                exit_code = _run_cmd(cmd, log, job_output=out_path)
                if exit_code != 0 or not out_path.is_file():
                    return ActionResult(False, f"Merge failed (exit {exit_code}).", log)
                return ActionResult(
                    True,
                    merge_success_summary(len(paths), merge_mode, out_path, chapters),
                    log,
                )

        list_path = Path(os.environ.get("TEMP", ".")) / f"DOpus_ffmpeg_merge_{time.time_ns()}.concat.txt"
        register_temp_path(list_path)
        try:
            write_concat_list_file(list_path, merge_paths)
        except OSError as ex:
            return ActionResult(False, f"Could not write concat list: {ex}", log)

        cmd = build_lossless_merge_cmd(
            list_path, meta_path, out_path, ext, merge_has_audio
        )
        log.append(f"{title} ({merge_mode}): {cmd}")
        exit_code = _run_cmd(cmd, log, job_output=out_path)
        if exit_code != 0 or not out_path.is_file():
            return ActionResult(False, f"Merge failed (exit {exit_code}).", log)

        return ActionResult(
            True,
            merge_success_summary(len(paths), merge_mode, out_path, chapters),
            log,
        )
    finally:
        _safe_delete(meta_path)
        if list_path:
            _safe_delete(list_path)
        for tmp in fixed_temps:
            _safe_delete(tmp)


def run_video_transform(paths: list[Path], vf_filter: str, log_title: str) -> ActionResult:
    log = _job_log()
    videos = [p for p in paths if is_thumb_video(p.name)]
    if not videos:
        return ActionResult(False, "Select one or more video files.", log)
    if len(videos) != len(paths):
        bad = [p.name for p in paths if not is_thumb_video(p.name)]
        return ActionResult(False, f"Not video: {', '.join(bad)}", log)

    ok = fail = 0
    for vid in videos:
        ext = file_ext_lower(vid.name)
        tmp = vid.parent / f"{vid.stem}.__opus_xform_tmp{ext}"
        bak = vid.parent / f"{vid.stem}.__opus_xform_orig{ext}"
        v_enc = video_encode_for_transform(ext)
        _safe_delete(tmp)
        _safe_delete(bak)
        cmd = (
            f"ffmpeg.exe -y -i {_quote(vid)} -vf \"{vf_filter}\" -map_metadata 0 -map_chapters 0 "
            f'-map 0:v:0 -map "0:a?" -c:v {v_enc} -c:a copy {_quote(tmp)}'
        )
        log.append(f"{log_title}: {cmd}")
        if _run_cmd(cmd, log) != 0 or not tmp.is_file():
            fail += 1
            continue
        if _replace_in_place(vid, tmp, bak, log, log_title):
            ok += 1
        else:
            fail += 1

    summary = f"{log_title} finished. OK: {ok}"
    if fail:
        summary += f", Failed: {fail}"
    return ActionResult(ok > 0, summary, log)


def run_trim_leading_frames(paths: list[Path], frame_count_str: str) -> ActionResult:
    log = _job_log()
    title = "Trim leading frames"
    frame_count = parse_trim_frame_count(frame_count_str)
    if frame_count < 1:
        return ActionResult(False, "Enter a whole number of frames to skip (1 or more).", log)

    videos = [p for p in paths if is_thumb_video(p.name)]
    if not videos:
        return ActionResult(False, "Select one or more video files.", log)
    if len(videos) != len(paths):
        bad = [p.name for p in paths if not is_thumb_video(p.name)]
        return ActionResult(False, f"Not video: {', '.join(bad)}", log)

    ok = fail = 0
    for vid in videos:
        fps = probe_video_avg_frame_rate(vid)
        if fps <= 0:
            log.append(f"{title}: could not read frame rate: {vid.name}")
            fail += 1
            continue
        ext = file_ext_lower(vid.name)
        tmp = vid.parent / f"{vid.stem}.__opus_trim_tmp{ext}"
        bak = vid.parent / f"{vid.stem}.__opus_trim_orig{ext}"
        _safe_delete(tmp)
        _safe_delete(bak)
        cmd = ffmpeg_trim_leading_frames_exec(vid, tmp, frame_count, fps, ext)
        log.append(f"{title} (skip {frame_count} @ {fps} fps): {cmd}")
        if _run_cmd(cmd, log) != 0 or not tmp.is_file():
            fail += 1
            continue
        if _replace_in_place(vid, tmp, bak, log, title):
            ok += 1
        else:
            fail += 1

    summary = f"{title} finished (skipped first {frame_count} frame(s)). OK: {ok}"
    if fail:
        summary += f", Failed: {fail}"
    return ActionResult(ok > 0, summary, log)


def run_action(
    action: str,
    paths: list[Path],
    settings: Settings,
    *,
    on_output: OutputSink | None = None,
) -> ActionResult:
    begin_job(on_output=on_output)
    try:
        return _run_action_impl(action, paths, settings)
    finally:
        global _output_sink
        _output_sink = None


def _run_action_impl(
    action: str,
    paths: list[Path],
    settings: Settings,
) -> ActionResult:
    if action == "convert":
        return run_convert(paths, settings.mode, settings.format_name, settings.quality)
    if action == "cover":
        return run_split_or_combine_cover(paths, settings.replace_video_with_image)
    if action == "mono":
        return run_audio_to_mono(paths)
    if action == "splitav":
        return run_split_av_copy(paths)
    if action == "splitch":
        return run_extract_audio_channels(paths)
    if action == "mergevid":
        return run_merge_videos(paths, settings.quality, fix_outliers=settings.merge_fix_outliers)
    if action == "rotatecw":
        return run_video_transform(paths, "transpose=1", "Rotate 90° CW")
    if action == "rotateccw":
        return run_video_transform(paths, "transpose=2", "Rotate 90° CCW")
    if action == "fliph":
        return run_video_transform(paths, "hflip", "Flip horizontal")
    if action == "flipv":
        return run_video_transform(paths, "vflip", "Flip vertical")
    if action == "trimstart":
        return run_trim_leading_frames(paths, settings.trim_frames)
    return ActionResult(False, f"Unknown action: {action}", [])


def format_action_log(result: ActionResult) -> str:
    lines = [result.summary, ""]
    if result.log:
        lines.extend(result.log)
    return "\n".join(lines)


def run_cli(argv: list[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="FFmpeg tool (GUI and CLI).")
    parser.add_argument("--gui", action="store_true", help="Open Dear PyGui GUI.")
    parser.add_argument("--repeat", action="store_true", help="Run last saved action.")
    parser.add_argument("--action", choices=LAST_ACTIONS, help="Action to run.")
    parser.add_argument("--only-list", metavar="FILE", help="UTF-8 file, one path per line.")
    parser.add_argument("--only-file", action="append", default=[], metavar="PATH")
    args = parser.parse_args(argv)

    settings = config_load_settings()

    if args.repeat and not args.action:
        args.action = config_load_last_action()

    if args.action:
        paths = paths_from_only_list(args.only_list, args.only_file)
        if not paths:
            paths = parse_file_paths(settings.files_text)
        if not paths:
            print("No files to process.", file=sys.stderr)
            return 2
        result = run_action(args.action, paths, settings)
        settings.last_action = args.action
        config_save_settings(settings, last_action=args.action)
        print(format_action_log(result))
        return 0 if result.ok else 1

    from ffmpeg_gui import run_gui

    run_gui(
        initial_only_list=args.only_list,
        initial_only_files=args.only_file or None,
    )
    return 0
