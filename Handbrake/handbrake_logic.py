"""HandBrakeCLI encode logic (ported from DOpus_handbrake.js)."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))
from recycle_delete import safe_delete as _safe_delete

OutputSink = Callable[[str, bool], None]

PRESET_DIR = Path(__file__).resolve().parent
DEFAULT_MAX_PICTURE_SIDE = 1920
# 0xC000013A STATUS_CONTROL_C_EXIT -- user closed console / Ctrl+C / killed process
HANDBRAKE_EXIT_CONTROL_C = -1073741510
# Replace original only when output is smaller by more than this fraction (10%).
REPLACE_MIN_SIZE_REDUCTION = 0.1

VIDEO_EXT = {
    ".mp4", ".m4v", ".mov", ".qt", ".mkv", ".webm", ".avi", ".wmv", ".asf",
    ".mpg", ".mpeg", ".mpe", ".m1v", ".m2v", ".mpv", ".vob", ".ts", ".mts",
    ".m2t", ".m2ts", ".3gp", ".3g2", ".flv", ".f4v", ".ogv", ".ogm", ".dv", ".mxf",
}

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "HandbrakeTool"
CONFIG_PATH = CONFIG_DIR / "settings.json"
LEGACY_INI_PATH = Path(os.environ.get("APPDATA", "")) / "DOpus_handbrake_settings.ini"


class ValidationError(ValueError):
    """Raised for bad user input; message is shown to the user as-is."""


@dataclass
class Settings:
    preset_file: str = ""
    max_side: str = str(DEFAULT_MAX_PICTURE_SIDE)
    video_quality: str = ""
    video_framerate: str = ""
    small_file_cutoff_mb: str = ""
    small_file_quality: str = ""
    small_file_framerate: str = ""
    frame_range_start: str = ""
    frame_range_end: str = ""
    output_format: str = ""
    replace_original: bool = False
    files_text: str = ""


@dataclass
class ActionResult:
    ok: bool
    summary: str
    log: list[str] = field(default_factory=list)


# --- settings persistence ---


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
            "preset_file": legacy.get("presetFile", ""),
            "max_side": legacy.get("maxSide", str(DEFAULT_MAX_PICTURE_SIDE)),
            "video_quality": legacy.get("videoQuality", ""),
            "video_framerate": legacy.get("videoFramerate", ""),
            "small_file_cutoff_mb": legacy.get("smallFileCutoffMb", ""),
            "small_file_quality": legacy.get("smallFileQuality", ""),
            "small_file_framerate": legacy.get("smallFileFramerate", ""),
            "frame_range_start": legacy.get("frameRangeStart", ""),
            "frame_range_end": legacy.get("frameRangeEnd", ""),
            "replace_original": legacy.get("replaceOriginal") == "1",
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
    return Settings(
        preset_file=str(data.get("preset_file") or ""),
        max_side=str(data.get("max_side") or DEFAULT_MAX_PICTURE_SIDE),
        video_quality=str(data.get("video_quality") or ""),
        video_framerate=str(data.get("video_framerate") or ""),
        small_file_cutoff_mb=str(data.get("small_file_cutoff_mb") or ""),
        small_file_quality=str(data.get("small_file_quality") or ""),
        small_file_framerate=str(data.get("small_file_framerate") or ""),
        frame_range_start=str(data.get("frame_range_start") or ""),
        frame_range_end=str(data.get("frame_range_end") or ""),
        output_format=str(data.get("output_format") or ""),
        replace_original=bool(data.get("replace_original")),
        files_text=str(data.get("files_text") or ""),
    )


def config_save_settings(settings: Settings) -> None:
    data = {
        "preset_file": settings.preset_file,
        "max_side": settings.max_side,
        "video_quality": settings.video_quality,
        "video_framerate": settings.video_framerate,
        "small_file_cutoff_mb": settings.small_file_cutoff_mb,
        "small_file_quality": settings.small_file_quality,
        "small_file_framerate": settings.small_file_framerate,
        "frame_range_start": settings.frame_range_start,
        "frame_range_end": settings.frame_range_end,
        "output_format": settings.output_format,
        "replace_original": settings.replace_original,
        "files_text": settings.files_text,
    }
    try:
        config_write(data)
    except OSError:
        pass


# --- file path handling ---


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


def paths_from_text_lines(text: str) -> list[Path]:
    paths: list[Path] = []
    for ln in text.splitlines():
        s = ln.strip()
        if not s:
            continue
        p = Path(s)
        if p.is_file() or p.is_dir():
            paths.append(p.resolve())
    return dedupe_paths(paths)


def _files_in_folder_recursive(folder: Path) -> list[Path]:
    out: list[Path] = []
    try:
        for child in sorted(folder.iterdir(), key=lambda p: p.name.casefold()):
            if child.is_dir():
                out.extend(_files_in_folder_recursive(child))
            elif child.is_file() and child.suffix.lower() in VIDEO_EXT:
                out.append(child.resolve())
    except OSError:
        pass
    return out


def expand_paths(paths: list[Path]) -> list[Path]:
    out: list[Path] = []
    for p in paths:
        if p.is_file():
            out.append(p.resolve())
        elif p.is_dir():
            out.extend(_files_in_folder_recursive(p))
    return dedupe_paths(out)


def parse_input_paths(text: str) -> list[Path]:
    return expand_paths(paths_from_text_lines(text))


def paths_from_only_list(list_path: Optional[str], only_files: Optional[list[str]]) -> list[Path]:
    lines: list[str] = []
    if list_path:
        try:
            lines.extend(Path(list_path).read_text(encoding="utf-8-sig").splitlines())
        except OSError:
            pass
    if only_files:
        lines.extend(only_files)
    return paths_from_text_lines("\n".join(lines))


def build_initial_files_text(
    saved: str,
    only_list: Optional[str],
    only_files: Optional[list[str]],
) -> str:
    from_dopus = paths_from_only_list(only_list, only_files)
    if from_dopus:
        return "\n".join(os.fspath(p) for p in from_dopus)
    return saved.strip()


def _delete_only_list_file(only_list: Optional[str]) -> None:
    if not only_list:
        return
    p = Path(only_list)
    name = p.name
    if name.startswith("HandbrakeTool_only_") and name.endswith(".txt"):
        _safe_delete(p)


# --- HandBrake / ffprobe discovery ---


def _win_subprocess_flags() -> int:
    if sys.platform != "win32":
        return 0
    return subprocess.CREATE_NO_WINDOW


def _program_files_roots() -> list[Path]:
    roots = []
    for key in ("ProgramFiles", "ProgramFiles(x86)"):
        val = os.environ.get(key)
        if val:
            roots.append(Path(val))
    return roots


def resolve_handbrake_cli() -> Optional[Path]:
    for root in _program_files_roots():
        std = root / "HandBrake" / "HandBrakeCLI.exe"
        if std.is_file():
            return std
    for root in _program_files_roots():
        hb_dir = root / "HandBrake"
        if not hb_dir.is_dir():
            continue
        try:
            for sub in hb_dir.iterdir():
                if sub.is_dir():
                    candidate = sub / "HandBrakeCLI.exe"
                    if candidate.is_file():
                        return candidate
        except OSError:
            pass
    return None


def resolve_handbrake_gui() -> Optional[Path]:
    for root in _program_files_roots():
        std = root / "HandBrake" / "HandBrake.exe"
        if std.is_file():
            return std
    return None


def resolve_ffprobe() -> str:
    found = shutil.which("ffprobe") or shutil.which("ffprobe.exe")
    if found:
        return found
    for root in _program_files_roots():
        candidate = root / "ffmpeg" / "bin" / "ffprobe.exe"
        if candidate.is_file():
            return os.fspath(candidate)
    return "ffprobe.exe"


def _parse_frame_rate_to_float(s: str) -> Optional[float]:
    s = s.strip().replace("\r", "").replace("\n", "")
    if not s:
        return None
    if "/" in s:
        num_s, _, den_s = s.partition("/")
        try:
            num = float(num_s)
            den = float(den_s)
            if den:
                return num / den
        except ValueError:
            pass
    try:
        return float(s)
    except ValueError:
        return None


def probe_source_video_frame_rate(media_path: Path) -> Optional[float]:
    ffprobe = resolve_ffprobe()
    for field_name in ("avg_frame_rate", "r_frame_rate"):
        cmd = [
            ffprobe, "-v", "error", "-select_streams", "v:0",
            "-show_entries", f"stream={field_name}",
            "-of", "default=noprint_wrappers=1:nokey=1",
            os.fspath(media_path),
        ]
        try:
            r = subprocess.run(
                cmd, capture_output=True, text=True, check=False,
                creationflags=_win_subprocess_flags(),
            )
        except OSError:
            continue
        for line in r.stdout.splitlines():
            v = _parse_frame_rate_to_float(line)
            if v is not None and v > 0:
                return v
    return None


def framerate_for_encode(requested_fps: Optional[float], source_fps: Optional[float]) -> Optional[float]:
    if requested_fps is None:
        return None
    if not source_fps or source_fps <= 0:
        return requested_fps
    return min(requested_fps, source_fps)


# --- preset JSON ---


def list_preset_json_files() -> list[Path]:
    if not PRESET_DIR.is_dir():
        return []
    rows = [p for p in PRESET_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".json"]
    rows.sort(key=lambda p: p.name.casefold())
    return rows


def output_ext_from_handbrake_file_format(file_format: str) -> str:
    key = (file_format or "").lower()
    if key == "av_mkv" or "mkv" in key:
        return ".mkv"
    if key == "av_mp4" or "mp4" in key:
        return ".mp4"
    if key == "av_webm" or "webm" in key:
        return ".webm"
    return ".mkv"


OUTPUT_FORMAT_CHOICES = ("", "mp4", "mkv")


def parse_output_format(raw: str) -> str:
    s = raw.strip().lower()
    if s not in OUTPUT_FORMAT_CHOICES:
        raise ValidationError('Output format must be blank (preset default), "mp4", or "mkv".')
    return s


def handbrake_format_flag(output_format: str) -> Optional[str]:
    if output_format == "mp4":
        return "av_mp4"
    if output_format == "mkv":
        return "av_mkv"
    return None


def active_preset_from_handbrake_json(preset_path: Path) -> tuple[str, str]:
    text = preset_path.read_text(encoding="utf-8-sig")
    root = json.loads(text)
    preset_list = root.get("PresetList")
    if not preset_list:
        raise ValidationError("PresetList missing or empty.")
    chosen = None
    for p in preset_list:
        if isinstance(p, dict) and p.get("Default") is True:
            chosen = p
            break
    if chosen is None:
        chosen = preset_list[0]
    if not chosen:
        raise ValidationError("No preset object in PresetList.")
    preset_name = chosen.get("PresetName")
    if not preset_name:
        raise ValidationError("Active preset has no PresetName.")
    return str(preset_name), output_ext_from_handbrake_file_format(chosen.get("FileFormat", ""))


# --- option parsing (mirrors DOpus_handbrake.js validation) ---


def parse_max_picture_side(raw: str) -> int:
    s = raw.strip()
    try:
        n = int(s)
    except ValueError:
        n = -1
    if n < 1:
        raise ValidationError("Enter a positive number for max picture size (pixels).")
    return n


def parse_video_quality_override(raw: str) -> Optional[float]:
    s = raw.strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        raise ValidationError("Video quality must be a number, or leave blank to use the preset.")


def parse_video_framerate_override(raw: str) -> Optional[float]:
    s = raw.strip()
    if s == "":
        return None
    try:
        n = float(s)
    except ValueError:
        n = -1
    if n <= 0:
        raise ValidationError("Frame rate must be a positive number, or leave blank to use the preset.")
    return n


@dataclass
class SmallFileRule:
    cutoff_mb: float
    quality: float
    framerate: Optional[float]


def parse_small_file_rule(
    cutoff_raw: str, quality_raw: str, framerate_raw: str
) -> Optional[SmallFileRule]:
    cutoff_s, quality_s, framerate_s = cutoff_raw.strip(), quality_raw.strip(), framerate_raw.strip()
    if cutoff_s == "" and quality_s == "" and framerate_s == "":
        return None
    if cutoff_s == "" or quality_s == "":
        raise ValidationError(
            "For small-file overrides, enter both the MB cutoff and -q, "
            "or leave cutoff, -q, and -r all blank."
        )
    try:
        cutoff_mb = float(cutoff_s)
    except ValueError:
        cutoff_mb = -1
    if cutoff_mb <= 0:
        raise ValidationError("Small-file cutoff must be a positive number (MB).")
    try:
        quality = float(quality_s)
    except ValueError:
        raise ValidationError("Small-file quality must be a number.")
    framerate = None
    if framerate_s != "":
        try:
            framerate = float(framerate_s)
        except ValueError:
            framerate = -1
        if framerate is None or framerate <= 0:
            raise ValidationError("Small-file frame rate must be a positive number, or leave -r blank.")
    return SmallFileRule(cutoff_mb=cutoff_mb, quality=quality, framerate=framerate)


@dataclass
class FrameRange:
    start_frame: int
    use_start_at: bool
    stop_duration: Optional[int]


def parse_frame_range(start_raw: str, end_raw: str) -> Optional[FrameRange]:
    start_s, end_s = start_raw.strip(), end_raw.strip()
    if start_s == "" and end_s == "":
        return None
    has_start, has_end = start_s != "", end_s != ""
    start_frame = 0
    if has_start:
        try:
            start_frame = int(start_s)
        except ValueError:
            start_frame = -1
        if start_frame < 0:
            raise ValidationError("Frame range start must be 0 or a positive whole number.")
    end_frame = 0
    if has_end:
        try:
            end_frame = int(end_s)
        except ValueError:
            end_frame = -1
        if end_frame < 0:
            raise ValidationError("Frame range end must be 0 or a positive whole number.")
    if has_start and has_end and end_frame < start_frame:
        raise ValidationError("Frame range end must be >= start.")
    stop_duration = None
    if has_end:
        stop_duration = end_frame - (start_frame if has_start else 0) + 1
        if stop_duration < 1:
            raise ValidationError("Frame range must include at least one frame.")
    return FrameRange(start_frame=start_frame if has_start else 0, use_start_at=has_start, stop_duration=stop_duration)


def frame_range_cli_args(frame_range: Optional[FrameRange]) -> list[str]:
    if not frame_range:
        return []
    args: list[str] = []
    if frame_range.use_start_at:
        args += ["--start-at", f"frames:{frame_range.start_frame}"]
    if frame_range.stop_duration is not None:
        args += ["--stop-at", f"frames:{frame_range.stop_duration}"]
    return args


def overrides_for_input(
    input_path: Path,
    video_quality: Optional[float],
    video_framerate: Optional[float],
    small_file_rule: Optional[SmallFileRule],
) -> tuple[Optional[float], Optional[float]]:
    if not small_file_rule:
        return video_quality, video_framerate
    size_mb = input_path.stat().st_size / (1024 * 1024)
    if size_mb < small_file_rule.cutoff_mb:
        framerate = small_file_rule.framerate if small_file_rule.framerate is not None else video_framerate
        return small_file_rule.quality, framerate
    return video_quality, video_framerate


@dataclass
class EncodeOptions:
    preset_path: Path
    max_picture_side: int
    video_quality: Optional[float]
    video_framerate: Optional[float]
    small_file_rule: Optional[SmallFileRule]
    frame_range: Optional[FrameRange]
    output_format: str
    replace_original: bool


def build_encode_options(settings: Settings, preset_path: Path) -> EncodeOptions:
    return EncodeOptions(
        preset_path=preset_path,
        max_picture_side=parse_max_picture_side(settings.max_side),
        video_quality=parse_video_quality_override(settings.video_quality),
        video_framerate=parse_video_framerate_override(settings.video_framerate),
        small_file_rule=parse_small_file_rule(
            settings.small_file_cutoff_mb, settings.small_file_quality, settings.small_file_framerate
        ),
        frame_range=parse_frame_range(settings.frame_range_start, settings.frame_range_end),
        output_format=parse_output_format(settings.output_format),
        replace_original=settings.replace_original,
    )


# --- output path / replace-original helpers ---


def _paths_equal_ignore_case(a: Path, b: Path) -> bool:
    return os.fspath(a).casefold() == os.fspath(b).casefold()


def output_path_for_input(input_path: Path, out_ext: str) -> Path:
    candidate = input_path.with_suffix(out_ext)
    if _paths_equal_ignore_case(candidate, input_path):
        return input_path.with_name(input_path.stem + "_hb" + out_ext)
    return candidate


def _size_reduction_exceeds_replace_threshold(input_path: Path, output_path: Path) -> bool:
    input_size = input_path.stat().st_size
    output_size = output_path.stat().st_size
    if input_size <= 0:
        return False
    return (input_size - output_size) / input_size > REPLACE_MIN_SIZE_REDUCTION


def replace_original_with_output(input_path: Path, output_path: Path, log: list[str]) -> bool:
    if not output_path.is_file():
        log.append(f"HandBrake: replace skipped -- no output file: {output_path}")
        return False
    if _paths_equal_ignore_case(input_path, output_path):
        return True
    same_ext = input_path.suffix.lower() == output_path.suffix.lower()
    if input_path.is_file():
        if not _safe_delete(input_path):
            log.append(f"HandBrake: could not recycle original: {input_path}")
            return False
        log.append(f"HandBrake: sent original to Recycle Bin: {input_path}")
    if same_ext:
        try:
            shutil.move(os.fspath(output_path), os.fspath(input_path))
            log.append(f"HandBrake: renamed output to replace original: {input_path}")
        except OSError:
            log.append(f"HandBrake: could not rename output to original name: {output_path}")
            return False
    return True


def delete_incomplete_output(output_path: Path, log: list[str]) -> None:
    if not output_path.is_file():
        return
    if _safe_delete(output_path):
        log.append(f"HandBrake: sent incomplete output to Recycle Bin: {output_path}")
    else:
        log.append(f"HandBrake: could not remove incomplete output ({output_path})")


# --- process execution / cancellation ---


_shutting_down = False
_active_procs: list[subprocess.Popen] = []
_active_procs_lock = threading.Lock()


def _cancelled() -> bool:
    return _shutting_down


def cancel_running_jobs() -> None:
    global _shutting_down
    _shutting_down = True
    with _active_procs_lock:
        procs = list(_active_procs)
        _active_procs.clear()
    for proc in procs:
        if proc.poll() is None:
            if sys.platform == "win32":
                subprocess.run(
                    f"taskkill /F /T /PID {proc.pid}", shell=True,
                    capture_output=True, check=False, creationflags=_win_subprocess_flags(),
                )
            else:
                proc.terminate()
    for proc in procs:
        try:
            proc.wait(timeout=3)
        except subprocess.TimeoutExpired:
            pass


def _reset_cancel_flag() -> None:
    global _shutting_down
    _shutting_down = False


def _stream_process_output(proc: subprocess.Popen, on_output: Optional[OutputSink]) -> None:
    if proc.stdout is None or on_output is None:
        return
    carry = ""
    try:
        while True:
            chunk = proc.stdout.read(4096)
            if not chunk:
                break
            carry += chunk.decode("utf-8", errors="replace")
            while carry:
                r = carry.find("\r")
                n = carry.find("\n")
                if r == -1 and n == -1:
                    break
                if n != -1 and (r == -1 or n < r):
                    line, carry = carry[:n], carry[n + 1:]
                    line = line.strip("\r")
                    if line:
                        on_output(line, False)
                else:
                    line, carry = carry[:r], carry[r + 1:]
                    if line:
                        on_output(line, True)
        tail = carry.strip("\r\n")
        if tail:
            on_output(tail, False)
    except OSError:
        pass


def _run_handbrake_cmd(cmd: list[str], on_output: Optional[OutputSink]) -> int:
    if _cancelled():
        return HANDBRAKE_EXIT_CONTROL_C
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            creationflags=_win_subprocess_flags(),
        )
    except OSError as ex:
        if on_output:
            on_output(f"Error: {ex}", False)
        return -1
    with _active_procs_lock:
        _active_procs.append(proc)
    reader = threading.Thread(target=_stream_process_output, args=(proc, on_output), daemon=True)
    reader.start()
    try:
        rc = proc.wait()
    finally:
        with _active_procs_lock:
            if proc in _active_procs:
                _active_procs.remove(proc)
        if proc.stdout is not None:
            try:
                proc.stdout.close()
            except OSError:
                pass
        reader.join(timeout=2.0)
    return rc


def _build_encode_cmd(
    cli: Path, input_path: Path, output_path: Path, preset_name: str,
    options: EncodeOptions, quality: Optional[float], framerate: Optional[float],
) -> list[str]:
    cmd = [
        os.fspath(cli),
        "--preset-import-file", os.fspath(options.preset_path),
        "-Z", preset_name,
        "--maxWidth", str(options.max_picture_side),
        "--maxHeight", str(options.max_picture_side),
        "--loose-anamorphic",
    ]
    format_flag = handbrake_format_flag(options.output_format)
    if format_flag is not None:
        cmd += ["-f", format_flag]
    if quality is not None:
        cmd += ["-q", str(quality)]
    if framerate is not None:
        cmd += ["-r", str(framerate)]
    cmd += frame_range_cli_args(options.frame_range)
    cmd += ["-i", os.fspath(input_path), "-o", os.fspath(output_path)]
    return cmd


def run_encode(
    paths: list[Path],
    options: EncodeOptions,
    on_output: Optional[OutputSink] = None,
) -> ActionResult:
    _reset_cancel_flag()
    log: list[str] = []

    def emit(text: str) -> None:
        log.append(text)
        if on_output:
            on_output(text, False)

    try:
        preset_name, out_ext = active_preset_from_handbrake_json(options.preset_path)
    except (ValidationError, OSError, json.JSONDecodeError) as ex:
        return ActionResult(False, f"Could not read preset JSON: {ex}\n\n{options.preset_path}", log)
    if options.output_format:
        out_ext = "." + options.output_format

    cli = resolve_handbrake_cli()
    if not cli:
        return ActionResult(False, "HandBrakeCLI.exe not found under Program Files\\HandBrake.", log)

    emit(
        f'HandBrake: using preset "{preset_name}" (max picture side {options.max_picture_side}, '
        f"output {out_ext})"
    )

    for input_path in paths:
        if _cancelled():
            emit("Cancelled.")
            return ActionResult(False, "Cancelled.", log)
        output_path = output_path_for_input(input_path, out_ext)
        quality, framerate = overrides_for_input(
            input_path, options.video_quality, options.video_framerate, options.small_file_rule
        )
        if framerate is not None:
            source_fps = probe_source_video_frame_rate(input_path)
            capped = framerate_for_encode(framerate, source_fps)
            if source_fps and capped is not None and capped < framerate:
                emit(
                    f"HandBrake: keeping source {source_fps} fps (requested {framerate} is higher): "
                    f"{input_path.name}"
                )
            framerate = capped

        cmd = _build_encode_cmd(cli, input_path, output_path, preset_name, options, quality, framerate)
        emit("HandBrakeCLI: " + " ".join(f'"{c}"' if " " in c else c for c in cmd))
        rc = _run_handbrake_cmd(cmd, on_output)

        if rc == HANDBRAKE_EXIT_CONTROL_C:
            emit(f"HandBrakeCLI: cancelled or interrupted (exit {rc}). Stopped after:\n{input_path}")
            delete_incomplete_output(output_path, log)
            return ActionResult(False, "Cancelled.", log)
        if rc != 0:
            delete_incomplete_output(output_path, log)
            return ActionResult(
                False, f"HandBrakeCLI exited with code {rc}.\n\nStopped after:\n{input_path}", log
            )

        if options.replace_original:
            if not _size_reduction_exceeds_replace_threshold(input_path, output_path):
                in_mb = input_path.stat().st_size / (1024 * 1024)
                out_mb = output_path.stat().st_size / (1024 * 1024)
                reduction_pct = (in_mb - out_mb) / in_mb * 100 if in_mb > 0 else 0
                emit(
                    f"HandBrake: replace skipped -- {reduction_pct:.1f}% smaller (need >10%). "
                    f"Keeping original and output beside it: {input_path.name} "
                    f"({in_mb:.1f} MB -> {out_mb:.1f} MB, {output_path.name})"
                )
            elif not replace_original_with_output(input_path, output_path, log):
                return ActionResult(
                    False, f"Encode finished but could not replace the original file.\n\n{input_path}", log
                )

    summary = f"HandBrake: finished {len(paths)} file(s)."
    emit(summary)
    return ActionResult(True, summary, log)


# --- CLI (headless "repeat last settings") ---


def resolve_preset_path(preset_file: str) -> Optional[Path]:
    rows = list_preset_json_files()
    if not rows:
        return None
    if preset_file:
        for p in rows:
            if p.name.casefold() == preset_file.casefold():
                return p
    return rows[0]


def run_repeat_last(only_list: Optional[str], only_files: Optional[list[str]]) -> int:
    settings = config_load_settings()
    paths = paths_from_only_list(only_list, only_files)
    paths = expand_paths(paths)
    if not paths:
        print("HandBrake: no files to encode.")
        return 1
    preset_path = resolve_preset_path(settings.preset_file)
    if not preset_path:
        print(f"No preset JSON files found in:\n{PRESET_DIR}")
        return 1
    try:
        options = build_encode_options(settings, preset_path)
    except ValidationError as ex:
        print(str(ex))
        return 1

    def on_output(text: str, replace_last: bool) -> None:
        print(text)

    result = run_encode(paths, options, on_output=on_output)
    _delete_only_list_file(only_list)
    print(result.summary)
    return 0 if result.ok else 1
