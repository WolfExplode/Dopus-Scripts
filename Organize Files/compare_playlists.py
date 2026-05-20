"""Compare a streaming playlist export to a local music file list (exact + fuzzy).

Library API: run_compare(), format_compare_report()
CLI: python compare_playlists.py playlist.txt library.txt [-o output_dir]
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from rapidfuzz import fuzz

AUDIO_EXTS = {".mp3", ".opus", ".m4a", ".flac", ".wav", ".ogg", ".aac"}
FUZZY_THRESHOLD = 80


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("utf-8", errors="replace")


def read_lines(path: Path) -> list[str]:
    text = read_text(path)
    lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if re.match(r"^\d+\|", line):
            line = line.split("|", 1)[1].strip()
        lines.append(line)
    return lines


def strip_audio_ext(name: str) -> str:
    path = Path(name)
    if path.suffix.lower() in AUDIO_EXTS:
        return path.stem
    return name


def split_artist_title(line: str) -> tuple[str, str] | None:
    m = re.match(r"^(.+?)\s+-\s*(.+)$", line)
    if not m:
        return None
    return m.group(1).strip(), m.group(2).strip()


def squash(text: str) -> str:
    return re.sub(r"[\s\-_/\\.:]+", "", text.lower())


def normalize_text(text: str) -> str:
    text = text.lower()
    text = text.replace("&", "and")
    text = re.sub(r"\bfeat\.?\b", "", text)
    text = re.sub(r"\bft\.?\b", "", text)
    text = re.sub(r"[\(\)\[\]\"''`]", " ", text)
    text = re.sub(r"[/\\|:]+", " ", text)
    text = re.sub(r"[_]+", " ", text)
    text = re.sub(r"[^a-z0-9\u3040-\u30ff\u4e00-\u9fff\s]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def artist_tokens(artist: str) -> list[str]:
    parts = re.split(r"[,;&+]|\band\b", artist, flags=re.I)
    tokens = []
    for part in parts:
        n = normalize_text(part)
        if n:
            tokens.append(n)
    return tokens or [normalize_text(artist)]


def track_keys(artist: str, title: str) -> set[str]:
    a = normalize_text(artist)
    t = normalize_text(title)
    keys = {
        f"{a}|{t}",
        f"{squash(a)}|{squash(t)}",
    }
    for token in artist_tokens(artist):
        keys.add(f"{token}|{t}")
        keys.add(f"{squash(token)}|{squash(t)}")
    keys.add(t)
    keys.add(squash(t))
    return keys


class Track:
    def __init__(self, raw: str, artist: str, title: str):
        self.raw = raw
        self.artist = artist
        self.title = title
        self.keys = track_keys(artist, title)
        self.compare = normalize_text(f"{artist} - {title}")


def load_playlist(path: Path) -> list[Track]:
    tracks = []
    for line in read_lines(path):
        parsed = split_artist_title(line)
        if not parsed:
            continue
        artist, title = parsed
        tracks.append(Track(line, artist, title))
    return tracks


def load_library(path: Path) -> list[Track]:
    tracks = []
    for line in read_lines(path):
        base = strip_audio_ext(line)
        parsed = split_artist_title(base)
        if not parsed:
            continue
        artist, title = parsed
        tracks.append(Track(line, artist, title))
    return tracks


def build_index(tracks: list[Track]) -> dict[str, list[Track]]:
    index: dict[str, list[Track]] = {}
    for track in tracks:
        for key in track.keys:
            index.setdefault(key, []).append(track)
    return index


def find_exact_match(track: Track, index: dict[str, list[Track]]) -> Track | None:
    for key in track.keys:
        hits = index.get(key)
        if hits:
            return hits[0]
    return None


def subsequence_coverage(needle: str, haystack: str) -> float:
    """Share of needle characters found in order inside haystack."""
    if not needle:
        return 100.0
    n = needle.lower()
    h = haystack.lower()
    j = 0
    for c in h:
        if j < len(n) and c == n[j]:
            j += 1
    return 100.0 * j / len(n)


def fuzzy_score(a: str, b: str) -> float:
    return max(
        fuzz.ratio(a, b),
        fuzz.partial_ratio(a, b),
        subsequence_coverage(a, b),
        subsequence_coverage(b, a),
    )


def find_fuzzy_match(
    track: Track, library: list[Track], used: set[int], threshold: float
) -> tuple[Track, float] | None:
    best: Track | None = None
    best_score = -1.0
    for cand in library:
        if id(cand) in used:
            continue
        score = fuzzy_score(track.compare, cand.compare)
        if score >= threshold and score > best_score:
            best_score = score
            best = cand
    if best is None:
        return None
    return best, best_score


def match_playlist_to_library(
    playlist: list[Track],
    library: list[Track],
    library_index: dict[str, list[Track]],
    threshold: float,
) -> tuple[list[Track], list[tuple[Track, Track, float]], set[int]]:
    """Return (unmatched playlist tracks, fuzzy pairs, used library track ids)."""
    used_lib: set[int] = set()
    unmatched: list[Track] = []
    fuzzy_pairs: list[tuple[Track, Track, float]] = []

    for track in playlist:
        hit = find_exact_match(track, library_index)
        if hit is not None and id(hit) not in used_lib:
            used_lib.add(id(hit))
            continue

        fuzzy = find_fuzzy_match(track, library, used_lib, threshold)
        if fuzzy is not None:
            lib_track, score = fuzzy
            used_lib.add(id(lib_track))
            fuzzy_pairs.append((track, lib_track, score))
            continue

        unmatched.append(track)

    return unmatched, fuzzy_pairs, used_lib


def write_list(path: Path, tracks: list[Track]) -> None:
    path.write_text(
        "\n".join(t.raw for t in tracks) + ("\n" if tracks else ""),
        encoding="utf-8",
    )


def write_fuzzy_matches(
    path: Path, pairs: list[tuple[Track, Track, float]]
) -> None:
    lines = []
    for playlist_track, lib_track, score in pairs:
        lines.append(f"{score:.0f}%  {playlist_track.raw}")
        lines.append(f"       -> {lib_track.raw}")
    path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")


@dataclass
class CompareResult:
    playlist_path: Path
    library_path: Path
    output_dir: Path
    threshold: int
    playlist_count: int
    library_count: int
    exact_matched: int
    missing_files: list[Track]
    extra_files: list[Track]
    fuzzy_pairs: list[tuple[Track, Track, float]]
    missing_path: Path
    extra_path: Path
    fuzzy_path: Path
    error: str | None = None

    @property
    def total_matched(self) -> int:
        return self.playlist_count - len(self.missing_files)

    @property
    def fuzzy_matched(self) -> int:
        return len(self.fuzzy_pairs)


def run_compare(
    playlist_path: Path,
    library_path: Path,
    *,
    output_dir: Path | None = None,
    threshold: int = FUZZY_THRESHOLD,
    write_reports: bool = True,
) -> CompareResult:
    """Compare two list files; optionally write report txt files under output_dir."""
    out_default = output_dir or playlist_path.parent
    base = CompareResult(
        playlist_path=playlist_path,
        library_path=library_path,
        output_dir=out_default,
        threshold=threshold,
        playlist_count=0,
        library_count=0,
        exact_matched=0,
        missing_files=[],
        extra_files=[],
        fuzzy_pairs=[],
        missing_path=out_default / "missing-from-library.txt",
        extra_path=out_default / "extra-not-in-playlist.txt",
        fuzzy_path=out_default / "fuzzy-matches.txt",
    )

    if not playlist_path.is_file():
        base.error = f"Playlist file not found:\n{playlist_path}"
        return base
    if not library_path.is_file():
        base.error = f"Library file not found:\n{library_path}"
        return base

    playlist = load_playlist(playlist_path)
    library = load_library(library_path)
    library_index = build_index(library)

    missing_files, fuzzy_pairs, used_lib = match_playlist_to_library(
        playlist, library, library_index, float(threshold)
    )
    extra_files = [t for t in library if id(t) not in used_lib]
    exact_matched = len(playlist) - len(missing_files) - len(fuzzy_pairs)

    out_dir = output_dir or playlist_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    missing_path = out_dir / "missing-from-library.txt"
    extra_path = out_dir / "extra-not-in-playlist.txt"
    fuzzy_path = out_dir / "fuzzy-matches.txt"

    if write_reports:
        write_list(missing_path, missing_files)
        write_list(extra_path, extra_files)
        write_fuzzy_matches(fuzzy_path, fuzzy_pairs)

    return CompareResult(
        playlist_path=playlist_path,
        library_path=library_path,
        output_dir=out_dir,
        threshold=threshold,
        playlist_count=len(playlist),
        library_count=len(library),
        exact_matched=exact_matched,
        missing_files=missing_files,
        extra_files=extra_files,
        fuzzy_pairs=fuzzy_pairs,
        missing_path=missing_path,
        extra_path=extra_path,
        fuzzy_path=fuzzy_path,
    )


def format_compare_report(result: CompareResult, *, preview_limit: int = 15) -> str:
    if result.error:
        return result.error

    lines = [
        "Playlist compare\n",
        f"Playlist:  {result.playlist_path}",
        f"Library:   {result.library_path}",
        f"Reports:   {result.output_dir}",
        f"Threshold: {result.threshold}%",
        "",
        f"Playlist tracks: {result.playlist_count}",
        f"Library files:   {result.library_count}",
        f"Exact matched:   {result.exact_matched}",
        f"Fuzzy matched:   {result.fuzzy_matched}",
        f"Total matched:   {result.total_matched}",
        "",
        f"In playlist but no matching file ({len(result.missing_files)}):",
        f"  -> {result.missing_path}",
    ]
    for t in result.missing_files[:preview_limit]:
        lines.append(f"     {t.raw}")
    if len(result.missing_files) > preview_limit:
        lines.append(f"     ... and {len(result.missing_files) - preview_limit} more")

    lines.append("")
    lines.append(f"On disk but not in playlist ({len(result.extra_files)}):")
    lines.append(f"  -> {result.extra_path}")
    for t in result.extra_files[:preview_limit]:
        lines.append(f"     {t.raw}")
    if len(result.extra_files) > preview_limit:
        lines.append(f"     ... and {len(result.extra_files) - preview_limit} more")

    if result.fuzzy_pairs:
        lines.append("")
        lines.append(f"Fuzzy matches ({len(result.fuzzy_pairs)}):")
        lines.append(f"  -> {result.fuzzy_path}")
        for playlist_track, lib_track, score in result.fuzzy_pairs[:preview_limit]:
            lines.append(f"  {score:.0f}%  {playlist_track.raw}")
            lines.append(f"       -> {lib_track.raw}")
        if len(result.fuzzy_pairs) > preview_limit:
            lines.append(f"     ... and {len(result.fuzzy_pairs) - preview_limit} more")

    return "\n".join(lines) + "\n"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Find songs missing from each side of two playlist lists."
    )
    p.add_argument(
        "playlist",
        type=Path,
        help="Playlist export (Artist - Title per line, e.g. www.txt)",
    )
    p.add_argument(
        "library",
        type=Path,
        help="Local file list (Artist - Title.ext per line, e.g. Music.txt)",
    )
    p.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Write report files here (default: same folder as playlist)",
    )
    p.add_argument(
        "--threshold",
        type=int,
        default=FUZZY_THRESHOLD,
        help=f"Fuzzy match minimum score (default: {FUZZY_THRESHOLD})",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    args = parse_args(argv)
    result = run_compare(
        args.playlist,
        args.library,
        output_dir=args.output_dir,
        threshold=args.threshold,
    )
    text = format_compare_report(result)
    try:
        print(text.rstrip("\n"))
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(
            text.encode(enc, errors="replace")
            .decode(enc, errors="replace")
            .rstrip("\n")
        )
    return 1 if result.error else 0


if __name__ == "__main__":
    raise SystemExit(main())
