# Dopus-Scripts

Directory Opus buttons backed by small Python desktop tools. Every tool is its own
window built on one shared PySide6 UI kit, so they all look and behave the same:
a rail of pages on the left, the page's options in the middle, and an activity
log on the right that streams the underlying program's output.

## Tools

| Folder | Window | What it does | Opus button |
|---|---|---|---|
| `FFmpeg/` | FFmpeg Tool | Convert, cut, rotate/flip, cover art, audio streams, merge with chapters, timelapse frame sampling | `DOpus_ffmpeg.js`, `Rarely used scripts/DOpus_FrameExtractor.js` (opens the Timelapse page) |
| `Handbrake/` | HandBrake Tool | Batch encode with a preset JSON from the folder, with quality/size/range overrides | `DOpus_handbrake.js` |
| `ImageConverter/` | Image Converter | Convert (ImageMagick, cjxl, texconv for Skyrim DDS) and resize images | `DOpus_image_converter.js` |
| `Organize Files/` | Organize Files | Match marks, title cleanup, filename tags, JPG transfer, copy from list, text compare | `DOpus_OrganizeFiles.js` |
| `TranslateFilename/` | Translate Filename | Translate file names to English (DeepSeek or Kimi), review, rename, revert | `DOpus_translate_filename.js` |
| `YtDlp/` | yt-dlp | Download the clipboard URL (or a pasted `yt-dlp …` command) into the current folder | `UserCommands/DOpus_yt-dlp.js` |
| `GalleryDl/` | Pinterest | Scrape a Pinterest profile or board with gallery-dl | `UserCommands/DOpus_gallery-dl_pinterest.js` |

The other `UserCommands/*.js` buttons launch external programs and have no UI of their own.

Common behaviour:

- **Click** in Opus opens the window with the selection loaded. **Ctrl+click** repeats the
  last action headlessly (FFmpeg, HandBrake, Image Converter, Organize), renames straight
  away (Translate) or downloads the clipboard URL (yt-dlp).
- Drop files or folders from Explorer onto any file list; Ctrl+V pastes paths; Delete removes
  the selected rows; double-click shows a file in Explorer.
- **Ctrl+Enter** runs the open page's main action; **Ctrl+1…9** switches pages.
- **Stop** in the activity panel cancels a running job; FFmpeg, HandBrake and the Image
  Converter also remove the half-written output.
- Settings live in `%APPDATA%\<ToolName>\settings.json`; window size, splitters and the
  last page are kept in `ui.json` next to it.

## Setup

Python 3.10+ on Windows. Install each tool's requirements, for example:

```powershell
pip install -r FFmpeg\requirements.txt
```

All tools need `PySide6-Essentials`. External programs: FFmpeg/ffprobe on `PATH`,
HandBrakeCLI under Program Files\HandBrake, portable ImageMagick 7 (folder set on the
Image Converter's Encoders page), yt-dlp, gallery-dl.

The Opus scripts find each tool next to themselves (`..\<Tool>\<Tool>.py`) and fall
back to `C:\Users\WXP\Documents\GitHub\Dopus-Scripts\…`; each script has a variable at
the top to point it elsewhere.

## Layout

```
Shared/
  uikit/              PySide6 UI kit used by every window
    theme.py          palette + stylesheet (BetterGhub-style dark theme, one accent per tool)
    shell.py          ToolWindow: rail, header, pages, inputs/activity splitters, toasts
    widgets.py        Card, field/row helpers, Segmented, OptionRow, ActionRow, Pill, buttons
    paths.py          PathList (drag-and-drop file list), PathField, file pickers
    console.py        ConsolePanel: streaming log with status pill, progress and Stop
    jobs.py           JobRunner / JobHost: run work on a thread, stream output, report results
  process_runner.py   run a console program with live output and cancel (downloaders)
  recycle_delete.py   send files to the Recycle Bin
<Tool>/
  <Tool>Tool.py       entry point: CLI flags, else the GUI
  *_logic.py          the work (no UI imports)
  *_gui.py            the window, built from uikit
```

For a design review, set `UIKIT_SNAPSHOT` to a folder and launch a tool: it renders every
page to `<folder>\<page>.png` and exits without touching saved settings.

```powershell
$env:UIKIT_SNAPSHOT = "$env:TEMP\snaps"; python FFmpeg\FFmpegTool.py
```
