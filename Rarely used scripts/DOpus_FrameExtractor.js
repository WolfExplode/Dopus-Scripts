// Timelapse-style frame sampling: extract every N seconds as a short H.264 MP4 (no audio).
// Requires ffmpeg; ffprobe should live beside ffmpeg.exe or on PATH (same as a full FFmpeg install).
// Defaults for the dialog (DOpus_FrameExtractorDlg.xml); user can change values each run.

var DEFAULT_START_OFFSET_SEC = 3; // first sampled frame ≈ this many seconds in
var DEFAULT_INTERVAL_SEC = 2; // seconds between samples
var DEFAULT_FRAME_COUNT = 48; // how many output frames (then stop)

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function popup(shell, text, title, flags) {
    shell.Popup(String(text), 0, String(title), flags == null ? 48 : flags);
}

function pathsEqualIgnoreCase(a, b) {
    return String(a).toLowerCase() === String(b).toLowerCase();
}

function outputPathForTimelapse(fso, inputPath) {
    var folder = fso.GetParentFolderName(inputPath);
    var base = fso.GetBaseName(inputPath);
    var candidate = folder + "\\" + base + "_frames.mp4";
    if (pathsEqualIgnoreCase(candidate, inputPath)) {
        return folder + "\\" + base + "_timelapse.mp4";
    }
    return candidate;
}

/** Resolve ffmpeg.exe / ffprobe.exe: PATH first, then common install dirs. */
function resolveTool(shell, fso, exeName) {
    var tmp = shell.ExpandEnvironmentStrings("%TEMP%") + "\\DOpus_where_" + exeName.replace(/[^a-z0-9]/gi, "_") + ".txt";
    try {
        if (fso.FileExists(tmp)) fso.DeleteFile(tmp);
    } catch (e0) {}
    var cmd =
        'cmd /c where ' +
        quoteArg(exeName) +
        ' 1> "' +
        tmp +
        '" 2>nul';
    shell.Run(cmd, 0, true);
    if (fso.FileExists(tmp)) {
        try {
            var ts = fso.OpenTextFile(tmp, 1);
            var first = ts.ReadLine().replace(/^\s+|\s+$/g, "");
            ts.Close();
            fso.DeleteFile(tmp);
            if (first && fso.FileExists(first)) return first;
        } catch (e1) {
            try {
                fso.DeleteFile(tmp);
            } catch (e2) {}
        }
    }
    var candidates = [
        shell.ExpandEnvironmentStrings("%ProgramFiles%\\ffmpeg\\bin\\" + exeName),
        shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\\ffmpeg\\bin\\" + exeName)
    ];
    var i;
    for (i = 0; i < candidates.length; i++) {
        if (fso.FileExists(candidates[i])) return candidates[i];
    }
    return "";
}

function parseFrameRateToFloat(s) {
    s = String(s).replace(/^\s+|\s+$/g, "").replace(/\r|\n/g, "");
    if (!s) return NaN;
    var slash = s.indexOf("/");
    if (slash > 0) {
        var num = parseFloat(s.substring(0, slash));
        var den = parseFloat(s.substring(slash + 1));
        if (den && !isNaN(num)) return num / den;
    }
    var v = parseFloat(s);
    return isNaN(v) ? NaN : v;
}

/** ffprobe.exe next to ffmpeg.exe (same install); PATH-based setups often lack a separate ffprobe on PATH. */
function ffprobeBesideFfmpeg(fso, ffmpegPath) {
    if (!ffmpegPath) return "";
    var p = fso.GetParentFolderName(ffmpegPath) + "\\ffprobe.exe";
    return fso.FileExists(p) ? p : "";
}

/** Cmd prefix: bare name uses PATH like DOpus_ffmpeg.js; full path is quoted. */
function ffprobeCmdToken(exePath) {
    if (exePath.indexOf("\\") >= 0 || exePath.indexOf("/") >= 0 || exePath.indexOf(":") >= 0) {
        return quoteArg(exePath);
    }
    return exePath;
}

function readTempFileAll(fso, path) {
    try {
        if (!fso.FileExists(path)) return "";
        var ts = fso.OpenTextFile(path, 1);
        var t = ts.ReadAll();
        ts.Close();
        return t;
    } catch (e) {
        return "";
    }
}

/**
 * Pick first line that parses as a positive frame rate (fraction or decimal).
 */
function extractFpsLineFromProbeOutput(text) {
    var lines = String(text).split(/\r?\n/);
    var i;
    for (i = 0; i < lines.length; i++) {
        var t = lines[i].replace(/^\s+|\s+$/g, "");
        if (!t) continue;
        var v = parseFrameRateToFloat(t);
        if (isFinite(v) && v > 0) return t;
    }
    return "";
}

/**
 * Run ffprobe for one stream field. Ignores exit code if stdout parse succeeds (ffprobe often returns non-zero when stderr has noise).
 */
function probeVideoField(shell, fso, ffprobeExe, mediaPath, streamField, tmpBase) {
    var tmp = shell.ExpandEnvironmentStrings("%TEMP%") + "\\" + tmpBase + ".txt";
    try {
        if (fso.FileExists(tmp)) fso.DeleteFile(tmp);
    } catch (e0) {}
    var cmd =
        "cmd /c " +
        ffprobeCmdToken(ffprobeExe) +
        " -v error -select_streams v:0 -show_entries stream=" +
        streamField +
        " -of default=noprint_wrappers=1:nokey=1 " +
        quoteArg(mediaPath) +
        " 1> \"" +
        tmp +
        '" 2>&1';
    try {
        shell.Run(cmd, 0, true);
    } catch (ex) {
        try {
            if (fso.FileExists(tmp)) fso.DeleteFile(tmp);
        } catch (e1) {}
        return "";
    }
    var raw = readTempFileAll(fso, tmp);
    try {
        if (fso.FileExists(tmp)) fso.DeleteFile(tmp);
    } catch (eD) {}
    return extractFpsLineFromProbeOutput(raw);
}

/**
 * Try sibling ffprobe, then resolved path, then bare ffprobe.exe on PATH (matches other DOpus scripts).
 */
function probeVideoFrameRate(shell, fso, ffmpegPath, mediaPath) {
    var candidates = [];
    var sib = ffprobeBesideFfmpeg(fso, ffmpegPath);
    if (sib) candidates.push(sib);
    var resolved = resolveTool(shell, fso, "ffprobe.exe");
    if (resolved) candidates.push(resolved);
    candidates.push("ffprobe.exe");

    var seen = {};
    var i;
    var fields = ["r_frame_rate", "avg_frame_rate"];
    for (i = 0; i < candidates.length; i++) {
        var exe = candidates[i];
        var key = String(exe).toLowerCase();
        if (seen[key]) continue;
        seen[key] = true;
        if (
            (exe.indexOf("\\") >= 0 || exe.indexOf("/") >= 0 || exe.indexOf(":") >= 0) &&
            !fso.FileExists(exe)
        ) {
            continue;
        }

        var fi;
        for (fi = 0; fi < fields.length; fi++) {
            var line = probeVideoField(
                shell,
                fso,
                exe,
                mediaPath,
                fields[fi],
                "DOpus_frame_fps_" + i + "_" + fi
            );
            if (line) return line;
        }
    }
    return "";
}

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

/**
 * After OK on DOpus_FrameExtractorDlg; returns { startSec, intervalSec, count } or null (after popup on bad input).
 */
function parseFrameExtractParams(shell, dlg) {
    var startSec = parseFloat(trimStr(dlg.control("start_edit").value));
    var intervalSec = parseFloat(trimStr(dlg.control("interval_edit").value));
    var countStr = trimStr(dlg.control("frame_count_edit").value);
    if (!isFinite(startSec) || startSec < 0) {
        popup(shell, "Start offset must be a number ≥ 0.", "Frame extract", 16);
        return null;
    }
    if (!isFinite(intervalSec) || intervalSec <= 0) {
        popup(shell, "Interval must be a number greater than 0.", "Frame extract", 16);
        return null;
    }
    if (!/^\d+$/.test(countStr)) {
        popup(shell, "Frame count must be a whole number (e.g. 48).", "Frame extract", 16);
        return null;
    }
    var countNum = parseInt(countStr, 10);
    if (countNum < 1) {
        popup(shell, "Frame count must be at least 1.", "Frame extract", 16);
        return null;
    }
    return { startSec: startSec, intervalSec: intervalSec, count: countNum };
}

/**
 * Same pattern as DOpus_yt-dlp.js / gallery-dl: detach + GetMsg loop.
 * A plain Show() without detach often never displays the dialog in DOpus.
 */
function showFrameExtractDialog(clickData, shell) {
    var dlg = DOpus.dlg;
    dlg.window = clickData.func.sourcetab;
    dlg.template = "DOpus_FrameExtractorDlg";
    dlg.detach = true;
    try {
        dlg.Create();
    } catch (eCreate) {
        popup(
            shell,
            "Could not create dialog (XML missing or wrong name?).\n\n" +
                "Button → Resources → dialog resource must be named exactly:\nDOpus_FrameExtractorDlg\n\n" +
                String(eCreate.message || eCreate),
            "Frame extract",
            16
        );
        return null;
    }
    try {
        dlg.control("start_edit").value = String(DEFAULT_START_OFFSET_SEC);
        dlg.control("interval_edit").value = String(DEFAULT_INTERVAL_SEC);
        dlg.control("frame_count_edit").value = String(DEFAULT_FRAME_COUNT);
    } catch (eCtl) {
        popup(
            shell,
            "Dialog opened but controls are missing.\n\n" +
                "Paste DOpus_FrameExtractorDlg.xml into the button Resources; " +
                "check control names start_edit, interval_edit, frame_count_edit.\n\n" +
                String(eCtl.message || eCtl),
            "Frame extract",
            16
        );
        return null;
    }
    dlg.Show();
    while (true) {
        var msg = dlg.GetMsg();
        if (!msg.result) {
            break;
        }
    }
    if (dlg.result != "1") {
        return null;
    }
    return parseFrameExtractParams(shell, dlg);
}

function buildFfmpegCmd(ffmpegPath, inputPath, outputPath, fpsFrac, startFrame, endFrame, stepFrames, frameCount) {
    var vf =
        "select='gte(n," +
        startFrame +
        ")*lt(n," +
        endFrame +
        ")*not(mod(n-" +
        startFrame +
        "," +
        stepFrames +
        "))',setpts=N/FRAME_RATE/TB";
    return (
        quoteArg(ffmpegPath) +
        " -hide_banner -loglevel warning -stats -y " +
        "-i " +
        quoteArg(inputPath) +
        " " +
        "-vf " +
        quoteArg(vf) +
        " " +
        "-frames:v " +
        frameCount +
        " " +
        "-c:v libx264 -crf 18 -preset medium " +
        "-r " +
        quoteArg(fpsFrac) +
        " " +
        "-pix_fmt yuv420p -an -movflags +faststart " +
        quoteArg(outputPath)
    );
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    if (!tab) {
        popup(shell, "No source folder tab.", "Frame extract", 16);
        return;
    }
    var fso = new ActiveXObject("Scripting.FileSystemObject");

    if (tab.selstats.selfiles === 0) {
        popup(shell, "Select one or more video files.", "Frame extract", 48);
        return;
    }

    var paths = [];
    var selectedFiles = tab.selected_files;
    var en = new Enumerator(selectedFiles);
    for (; !en.atEnd(); en.moveNext()) {
        paths.push(en.item().realpath + "");
    }
    if (paths.length === 0) {
        popup(shell, "No file paths from selection.", "Frame extract", 16);
        return;
    }

    var filePaths = [];
    var i;
    for (i = 0; i < paths.length; i++) {
        var p = paths[i];
        if (!fso.FileExists(p)) continue;
        try {
            fso.GetFile(p);
            filePaths.push(p);
        } catch (eSkip) {
            DOpus.Output("Frame extract: skip (not a file): " + p);
        }
    }
    if (filePaths.length === 0) {
        popup(shell, "No files in selection (folders are skipped).", "Frame extract", 48);
        return;
    }

    var params = showFrameExtractDialog(clickData, shell);
    if (!params) {
        return;
    }
    var startSec = params.startSec;
    var intervalSec = params.intervalSec;
    var count = params.count;

    var ffmpegPath = resolveTool(shell, fso, "ffmpeg.exe");
    if (!ffmpegPath) {
        popup(
            shell,
            "ffmpeg.exe not found. Install FFmpeg and ensure it is on PATH,\n" +
                "or install under Program Files\\ffmpeg\\bin.",
            "Frame extract",
            16
        );
        return;
    }

    for (i = 0; i < filePaths.length; i++) {
        var inputPath = filePaths[i];

        var fpsFrac = probeVideoFrameRate(shell, fso, ffmpegPath, inputPath);
        if (!fpsFrac) {
            DOpus.Output(
                "Frame extract: ffprobe could not read FPS (no video stream, or ffprobe missing next to ffmpeg / on PATH): " +
                    inputPath
            );
            popup(
                shell,
                "Could not read video frame rate.\n\n" +
                    "ffprobe must be available (usually next to ffmpeg in the same folder), and the file needs a video stream.\n\n" +
                    inputPath,
                "Frame extract",
                16
            );
            return;
        }

        var fpsVal = parseFrameRateToFloat(fpsFrac);
        if (!isFinite(fpsVal) || fpsVal <= 0) {
            popup(shell, "Invalid frame rate from ffprobe: " + fpsFrac + "\n" + inputPath, "Frame extract", 16);
            return;
        }

        var startFrame = Math.round(startSec * fpsVal);
        var stepFrames = Math.round(intervalSec * fpsVal);
        var endFrame = startFrame + count * stepFrames;

        if (stepFrames <= 0) {
            popup(
                shell,
                "Interval too small for this frame rate (step frames = 0).\n" + inputPath,
                "Frame extract",
                16
            );
            return;
        }

        var outputPath = outputPathForTimelapse(fso, inputPath);

        DOpus.Output(
            "Frame extract: FPS=" +
                fpsFrac +
                " (~" +
                fpsVal +
                ") startFrame=" +
                startFrame +
                " step=" +
                stepFrames +
                " endFrame=" +
                endFrame +
                " frames=" +
                count
        );
        DOpus.Output("  in:  " + inputPath);
        DOpus.Output("  out: " + outputPath);

        var cmd = buildFfmpegCmd(
            ffmpegPath,
            inputPath,
            outputPath,
            fpsFrac,
            startFrame,
            endFrame,
            stepFrames,
            count
        );
        DOpus.Output("ffmpeg: " + cmd);

        var rc = shell.Run(cmd, 1, true);
        if (rc !== 0) {
            popup(
                shell,
                "ffmpeg exited with code " + rc + ".\n\nStopped at:\n" + inputPath,
                "Frame extract",
                16
            );
            return;
        }
    }

    DOpus.Output("Frame extract: finished " + filePaths.length + " file(s).");
    try {
        clickData.func.command.RunCommand("Go REFRESH");
    } catch (eRf) {}
}
