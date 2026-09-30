// Frame extract (timelapse sample): now part of the FFmpeg Tool.
//
// Opens FFmpeg\FFmpegTool.py on its Timelapse page with the selected videos loaded.
// Set the start offset, interval and frame count there; it writes <name>_frames.mp4
// (H.264, no audio) next to each video.

/** Optional full path to FFmpegTool.py if auto-detect fails. */
var FFMPEG_PY = "";

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function resolveToolPy(fso) {
    if (FFMPEG_PY && fso.FileExists(FFMPEG_PY)) {
        return FFMPEG_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.GetAbsolutePathName(
                fso.BuildPath(fso.GetParentFolderName(Script.file), "..\\FFmpeg\\FFmpegTool.py")
            );
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback = "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\FFmpeg\\FFmpegTool.py";
    return fso.FileExists(fallback) ? fallback : "";
}

/** UTF-8 list file named like DOpus_ffmpeg.js's, so the FFmpeg Tool deletes it on exit. */
function writeOnlyListFile(shell, fso, paths) {
    var name = "FFmpegTool_only_" + Math.floor(Math.random() * 1000000000) + ".txt";
    var file = fso.BuildPath(shell.ExpandEnvironmentStrings("%TEMP%"), name);
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 2;
    stream.Charset = "utf-8";
    stream.Open();
    for (var i = 0; i < paths.length; i++) {
        if (i > 0) {
            stream.WriteText("\r\n");
        }
        stream.WriteText(paths[i]);
    }
    stream.SaveToFile(file, 2);
    stream.Close();
    return file;
}

function OnClick(clickData) {
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var toolPy = resolveToolPy(fso);
    if (!toolPy) {
        shell.Popup("FFmpegTool.py not found.\n\nSet FFMPEG_PY in DOpus_FrameExtractor.js.", 0, "Frame extract", 16);
        return;
    }

    var paths = [];
    var tab = clickData.func.sourcetab;
    if (tab && tab.selstats.selfiles > 0) {
        var en = new Enumerator(tab.selected_files);
        for (; !en.atEnd(); en.moveNext()) {
            var p = String(en.item().realpath + "");
            if (fso.FileExists(p)) {
                paths.push(p);
            }
        }
    }

    var exec = quoteArg("pythonw") + " " + quoteArg(toolPy) + " --gui --section timelapse";
    if (paths.length > 0) {
        exec += " --only-list " + quoteArg(writeOnlyListFile(shell, fso, paths));
    }
    DOpus.Output("Frame extract: " + exec);
    shell.Run(exec, 0, false);
}
