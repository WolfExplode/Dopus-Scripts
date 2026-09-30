// yt-dlp Downloader: launches YtDlp\YtDlpTool.py (Python / PySide6) from this repo.
//
// Click: open the downloader window. The URL (or a whole "yt-dlp ... URL" command)
//        is read from the clipboard; files are saved into the current folder.
// Ctrl+click: download the clipboard URL right away with the saved settings, in a
//        console window at the bottom-right of the screen (no dialog).
// Settings: %APPDATA%\YtDlpTool\settings.json (migrates the old DOpus_ytdlp_settings.ini).

/** Optional full path to YtDlpTool.py if auto-detect fails. */
var YTDLP_PY = "";

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function resolveToolPy(fso) {
    if (YTDLP_PY && fso.FileExists(YTDLP_PY)) {
        return YTDLP_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.GetAbsolutePathName(
                fso.BuildPath(fso.GetParentFolderName(Script.file), "..\\YtDlp\\YtDlpTool.py")
            );
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback = "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\YtDlp\\YtDlpTool.py";
    return fso.FileExists(fallback) ? fallback : "";
}

function OnClick(clickData) {
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var toolPy = resolveToolPy(fso);
    if (!toolPy) {
        shell.Popup("YtDlpTool.py not found.\n\nSet YTDLP_PY in DOpus_yt-dlp.js.", 0, "yt-dlp", 16);
        return;
    }

    var dest = "";
    try {
        dest = String(clickData.func.sourcetab.path + "");
    } catch (eTab) {}
    // A trailing backslash would escape the closing quote on the command line (C:\ -> C:\.).
    if (/\\$/.test(dest)) {
        dest += ".";
    }

    var qualStr = "";
    try {
        qualStr = String(clickData.func.qualifiers + "").toLowerCase();
    } catch (eq) {}

    if (qualStr.indexOf("ctrl") >= 0) {
        var execRepeat = quoteArg("python") + " " + quoteArg(toolPy) + " --repeat --dest " + quoteArg(dest);
        DOpus.Output("yt-dlp (Ctrl+click, saved settings): " + execRepeat);
        shell.Run(execRepeat, 1, false);
        return;
    }

    var execGui = quoteArg("pythonw") + " " + quoteArg(toolPy) + " --gui --dest " + quoteArg(dest);
    DOpus.Output("yt-dlp: " + execGui);
    shell.Run(execGui, 0, false);
}
