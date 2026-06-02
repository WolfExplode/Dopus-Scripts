// FFmpeg Tool — launches FFmpegTool.py (Python / Dear PyGui) from this repo.
//
// Click: open GUI. Selected files and folders fill the file list (folders → all media inside).
// Ctrl+click: run the last GUI action on the selection (no dialog).
// Settings: %APPDATA%\FFmpegTool\settings.json (migrates legacy DOpus_ffmpeg_settings.ini).
//
/** Optional full path to FFmpegTool.py if auto-detect fails. */
var FFMPEG_PY = "";

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function pushItemPath(paths, item, fso) {
    var pathObj = item.realpath;
    pathObj.Resolve();
    var p = trimStr(pathObj + "");
    if (p && (fso.FileExists(p) || fso.FolderExists(p))) {
        paths.push(p);
    }
}

function collectSelectedFilePaths(tab, fso) {
    var paths = [];
    if (!tab) {
        return paths;
    }
    var checkboxMode = false;
    try {
        checkboxMode = tab.selstats.checkbox_mode;
    } catch (e0) {}
    if (checkboxMode) {
        if (tab.selstats.checkeditems === 0) {
            return paths;
        }
        var enChecked = new Enumerator(tab.files);
        for (; !enChecked.atEnd(); enChecked.moveNext()) {
            var checkedItem = enChecked.item();
            if (!checkedItem.checked) {
                continue;
            }
            pushItemPath(paths, checkedItem, fso);
        }
        return paths;
    }
    if (tab.selstats.selitems === 0) {
        return paths;
    }
    var en = new Enumerator(tab.selected);
    for (; !en.atEnd(); en.moveNext()) {
        pushItemPath(paths, en.item(), fso);
    }
    return paths;
}

function mergeUniquePaths(existing, extra) {
    var seen = {};
    var out = [];
    var i;
    for (i = 0; i < existing.length; i++) {
        var key = existing[i].toLowerCase();
        if (!seen[key]) {
            seen[key] = true;
            out.push(existing[i]);
        }
    }
    for (i = 0; i < extra.length; i++) {
        key = extra[i].toLowerCase();
        if (!seen[key]) {
            seen[key] = true;
            out.push(extra[i]);
        }
    }
    return out;
}

function writeOnlyListFile(shell, fso, paths) {
    if (!paths || paths.length === 0) {
        return "";
    }
    var name = "FFmpegTool_only_" + Math.floor(Math.random() * 1000000000) + ".txt";
    var file = fso.BuildPath(shell.ExpandEnvironmentStrings("%TEMP%"), name);
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 2;
    stream.Charset = "utf-8";
    stream.Open();
    var i;
    for (i = 0; i < paths.length; i++) {
        if (i > 0) {
            stream.WriteText("\r\n");
        }
        stream.WriteText(paths[i]);
    }
    stream.SaveToFile(file, 2);
    stream.Close();
    return file;
}

function appendOnlyList(exec, onlyListPath) {
    if (onlyListPath) {
        return exec + " --only-list " + quoteArg(onlyListPath);
    }
    return exec;
}

function appendOnlyFiles(exec, paths, maxArgs) {
    if (!paths || paths.length === 0) {
        return exec;
    }
    var limit = maxArgs || 40;
    if (paths.length > limit) {
        return exec;
    }
    var i;
    for (i = 0; i < paths.length; i++) {
        exec += " --only-file " + quoteArg(paths[i]);
    }
    return exec;
}

function resolveFfmpegPy(shell, fso) {
    if (FFMPEG_PY && fso.FileExists(FFMPEG_PY)) {
        return FFMPEG_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.BuildPath(
                fso.GetParentFolderName(Script.file),
                "..\\FFmpeg\\FFmpegTool.py"
            );
            sibling = fso.GetAbsolutePathName(sibling);
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback =
        "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\FFmpeg\\FFmpegTool.py";
    if (fso.FileExists(fallback)) {
        return fallback;
    }
    return "";
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var ffmpegPy = resolveFfmpegPy(shell, fso);

    if (!ffmpegPy) {
        shell.Popup(
            "FFmpegTool.py not found.\n\n" +
                "Copy it under FFmpeg\\ next to this repo, set FFMPEG_PY in DOpus_ffmpeg.js, " +
                "or install both under Script AddIns.",
            0,
            "FFmpeg Tool",
            16
        );
        return;
    }

    var qualStr = "";
    try {
        qualStr = String(clickData.func.qualifiers + "").toLowerCase();
    } catch (eq) {
        qualStr = "";
    }
    var isCtrl = qualStr.indexOf("ctrl") >= 0;

    var selected = [];
    if (tab) {
        selected = mergeUniquePaths(selected, collectSelectedFilePaths(tab, fso));
    }
    var onlyListPath = writeOnlyListFile(shell, fso, selected);

    if (isCtrl) {
        var execRepeat =
            quoteArg("python") + " " + quoteArg(ffmpegPy) + " --repeat";
        execRepeat = appendOnlyList(execRepeat, onlyListPath);
        if (!onlyListPath) {
            execRepeat = appendOnlyFiles(execRepeat, selected);
        }
        DOpus.Output("FFmpeg Tool (repeat last): " + execRepeat);
        var rc = shell.Run(execRepeat, 1, true);
        DOpus.Output("FFmpeg Tool (repeat last) exit code: " + rc);
        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf) {}
        return;
    }

    var execGui = quoteArg("pythonw") + " " + quoteArg(ffmpegPy) + " --gui";
    execGui = appendOnlyList(execGui, onlyListPath);
    execGui = appendOnlyFiles(execGui, selected);
    if (onlyListPath) {
        DOpus.Output(
            "FFmpeg Tool: " + selected.length + " selected path(s) in file list."
        );
    }
    DOpus.Output("FFmpeg Tool (GUI): " + execGui);
    shell.Run(execGui, 0, false);
}
