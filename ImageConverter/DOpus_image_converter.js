// Image Converter — launches ImageConverterTool.py (Python / Dear PyGui) from this repo.
//
// Click: open GUI. Selected files and folders fill the input list.
// No selection: current tab folder is passed as the default input.
// Ctrl+click: convert immediately with last saved settings (no dialog).
//
// Settings: %APPDATA%\ImageConverter\settings.json

/** Optional full path to ImageConverterTool.py if auto-detect fails. */
var CONVERTER_PY = "";

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function tabFolderPath(tab, fso) {
    if (!tab || !tab.path) {
        return "";
    }
    var pathObj = tab.path;
    try {
        pathObj.Resolve();
    } catch (e0) {}
    var p = trimStr(pathObj + "");
    if (p && fso.FolderExists(p)) {
        return p;
    }
    return "";
}

function pushItemPath(paths, item, fso) {
    var pathObj = item.realpath;
    pathObj.Resolve();
    var p = trimStr(pathObj + "");
    if (p && (fso.FileExists(p) || fso.FolderExists(p))) {
        paths.push(p);
    }
}

function collectSelectedPaths(tab, fso) {
    var paths = [];
    if (!tab) {
        return paths;
    }

    var en;
    en = new Enumerator(tab.selected_files);
    for (; !en.atEnd(); en.moveNext()) {
        pushItemPath(paths, en.item(), fso);
    }
    en = new Enumerator(tab.selected_dirs);
    for (; !en.atEnd(); en.moveNext()) {
        pushItemPath(paths, en.item(), fso);
    }

    if (paths.length === 0) {
        try {
            if (tab.selstats.selitems > 0) {
                en = new Enumerator(tab.selected);
                for (; !en.atEnd(); en.moveNext()) {
                    pushItemPath(paths, en.item(), fso);
                }
            }
        } catch (e1) {}
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
            seen[key] = 1;
            out.push(existing[i]);
        }
    }
    for (i = 0; i < extra.length; i++) {
        key = extra[i].toLowerCase();
        if (!seen[key]) {
            seen[key] = 1;
            out.push(extra[i]);
        }
    }
    return out;
}

function writeOnlyListFile(shell, fso, paths) {
    if (!paths || paths.length === 0) {
        return "";
    }
    var name = "ImageConverter_only_" + Math.floor(Math.random() * 1000000000) + ".txt";
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

function resolveConverterPy(shell, fso) {
    if (CONVERTER_PY && fso.FileExists(CONVERTER_PY)) {
        return CONVERTER_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.BuildPath(
                fso.GetParentFolderName(Script.file),
                "ImageConverterTool.py"
            );
            sibling = fso.GetAbsolutePathName(sibling);
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback =
        "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\ImageConverter\\ImageConverterTool.py";
    if (fso.FileExists(fallback)) {
        return fallback;
    }
    return "";
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var converterPy = resolveConverterPy(shell, fso);

    if (!converterPy) {
        shell.Popup(
            "ImageConverterTool.py not found.\n\n" +
                "Copy it under ImageConverter\\ next to this repo, set CONVERTER_PY in DOpus_image_converter.js, " +
                "or install both under Script AddIns.",
            0,
            "Image Converter",
            16
        );
        return;
    }

    var tabFolder = tab ? tabFolderPath(tab, fso) : "";
    var qualStr = "";
    try {
        qualStr = String(clickData.func.qualifiers + "").toLowerCase();
    } catch (eq) {
        qualStr = "";
    }
    var isCtrl = qualStr.indexOf("ctrl") >= 0;

    var selected = [];
    if (tab) {
        selected = mergeUniquePaths(selected, collectSelectedPaths(tab, fso));
    }
    var onlyListPath = writeOnlyListFile(shell, fso, selected);
    var hasSelection = selected.length > 0;

    if (isCtrl) {
        var execRepeat = quoteArg("python") + " " + quoteArg(converterPy) + " --repeat";
        if (tabFolder && !hasSelection) {
            execRepeat += " --tab-folder " + quoteArg(tabFolder);
        }
        execRepeat = appendOnlyList(execRepeat, onlyListPath);
        if (!onlyListPath) {
            execRepeat = appendOnlyFiles(execRepeat, selected);
        }
        DOpus.Output("convert (repeat last): " + execRepeat);
        var rc = shell.Run(execRepeat, 1, true);
        DOpus.Output("convert (repeat last) exit code: " + rc);
        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf) {}
        return;
    }

    var execGui = quoteArg("pythonw") + " " + quoteArg(converterPy) + " --gui";
    if (tabFolder && !hasSelection) {
        execGui += " --tab-folder " + quoteArg(tabFolder);
    }
    execGui = appendOnlyList(execGui, onlyListPath);
    execGui = appendOnlyFiles(execGui, selected);
    if (onlyListPath) {
        DOpus.Output(
            "convert: " + selected.length + " selected path(s) in input list."
        );
    }
    DOpus.Output("convert (GUI): " + execGui);
    shell.Run(execGui, 0, false);
}
