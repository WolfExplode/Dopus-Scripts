// Organize Files — launches OrganizeFiles.py (Python / Dear PyGui) from this repo.
//
// Click: open GUI. Dual pane: source + destination tab paths as hints.
// Files selected in source or destination pane: only those files are processed
// (checkbox on in GUI; Ctrl+preview uses the same list).
// No selection: whole source/target folders as before.
//
// Ctrl+click: run the last GUI action (e.g. tag apply) on the selection — always
// Apply, not Preview. Uses saved tag/strip paths from %APPDATA%\OrganizeFiles\settings.json.
// GUI also saves which panels were expanded (e.g. Filename tags) when you close the window.
//
/** Optional full path to OrganizeFiles.py if auto-detect fails. */
var ORGANIZE_PY = "";

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function tabFolderPath(tab) {
    if (!tab || !tab.path) {
        return "";
    }
    return trimStr(tab.path + "");
}

function pushFilePath(paths, item, fso) {
    var pathObj = item.realpath;
    pathObj.Resolve();
    var p = trimStr(pathObj + "");
    if (p && fso.FileExists(p)) {
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
        if (tab.selstats.checkedfiles === 0) {
            return paths;
        }
        var enChecked = new Enumerator(tab.files);
        for (; !enChecked.atEnd(); enChecked.moveNext()) {
            var checkedItem = enChecked.item();
            if (!checkedItem.checked || checkedItem.is_dir) {
                continue;
            }
            pushFilePath(paths, checkedItem, fso);
        }
        return paths;
    }
    if (tab.selstats.selfiles === 0) {
        return paths;
    }
    var en = new Enumerator(tab.selected_files);
    for (; !en.atEnd(); en.moveNext()) {
        pushFilePath(paths, en.item(), fso);
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
    var name =
        "OrganizeFiles_only_" + Math.floor(Math.random() * 1000000000) + ".txt";
    var file = fso.BuildPath(shell.ExpandEnvironmentStrings("%TEMP%"), name);
    // FSO CreateTextFile(..., false) is ANSI and fails on emoji / CJK in paths.
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

function resolveOrganizePy(shell, fso) {
    if (ORGANIZE_PY && fso.FileExists(ORGANIZE_PY)) {
        return ORGANIZE_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.BuildPath(
                fso.GetParentFolderName(Script.file),
                "OrganizeFiles.py"
            );
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback =
        "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\Organize Files\\OrganizeFiles.py";
    if (fso.FileExists(fallback)) {
        return fallback;
    }
    return "";
}

function resolvePythonExe() {
    return "python";
}

function resolvePythonwExe() {
    return "pythonw";
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var organizePy = resolveOrganizePy(shell, fso);

    if (!organizePy) {
        shell.Popup(
            "OrganizeFiles.py not found.\n\n" +
                "Copy it next to this script, set ORGANIZE_PY in DOpus_OrganizeFiles.js, " +
                "or install both under Script AddIns.",
            0,
            "Organize Files",
            16
        );
        return;
    }

    var srcHint = tab ? tabFolderPath(tab) : "";
    var destTab = clickData.func.desttab;
    var tgtHint = tabFolderPath(destTab);
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
    if (destTab) {
        selected = mergeUniquePaths(
            selected,
            collectSelectedFilePaths(destTab, fso)
        );
    }
    var onlyListPath = writeOnlyListFile(shell, fso, selected);

    if (isCtrl) {
        var py = resolvePythonExe();
        var execRepeat = quoteArg(py) + " " + quoteArg(organizePy) + " --repeat";
        if (srcHint && !onlyListPath) {
            execRepeat += " --source " + quoteArg(srcHint);
        }
        if (tgtHint) {
            execRepeat += " --target " + quoteArg(tgtHint);
        }
        execRepeat += " --apply";
        execRepeat = appendOnlyList(execRepeat, onlyListPath);
        if (!onlyListPath) {
            execRepeat = appendOnlyFiles(execRepeat, selected);
        }
        DOpus.Output("Organize Files (repeat last): " + execRepeat);
        var rc = shell.Run(execRepeat, 1, true);
        DOpus.Output("Organize Files (repeat last) exit code: " + rc);
        return;
    }

    var pyw = resolvePythonwExe();
    var execGui = quoteArg(pyw) + " " + quoteArg(organizePy) + " --gui";
    if (srcHint && !onlyListPath) {
        execGui += " --source " + quoteArg(srcHint);
    }
    if (tgtHint) {
        execGui += " --target " + quoteArg(tgtHint);
    }
    execGui = appendOnlyList(execGui, onlyListPath);
    execGui = appendOnlyFiles(execGui, selected);
    if (onlyListPath) {
        DOpus.Output(
            "Organize Files: " + selected.length + " selected file(s) only."
        );
    }
    DOpus.Output("Organize Files (GUI): " + execGui);
    shell.Run(execGui, 0, false);
}
