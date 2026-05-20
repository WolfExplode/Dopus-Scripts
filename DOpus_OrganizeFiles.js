// Organize Files — launches OrganizeFiles.pyw (Python / Dear PyGui) from this repo.
//
// Click: open GUI. Dual pane: source + destination tab paths as hints.
// Files selected in source or destination pane: only those files are processed
// (checkbox on in GUI; Ctrl+preview uses the same list).
// No selection: whole source/target folders as before.
//
// Ctrl+click: preview checkmark scan in a console.
//
// Install: copy DOpus_OrganizeFiles.js and OrganizeFiles.pyw to the same folder
// (e.g. Script AddIns), or set ORGANIZE_PYW below. Python must be on PATH.

/** Optional full path to OrganizeFiles.pyw if auto-detect fails. */
var ORGANIZE_PYW = "";

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

function collectSelectedFilePaths(tab, fso) {
    var paths = [];
    if (!tab || tab.selstats.selfiles === 0) {
        return paths;
    }
    var en = new Enumerator(tab.selected_files);
    for (; !en.atEnd(); en.moveNext()) {
        var pathObj = en.item().realpath;
        pathObj.Resolve();
        var p = trimStr(pathObj + "");
        if (p && fso.FileExists(p)) {
            paths.push(p);
        }
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
    var stream = fso.CreateTextFile(file, true, false);
    var i;
    for (i = 0; i < paths.length; i++) {
        stream.WriteLine(paths[i]);
    }
    stream.Close();
    return file;
}

function appendOnlyList(exec, onlyListPath) {
    if (onlyListPath) {
        return exec + " --only-list " + quoteArg(onlyListPath);
    }
    return exec;
}

function resolveOrganizePyw(shell, fso) {
    if (ORGANIZE_PYW && fso.FileExists(ORGANIZE_PYW)) {
        return ORGANIZE_PYW;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.BuildPath(
                fso.GetParentFolderName(Script.file),
                "OrganizeFiles.pyw"
            );
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback =
        "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\OrganizeFiles.pyw";
    if (fso.FileExists(fallback)) {
        return fallback;
    }
    return "";
}

function resolvePythonExe(fso, projectDir) {
    var venv1 = projectDir + "\\.venv\\Scripts\\python.exe";
    var venv2 = projectDir + "\\venv\\Scripts\\python.exe";
    if (fso.FileExists(venv1)) {
        return venv1;
    }
    if (fso.FileExists(venv2)) {
        return venv2;
    }
    return "python";
}

function resolvePythonwExe(fso, projectDir) {
    var py = resolvePythonExe(fso, projectDir);
    if (py === "python") {
        return "pythonw";
    }
    var dir = fso.GetParentFolderName(py);
    var pyw = dir + "\\pythonw.exe";
    if (fso.FileExists(pyw)) {
        return pyw;
    }
    return py;
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var pywPath = resolveOrganizePyw(shell, fso);

    if (!pywPath) {
        shell.Popup(
            "OrganizeFiles.pyw not found.\n\n" +
                "Copy it next to this script, set ORGANIZE_PYW in DOpus_OrganizeFiles.js, " +
                "or install both under Script AddIns.",
            0,
            "Organize Files",
            16
        );
        return;
    }

    var projectDir = fso.GetParentFolderName(pywPath);
    var srcHint = tab ? tabFolderPath(tab) : "";
    var destTab = clickData.func.desttab;
    var tgtHint = tabFolderPath(destTab);
    var qual = String(clickData.func.qualifiers || "").toLowerCase();
    var isCtrl = qual.indexOf("ctrl") >= 0;

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
        var py = resolvePythonExe(fso, projectDir);
        var execPreview =
            quoteArg(py) + " " + quoteArg(pywPath) + " --action mark --preview";
        if (srcHint) {
            execPreview += " --source " + quoteArg(srcHint);
        }
        if (tgtHint) {
            execPreview += " --target " + quoteArg(tgtHint);
        }
        execPreview = appendOnlyList(execPreview, onlyListPath);
        DOpus.Output("Organize Files (preview): " + execPreview);
        var rc = shell.Run(execPreview, 1, true);
        DOpus.Output("Organize Files (preview) exit code: " + rc);
        return;
    }

    var pyw = resolvePythonwExe(fso, projectDir);
    var execGui = quoteArg(pyw) + " " + quoteArg(pywPath) + " --gui";
    if (srcHint) {
        execGui += " --source " + quoteArg(srcHint);
    }
    if (tgtHint) {
        execGui += " --target " + quoteArg(tgtHint);
    }
    execGui = appendOnlyList(execGui, onlyListPath);
    if (onlyListPath) {
        DOpus.Output(
            "Organize Files: " + selected.length + " selected file(s) only."
        );
    }
    DOpus.Output("Organize Files (GUI): " + execGui);
    shell.Run(execGui, 1, false);
}
