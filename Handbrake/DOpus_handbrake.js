// HandBrake Tool — launches HandbrakeTool.py (Python / Dear PyGui) from this repo.
//
// Click: open GUI, prefilled with the selection. No selection: launch full HandBrake.exe.
// Ctrl+click: re-run the last-used preset/settings on the selection (no dialog).
//
/** Optional full path to HandbrakeTool.py if auto-detect fails. */
var HANDBRAKE_PY = "";
var HANDBRAKE_GUI = "%ProgramFiles%\\HandBrake\\HandBrake.exe";

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

/** shell.Popup avoids DOpus.dlg.message 0x8000ffff in some contexts. flags: 16=critical, 48=warn, 64=info */
function popup(shell, text, title, flags) {
    shell.Popup(String(text), 0, String(title), flags == null ? 48 : flags);
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
    if (!tab || tab.selstats.selitems === 0) {
        return paths;
    }
    var en = new Enumerator(tab.selected);
    for (; !en.atEnd(); en.moveNext()) {
        pushItemPath(paths, en.item(), fso);
    }
    return paths;
}

function writeOnlyListFile(shell, fso, paths) {
    if (!paths || paths.length === 0) {
        return "";
    }
    var name = "HandbrakeTool_only_" + Math.floor(Math.random() * 1000000000) + ".txt";
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

function resolveHandbrakePy(fso) {
    if (HANDBRAKE_PY && fso.FileExists(HANDBRAKE_PY)) {
        return HANDBRAKE_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.BuildPath(
                fso.GetParentFolderName(Script.file),
                "HandbrakeTool.py"
            );
            sibling = fso.GetAbsolutePathName(sibling);
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback =
        "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\Handbrake\\HandbrakeTool.py";
    if (fso.FileExists(fallback)) {
        return fallback;
    }
    return "";
}

function resolveHandBrakeGui(shell, fso) {
    var std = [
        shell.ExpandEnvironmentStrings(HANDBRAKE_GUI),
        shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\\HandBrake\\HandBrake.exe")
    ];
    var i;
    for (i = 0; i < std.length; i++) {
        if (fso.FileExists(std[i])) return std[i];
    }
    return "";
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");

    var selected = tab ? collectSelectedPaths(tab, fso) : [];

    if (selected.length === 0) {
        var gui = resolveHandBrakeGui(shell, fso);
        if (!gui) {
            popup(shell, "HandBrake.exe not found under Program Files\\HandBrake.", "HandBrake", 16);
            return;
        }
        DOpus.Output("HandBrake (GUI): " + quoteArg(gui));
        shell.Run(quoteArg(gui), 1, false);
        return;
    }

    var handbrakePy = resolveHandbrakePy(fso);
    if (!handbrakePy) {
        popup(
            shell,
            "HandbrakeTool.py not found.\n\n" +
                "Copy it under Handbrake\\ next to this repo, set HANDBRAKE_PY in DOpus_handbrake.js, " +
                "or install both under Script AddIns.",
            "HandBrake",
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

    var onlyListPath = writeOnlyListFile(shell, fso, selected);

    if (isCtrl) {
        var execRepeat =
            quoteArg("python") + " " + quoteArg(handbrakePy) + " --repeat";
        execRepeat = appendOnlyList(execRepeat, onlyListPath);
        if (!onlyListPath) {
            execRepeat = appendOnlyFiles(execRepeat, selected);
        }
        DOpus.Output("HandBrake Tool (repeat last): " + execRepeat);
        var rc = shell.Run(execRepeat, 1, true);
        DOpus.Output("HandBrake Tool (repeat last) exit code: " + rc);
        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf) {}
        return;
    }

    var execGui = quoteArg("pythonw") + " " + quoteArg(handbrakePy) + " --gui";
    execGui = appendOnlyList(execGui, onlyListPath);
    execGui = appendOnlyFiles(execGui, selected);
    if (onlyListPath) {
        DOpus.Output(
            "HandBrake Tool: " + selected.length + " selected path(s) in file list."
        );
    }
    DOpus.Output("HandBrake Tool (GUI): " + execGui);
    shell.Run(execGui, 0, false);
}
