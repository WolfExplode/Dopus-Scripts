// Krokiet (Czkawka GUI) launcher for Directory Opus (JScript):
// - Source tab folder -> Krokiet included path
// - Destination tab folder (dual-display) -> second included path when different
// - No real folder paths (e.g. lib:// only) -> start Krokiet with no CLI paths
//
// Krokiet CLI: krokiet_gui [--maximize] [FOLDERS...]  (-e / -r for exclude / referenced)

var KROKET_EXE_CANDIDATES = [
    "C:\\Users\\WXP\\Documents\\GitHub\\czkawka\\target\\fast_release\\krokiet.exe",
    "C:\\Users\\WXP\\Documents\\GitHub\\czkawka\\target\\debug\\deps\\krokiet.exe"
];

function fileModifiedMs(fso, path) {
    var raw = fso.GetFile(path).DateLastModified;
    return (new Date(raw)).getTime();
}

function resolveKrokietExe(fso, shell) {
    var best = "";
    var bestTime = 0;
    var i;
    for (i = 0; i < KROKET_EXE_CANDIDATES.length; i++) {
        var path = shell.ExpandEnvironmentStrings(KROKET_EXE_CANDIDATES[i]);
        if (!fso.FileExists(path)) {
            continue;
        }
        var mtime = fileModifiedMs(fso, path);
        if (!bestTime || mtime > bestTime) {
            bestTime = mtime;
            best = path;
        }
    }
    return best;
}

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function tabFolderPath(tab, fso) {
    if (!tab || !tab.path) {
        return "";
    }
    var pathObj = tab.path;
    try {
        pathObj.Resolve();
    } catch (e) {}
    var p = trimStr(pathObj + "");
    if (p && fso.FolderExists(p)) {
        return p;
    }
    return "";
}

function pushUniqueFolder(paths, seen, folder) {
    if (!folder) {
        return;
    }
    var key = folder.toLowerCase();
    if (seen[key]) {
        return;
    }
    seen[key] = true;
    paths.push(folder);
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    if (!tab) {
        DOpus.dlg.message("No source folder tab.", "Krokiet");
        return;
    }

    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var exe = resolveKrokietExe(fso, shell);

    if (!exe) {
        DOpus.dlg.message(
            "Krokiet not found. Checked:\n" + KROKET_EXE_CANDIDATES.join("\n"),
            "Krokiet"
        );
        return;
    }

    var paths = [];
    var seen = {};

    pushUniqueFolder(paths, seen, tabFolderPath(tab, fso));

    var destTab = clickData.func.desttab;
    if (destTab) {
        pushUniqueFolder(paths, seen, tabFolderPath(destTab, fso));
    }

    var exec = '"' + exe + '" --maximize';
    var i;
    for (i = 0; i < paths.length; i++) {
        exec += ' "' + paths[i] + '"';
    }

    var exeDir = fso.GetParentFolderName(exe);
    shell.CurrentDirectory = exeDir;

    DOpus.Output("Krokiet (" + exe + "): " + exec);
    shell.Run(exec, 1, false);
}
