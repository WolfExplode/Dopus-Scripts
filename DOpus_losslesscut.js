// LosslessCut launcher for Directory Opus (JScript):
// - No file selected -> start LosslessCut with no arguments
// - One or more files -> open the first in LosslessCut (see DOpus.Output if multiple)
//
// Uses tab.selected_files and Item.realpath (not Tab.selected / raw path strings)
// to avoid 0x8000ffff COM issues in JScript.

var LOSSLESSCUT_EXE = "C:\\Users\\WXP\\Desktop\\Tools\\LosslessCut-win-x64\\LosslessCut.exe";

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    if (!tab) {
        DOpus.dlg.message("No source folder tab.", "LosslessCut");
        return;
    }

    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var exe = shell.ExpandEnvironmentStrings(LOSSLESSCUT_EXE);

    if (!fso.FileExists(exe)) {
        DOpus.dlg.message("LosslessCut not found at:\n" + exe, "LosslessCut");
        return;
    }

    if (tab.selstats.selfiles === 0) {
        var execBare = '"' + exe + '"';
        DOpus.Output("LosslessCut: " + execBare);
        shell.Run(execBare, 1, false);
        return;
    }

    var paths = [];
    var en = new Enumerator(tab.selected_files);
    for (; !en.atEnd(); en.moveNext()) {
        var pathObj = en.item().realpath;
        pathObj.Resolve();
        paths.push(pathObj + "");
    }

    if (paths.length === 0) {
        var execFallback = '"' + exe + '"';
        DOpus.Output("LosslessCut (no paths after enumerate): " + execFallback);
        shell.Run(execFallback, 1, false);
        return;
    }

    var targetPath = paths[0];
    if (paths.length > 1) {
        DOpus.Output("LosslessCut: " + paths.length + " files selected; opening first only: " + targetPath);
    }

    var exec = '"' + exe + '" "' + targetPath + '"';
    DOpus.Output("LosslessCut: " + exec);
    shell.Run(exec, 1, false);
}
