// Bak-Tsiu launcher for Directory Opus (JScript):
// - No file selected -> start Bak-Tsiu with no arguments
// - One file        -> open that file (normal view)
// - Two or more     -> --split + all paths (split view, all images imported)
//
// Uses tab.selstats.selfiles and tab.selected_files (not Tab.selected / raw
// Item.realpath strings) to avoid 0x8000ffff COM issues in JScript.

var BAKTSIU_EXE = "C:\\Users\\WXP\\Documents\\GitHub\\baktsiu\\build\\src\\Release\\baktsiu.exe";

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    if (!tab) {
        DOpus.dlg.message("No source folder tab.", "Bak-Tsiu");
        return;
    }

    var shell = new ActiveXObject("WScript.Shell");
    var exe = shell.ExpandEnvironmentStrings(BAKTSIU_EXE);

    if (tab.selstats.selfiles == 0) {
        var execBare = '"' + exe + '"';
        DOpus.Output("Bak-Tsiu: " + execBare);
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
        DOpus.Output("Bak-Tsiu (no paths after enumerate): " + execFallback);
        shell.Run(execFallback, 1, false);
        return;
    }

    var exec = '"' + exe + '"';
    if (paths.length > 1) {
        exec += " --split";
    }
    for (var i = 0; i < paths.length; i++) {
        exec += ' "' + paths[i] + '"';
    }

    DOpus.Output("Bak-Tsiu: " + exec);
    shell.Run(exec, 1, false);
}
