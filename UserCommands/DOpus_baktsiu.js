// Bak-Tsiu launcher for Directory Opus (JScript):
// - No file selected in any visible Lister display -> start Bak-Tsiu with no arguments
// - One file                            -> open that file (normal view)
// - Two or more                         -> --split + all paths (split view, all images imported)
//
// Collects selections from every visible Lister display. Uses tab.selstats.selfiles
// and tab.selected_files (not Tab.selected / raw Item.realpath strings) to
// avoid 0x8000ffff COM issues in JScript.

var BAKTSIU_EXE = "C:\\Users\\WXP\\Documents\\GitHub\\baktsiu\\build\\src\\Release\\baktsiu.exe";

function addSelectedFiles(tab, paths, seen) {
    if (!tab || !tab.visible || tab.selstats.selfiles === 0) {
        return;
    }

    var en = new Enumerator(tab.selected_files);
    for (; !en.atEnd(); en.moveNext()) {
        var pathObj = en.item().realpath;
        pathObj.Resolve();
        var path = pathObj + "";
        // Windows paths are case-insensitive. A file can be encountered twice
        // when the source tab is collected first and again via DOpus.listers.
        var key = path.toLowerCase();
        if (!seen[key]) {
            seen[key] = true;
            paths.push(path);
        }
    }
}

function collectSelectedFiles(sourceTab) {
    var paths = [];
    var seen = {};

    // Keep the invoking tab's selection first so its first item is the
    // initial/main Bak-Tsiu item when more than one file is selected.
    addSelectedFiles(sourceTab, paths, seen);

    var listers = DOpus.listers;
    var listerEnum = new Enumerator(listers);
    for (; !listerEnum.atEnd(); listerEnum.moveNext()) {
        var tabEnum = new Enumerator(listerEnum.item().tabs);
        for (; !tabEnum.atEnd(); tabEnum.moveNext()) {
            addSelectedFiles(tabEnum.item(), paths, seen);
        }
    }
    return paths;
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    if (!tab) {
        DOpus.dlg.message("No source folder tab.", "Bak-Tsiu");
        return;
    }

    var shell = new ActiveXObject("WScript.Shell");
    var exe = shell.ExpandEnvironmentStrings(BAKTSIU_EXE);

    var paths = collectSelectedFiles(tab);

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
