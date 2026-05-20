// BPM Analysis (Heartbeat BPM Analyzer) for Directory Opus (JScript):
// - No file selected -> start GUI (main.py) with no arguments
// - One or more paths selected -> headless batch_cli.py with those paths (all items;
//   the CLI skips anything that is not a file). Console stays open until the batch finishes.
//
// Uses tab.selstats.selfiles and tab.selected_files + Item.realpath
// to avoid 0x8000ffff COM issues in JScript.
//
// Python: set BPM_PYTHON_EXE if needed. Otherwise uses .venv under the project
// folder if present, else "python" on PATH. (Avoid "py -3" here — it can pick a
// different install than your shell and exit immediately with missing modules.)

var BPM_MAIN_PY = "C:\\Users\\WXP\\Documents\\GitHub\\bpm_analysis\\main.py";
/** Optional full path to python.exe that has requirements installed. */
var BPM_PYTHON_EXE = "";

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

/** Quoted exe path, or plain "python" on PATH (not quoted). */
function resolvePythonPrefix(fso, projectDir) {
    if (BPM_PYTHON_EXE && fso.FileExists(BPM_PYTHON_EXE)) {
        return quoteArg(BPM_PYTHON_EXE);
    }
    var venv1 = projectDir + "\\.venv\\Scripts\\python.exe";
    var venv2 = projectDir + "\\venv\\Scripts\\python.exe";
    if (fso.FileExists(venv1)) return quoteArg(venv1);
    if (fso.FileExists(venv2)) return quoteArg(venv2);
    return "python";
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    if (!tab) {
        DOpus.dlg.message("No source folder tab.", "BPM Analysis");
        return;
    }

    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var mainPy = shell.ExpandEnvironmentStrings(BPM_MAIN_PY);
    var projectDir = fso.GetParentFolderName(mainPy);
    var batchCli = projectDir + "\\batch_cli.py";
    var py = resolvePythonPrefix(fso, projectDir);

    if (!fso.FileExists(mainPy)) {
        DOpus.dlg.message("main.py not found at:\n" + mainPy, "BPM Analysis");
        return;
    }

    if (tab.selstats.selfiles === 0) {
        var execBare = py + " " + quoteArg(mainPy);
        DOpus.Output("BPM Analysis (GUI): " + execBare);
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
        var execFallback = py + " " + quoteArg(mainPy);
        DOpus.Output("BPM Analysis (GUI, no paths after enumerate): " + execFallback);
        shell.Run(execFallback, 1, false);
        return;
    }

    if (!fso.FileExists(batchCli)) {
        DOpus.dlg.message("batch_cli.py not found at:\n" + batchCli, "BPM Analysis");
        return;
    }

    var exec = py + " " + quoteArg(batchCli);
    var i;
    for (i = 0; i < paths.length; i++) {
        exec += " " + quoteArg(paths[i]);
    }
    DOpus.Output("BPM Analysis (batch_cli): " + exec);
    var rc = shell.Run(exec, 1, true);
    DOpus.Output("BPM Analysis (batch_cli) exit code: " + rc);
}
