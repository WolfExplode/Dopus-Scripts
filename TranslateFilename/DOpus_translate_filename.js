// Translate Filename — launches TranslateFilenameTool.py (Python / Dear PyGui) from this repo.
//
// Click: open GUI with the selected files/folders pre-loaded into Inputs.
//        Translation starts automatically; renaming only happens if you
//        click Apply (or if "Auto-rename after Translate" is turned on in
//        Settings).
// Ctrl+click: translate and rename the selected files/folders immediately,
//             no dialog. Failures are skipped (other items still get
//             renamed) and a summary popup lists any failures.
//
// Settings (incl. DeepSeek API key, history used by Untranslate):
//   %APPDATA%\TranslateFilename\settings.json

/** Optional full path to TranslateFilenameTool.py if auto-detect fails. */
var TRANSLATOR_PY = "";

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

function collectSelectedFiles(tab, fso) {
    var paths = [];
    if (!tab) {
        return paths;
    }

    var en = new Enumerator(tab.selected_files);
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

function writeOnlyListFile(shell, fso, paths) {
    if (!paths || paths.length === 0) {
        return "";
    }
    var name = "TranslateFilename_only_" + Math.floor(Math.random() * 1000000000) + ".txt";
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

function resolveTranslatorPy(fso) {
    if (TRANSLATOR_PY && fso.FileExists(TRANSLATOR_PY)) {
        return TRANSLATOR_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.BuildPath(
                fso.GetParentFolderName(Script.file),
                "TranslateFilenameTool.py"
            );
            sibling = fso.GetAbsolutePathName(sibling);
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback =
        "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\TranslateFilename\\TranslateFilenameTool.py";
    if (fso.FileExists(fallback)) {
        return fallback;
    }
    return "";
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var translatorPy = resolveTranslatorPy(fso);

    if (!translatorPy) {
        shell.Popup(
            "TranslateFilenameTool.py not found.\n\n" +
                "Copy it under TranslateFilename\\ next to this repo, or set TRANSLATOR_PY in DOpus_translate_filename.js.",
            0,
            "Translate Filename",
            16
        );
        return;
    }

    DOpus.Output("translate filename: translatorPy = " + translatorPy);

    var selected = tab ? collectSelectedFiles(tab, fso) : [];
    DOpus.Output("translate filename: selected = " + selected.join(" | "));
    var onlyListPath = writeOnlyListFile(shell, fso, selected);
    DOpus.Output("translate filename: onlyListPath = " + onlyListPath);

    var qualStr = "";
    try {
        qualStr = String(clickData.func.qualifiers + "").toLowerCase();
    } catch (eq) {
        qualStr = "";
    }
    var isCtrl = qualStr.indexOf("ctrl") >= 0;

    if (isCtrl) {
        if (selected.length === 0) {
            shell.Popup("Select one or more files first.", 0, "Translate Filename", 48);
            return;
        }

        var execCli = quoteArg("python") + " " + quoteArg(translatorPy);
        execCli = appendOnlyList(execCli, onlyListPath);
        if (!onlyListPath) {
            execCli = appendOnlyFiles(execCli, selected);
        }
        DOpus.Output("translate filename (ctrl): " + execCli);

        // Run via a temp .bat rather than "cmd /c <string>": once the command
        // line after /c contains more than one pair of quotes, cmd.exe's
        // legacy quote-stripping rule (strip only the very first and very
        // last quote char in the whole line) mangles everything in between —
        // the process never even opens the redirection target. A .bat file
        // is parsed by cmd.exe's normal line parser instead, which has no
        // such quirk.
        var tmp = shell.ExpandEnvironmentStrings("%TEMP%") + "\\TranslateFilename_out_" +
            Math.floor(Math.random() * 1000000000) + ".txt";
        var batPath = shell.ExpandEnvironmentStrings("%TEMP%") + "\\TranslateFilename_run_" +
            Math.floor(Math.random() * 1000000000) + ".bat";
        var batStream = new ActiveXObject("ADODB.Stream");
        batStream.Type = 2;
        batStream.Charset = "utf-8";
        batStream.Open();
        batStream.WriteText("@echo off\r\n" + execCli + " > " + quoteArg(tmp) + " 2>&1\r\n");
        batStream.SaveToFile(batPath, 2);
        batStream.Close();

        DOpus.Output("translate filename (ctrl) bat file: " + batPath);
        DOpus.Output("translate filename (ctrl) tmp file: " + tmp);
        var rc = shell.Run(quoteArg(batPath), 0, true);
        DOpus.Output("translate filename (ctrl) exit code: " + rc);

        var output = "";
        try {
            DOpus.Output("translate filename (ctrl) tmp file exists: " + fso.FileExists(tmp));
            var stream = new ActiveXObject("ADODB.Stream");
            stream.Type = 2;
            stream.Charset = "utf-8";
            stream.Open();
            stream.LoadFromFile(tmp);
            output = stream.ReadText();
            stream.Close();
            fso.DeleteFile(tmp, true);
        } catch (eOut) {
            DOpus.Output("translate filename (ctrl) error reading tmp output: " + eOut.description);
        }
        DOpus.Output("translate filename (ctrl) output: " + output);
        try {
            fso.DeleteFile(batPath, true);
        } catch (eBat) {}

        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf) {}

        if (rc !== 0) {
            var failureLines = [];
            var lines = trimStr(output).split(/\r?\n/);
            var li;
            for (li = 0; li < lines.length; li++) {
                if (lines[li].indexOf(": ERROR - ") >= 0) {
                    failureLines.push(lines[li]);
                }
            }
            var failureMsg = failureLines.length > 0
                ? failureLines.join("\n")
                : (trimStr(output) || "Some files failed to translate.");
            shell.Popup(failureMsg, 0, "Translate Filename — failures", 48);
        }
        return;
    }

    var execGui = quoteArg("pythonw") + " " + quoteArg(translatorPy) + " --gui";
    execGui = appendOnlyList(execGui, onlyListPath);
    execGui = appendOnlyFiles(execGui, selected);
    DOpus.Output("translate filename (GUI): " + execGui);
    shell.Run(execGui, 0, false);
}
