// HandBrakeCLI for Directory Opus

var HANDBRAKE_CLI = "%ProgramFiles%\\HandBrake\\HandBrakeCLI.exe";
var HANDBRAKE_GUI = "%ProgramFiles%\\HandBrake\\HandBrake.exe";
var PRESET_JSON = "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\handbrake.json";
/** 0xC000013A STATUS_CONTROL_C_EXIT — user closed console / Ctrl+C / killed process */
var HANDBRAKE_EXIT_CONTROL_C = -1073741510;

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

/** shell.Popup avoids DOpus.dlg.message 0x8000ffff in some contexts. flags: 16=critical, 48=warn, 64=info */
function popup(shell, text, title, flags) {
    shell.Popup(String(text), 0, String(title), flags == null ? 48 : flags);
}

function pathsEqualIgnoreCase(a, b) {
    return String(a).toLowerCase() === String(b).toLowerCase();
}

function outputPathForInput(fso, inputPath, outExt) {
    var folder = fso.GetParentFolderName(inputPath);
    var base = fso.GetBaseName(inputPath);
    var candidate = folder + "\\" + base + outExt;
    if (pathsEqualIgnoreCase(candidate, inputPath)) {
        return folder + "\\" + base + "_hb" + outExt;
    }
    return candidate;
}

/** Remove partial encode output; force=true clears read-only. Logs if delete fails. */
function deleteIncompleteOutput(fso, outputPath) {
    if (!fso.FileExists(outputPath)) return;
    try {
        fso.DeleteFile(outputPath, true);
        DOpus.Output("HandBrake: removed incomplete output: " + outputPath);
    } catch (e) {
        DOpus.Output(
            "HandBrake: could not remove incomplete output (" +
                outputPath +
                "): " +
                e.message
        );
    }
}

function resolveHandBrakeCli(shell, fso) {
    var std = [
        shell.ExpandEnvironmentStrings(HANDBRAKE_CLI),
        shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\\HandBrake\\HandBrakeCLI.exe")
    ];
    var i;
    for (i = 0; i < std.length; i++) {
        if (fso.FileExists(std[i])) return std[i];
    }
    var roots = [
        shell.ExpandEnvironmentStrings("%ProgramFiles%\\HandBrake"),
        shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\\HandBrake")
    ];
    for (i = 0; i < roots.length; i++) {
        if (!fso.FolderExists(roots[i])) continue;
        var subs = new Enumerator(fso.GetFolder(roots[i]).SubFolders);
        for (; !subs.atEnd(); subs.moveNext()) {
            var exePath = subs.item().Path + "\\HandBrakeCLI.exe";
            if (fso.FileExists(exePath)) return exePath;
        }
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

function outputExtFromHandbrakeFileFormat(fileFormat) {
    var key = String(fileFormat || "").toLowerCase();
    if (key === "av_mkv") return ".mkv";
    if (key === "av_mp4") return ".mp4";
    if (key === "av_webm") return ".webm";
    if (key.indexOf("mkv") >= 0) return ".mkv";
    if (key.indexOf("mp4") >= 0) return ".mp4";
    if (key.indexOf("webm") >= 0) return ".webm";
    return ".mkv";
}

/**
 * Read UTF-8 handbrake.json: pick default preset (Default true) or first in PresetList.
 * Returns { presetName, maxPictureSide, outputExt }.
 */
function activePresetFromHandbrakeJson(fso, presetPath) {
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 2;
    stream.Charset = "UTF-8";
    stream.Open();
    stream.LoadFromFile(presetPath);
    var text = stream.ReadText(-1);
    stream.Close();
    var root = eval("(" + text + ")");
    var list = root.PresetList;
    if (!list || !list.length) {
        throw new Error("PresetList missing or empty.");
    }
    var i;
    var p = null;
    for (i = 0; i < list.length; i++) {
        if (list[i] && list[i].Default === true) {
            p = list[i];
            break;
        }
    }
    if (!p) {
        p = list[0];
    }
    if (!p) {
        throw new Error("No preset object in PresetList.");
    }
    var presetName = p.PresetName;
    if (presetName == null || String(presetName) === "") {
        throw new Error("Active preset has no PresetName.");
    }
    presetName = String(presetName);
    var w = Number(p.PictureWidth) || 0;
    var h = Number(p.PictureHeight) || 0;
    var maxPictureSide = Math.max(w, h);
    if (maxPictureSide < 1) {
        throw new Error("Invalid PictureWidth/Height on preset \"" + presetName + "\".");
    }
    return {
        presetName: presetName,
        maxPictureSide: maxPictureSide,
        outputExt: outputExtFromHandbrakeFileFormat(p.FileFormat)
    };
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var shell = new ActiveXObject("WScript.Shell");
    if (!tab) {
        popup(shell, "No source folder tab.", "HandBrake", 16);
        return;
    }
    var fso = new ActiveXObject("Scripting.FileSystemObject");

    if (tab.selstats.selfiles == 0) {
        var gui = resolveHandBrakeGui(shell, fso);
        if (!gui) {
            popup(shell, "HandBrake.exe not found under Program Files\\HandBrake.", "HandBrake", 16);
            return;
        }
        var execGui = quoteArg(gui);
        DOpus.Output("HandBrake (GUI): " + execGui);
        shell.Run(execGui, 1, false);
        return;
    }

    var cli = resolveHandBrakeCli(shell, fso);
    var presetPath = shell.ExpandEnvironmentStrings(PRESET_JSON);

    if (!cli) {
        popup(shell, "HandBrakeCLI.exe not found under Program Files\\HandBrake.", "HandBrake", 16);
        return;
    }
    if (!fso.FileExists(presetPath)) {
        popup(shell, "Preset JSON not found at:\n" + presetPath, "HandBrake", 16);
        return;
    }

    var active;
    try {
        active = activePresetFromHandbrakeJson(fso, presetPath);
    } catch (e) {
        popup(
            shell,
            "Could not read handbrake.json:\n" + e.message + "\n\n" + presetPath,
            "HandBrake",
            16
        );
        return;
    }
    var presetName = active.presetName;
    var maxPictureSide = active.maxPictureSide;
    var outputExt = active.outputExt;
    DOpus.Output(
        "HandBrake: using preset \"" +
            presetName +
            "\" (max picture side " +
            maxPictureSide +
            ", output " +
            outputExt +
            ")"
    );

    var paths = [];
    var selectedFiles = tab.selected_files;
    var en = new Enumerator(selectedFiles);
    for (; !en.atEnd(); en.moveNext()) {
        paths.push(en.item().realpath + "");
    }

    if (paths.length === 0) {
        var guiFallback = resolveHandBrakeGui(shell, fso);
        if (guiFallback) {
            DOpus.Output("HandBrake (GUI, no paths after enumerate): " + quoteArg(guiFallback));
            shell.Run(quoteArg(guiFallback), 1, false);
        } else {
            popup(shell, "HandBrake.exe not found under Program Files\\HandBrake.", "HandBrake", 16);
        }
        return;
    }

    var presetImport = quoteArg(presetPath);
    var presetFlag = quoteArg(presetName);
    var i;

    for (i = 0; i < paths.length; i++) {
        var inputPath = paths[i];
        var outputPath = outputPathForInput(fso, inputPath, outputExt);
        var cmd =
            quoteArg(cli) +
            " --preset-import-file " +
            presetImport +
            " -Z " +
            presetFlag +
            " --maxWidth " +
            maxPictureSide +
            " --maxHeight " +
            maxPictureSide +
            " --loose-anamorphic" +
            " -i " +
            quoteArg(inputPath) +
            " -o " +
            quoteArg(outputPath);

        DOpus.Output("HandBrakeCLI: " + cmd);
        var rc = shell.Run(cmd, 1, true);
        if (rc === HANDBRAKE_EXIT_CONTROL_C) {
            DOpus.Output(
                "HandBrakeCLI: cancelled or interrupted (exit " +
                    rc +
                    "). Stopped after:\n" +
                    inputPath
            );
            deleteIncompleteOutput(fso, outputPath);
            return;
        }
        if (rc !== 0) {
            popup(
                shell,
                "HandBrakeCLI exited with code " + rc + ".\n\nStopped after:\n" + inputPath,
                "HandBrake",
                16
            );
            return;
        }
    }

    DOpus.Output("HandBrake: finished " + paths.length + " file(s).");
}
