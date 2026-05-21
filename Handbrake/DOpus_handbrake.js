// HandBrakeCLI for Directory Opus

var HANDBRAKE_CLI = "%ProgramFiles%\\HandBrake\\HandBrakeCLI.exe";
var HANDBRAKE_GUI = "%ProgramFiles%\\HandBrake\\HandBrake.exe";
var PRESET_DIR = "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\Handbrake";
var DEFAULT_MAX_PICTURE_SIDE = 1920;
var SETTINGS_FILE = null;
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

function getSettingsPath(shell) {
    if (!SETTINGS_FILE) {
        SETTINGS_FILE =
            shell.ExpandEnvironmentStrings("%APPDATA%") + "\\DOpus_handbrake_settings.ini";
    }
    return SETTINGS_FILE;
}

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function loadHandbrakeSettings(shell, fso) {
    var out = {
        presetFile: "",
        maxSide: String(DEFAULT_MAX_PICTURE_SIDE),
        videoQuality: "",
        videoFramerate: ""
    };
    try {
        var path = getSettingsPath(shell);
        if (fso.FileExists(path)) {
            var stream = fso.OpenTextFile(path, 1, false);
            var content = stream.ReadAll();
            stream.Close();
            var lines = content.split("\n");
            var i;
            for (i = 0; i < lines.length; i++) {
                var line = lines[i].replace(/\r$/, "");
                var eq = line.indexOf("=");
                if (eq > 0) {
                    var key = line.substring(0, eq);
                    var val = line.substring(eq + 1);
                    if (key === "presetFile") out.presetFile = val;
                    else if (key === "maxSide") out.maxSide = val;
                    else if (key === "videoQuality") out.videoQuality = val;
                    else if (key === "videoFramerate") out.videoFramerate = val;
                }
            }
        }
    } catch (e) { /* use defaults */ }
    return out;
}

function saveHandbrakeSettings(shell, fso, presetFile, maxSide, videoQuality, videoFramerate) {
    try {
        var path = getSettingsPath(shell);
        var stream = fso.OpenTextFile(path, 2, true);
        stream.WriteLine("presetFile=" + presetFile);
        stream.WriteLine("maxSide=" + maxSide);
        stream.WriteLine("videoQuality=" + (videoQuality == null ? "" : videoQuality));
        stream.WriteLine("videoFramerate=" + (videoFramerate == null ? "" : videoFramerate));
        stream.Close();
    } catch (e) { /* ignore */ }
}

function parseMaxPictureSide(shell, raw) {
    var n = parseInt(trimStr(raw), 10);
    if (isNaN(n) || n < 1) {
        popup(shell, "Enter a positive number for max picture size (pixels).", "HandBrake", 16);
        return 0;
    }
    return n;
}

/** Blank = no override (null). Otherwise number for HandBrakeCLI -q. Invalid input returns false. */
function parseVideoQualityOverride(shell, raw) {
    var s = trimStr(raw);
    if (s === "") {
        return null;
    }
    var n = parseFloat(s);
    if (isNaN(n)) {
        popup(shell, "Video quality must be a number, or leave blank to use the preset.", "HandBrake", 16);
        return false;
    }
    return n;
}

/** Blank = no override (null). Otherwise fps for HandBrakeCLI -r. Invalid input returns false. */
function parseVideoFramerateOverride(shell, raw) {
    var s = trimStr(raw);
    if (s === "") {
        return null;
    }
    var n = parseFloat(s);
    if (isNaN(n) || n <= 0) {
        popup(shell, "Frame rate must be a positive number, or leave blank to use the preset.", "HandBrake", 16);
        return false;
    }
    return n;
}

/** Sorted list of { label, path } for *.json in PRESET_DIR. */
function listPresetJsonFiles(fso) {
    var rows = [];
    if (!fso.FolderExists(PRESET_DIR)) {
        return rows;
    }
    var folder = fso.GetFolder(PRESET_DIR);
    var files = new Enumerator(folder.Files);
    for (; !files.atEnd(); files.moveNext()) {
        var file = files.item();
        var name = String(file.Name);
        if (name.toLowerCase().indexOf(".json") === name.length - 5) {
            rows.push({ label: name, path: file.Path });
        }
    }
    rows.sort(function (a, b) {
        return String(a.label).toLowerCase() < String(b.label).toLowerCase() ? -1 : 1;
    });
    return rows;
}

function comboPresetPath(raw, rows) {
    var n = parseInt(String(raw), 10);
    if (!isNaN(n) && n >= 0 && n < rows.length) {
        return rows[n].path;
    }
    return rows.length ? rows[0].path : "";
}

function populatePresetCombo(dlg, rows, selectFileName) {
    var combo = dlg.control("preset_combo");
    combo.RemoveItem(-1);
    var i;
    for (i = 0; i < rows.length; i++) {
        combo.AddItem(rows[i].label, rows[i].path);
    }
    var sel = 0;
    if (selectFileName) {
        for (i = 0; i < rows.length; i++) {
            if (String(rows[i].label).toLowerCase() === String(selectFileName).toLowerCase()) {
                sel = i;
                break;
            }
        }
    }
    if (rows.length > 0) {
        combo.SelectItem(sel);
    }
}

/**
 * Show encode dialog. Returns encode options or null if cancelled / invalid input.
 * videoQuality / videoFramerate are null when blank (use preset values).
 */
function pickHandbrakeEncodeOptions(clickData, shell, fso) {
    var rows = listPresetJsonFiles(fso);
    if (rows.length === 0) {
        popup(
            shell,
            "No preset JSON files found in:\n" + PRESET_DIR + "\n\nAdd HandBrake-exported *.json files there.",
            "HandBrake",
            16
        );
        return null;
    }

    var saved = loadHandbrakeSettings(shell, fso);
    var dlg = DOpus.dlg;
    dlg.window = clickData.func.sourcetab;
    dlg.template = "DOpus_handbrake_Dlg";
    dlg.detach = true;
    dlg.Create();
    populatePresetCombo(dlg, rows, saved.presetFile);
    dlg.control("maxside_edit").value = saved.maxSide || String(DEFAULT_MAX_PICTURE_SIDE);
    dlg.control("quality_edit").value = saved.videoQuality || "";
    dlg.control("framerate_edit").value = saved.videoFramerate || "";
    dlg.Show();

    var dialogResult = 0;
    while (true) {
        var msg = dlg.GetMsg();
        if (!msg.result) {
            dialogResult = dlg.result;
            break;
        }
    }

    if (dialogResult != "1") {
        return null;
    }

    var presetPath = comboPresetPath(dlg.control("preset_combo").value, rows);
    if (!presetPath) {
        return null;
    }
    var maxPictureSide = parseMaxPictureSide(shell, dlg.control("maxside_edit").value);
    if (!maxPictureSide) {
        return null;
    }
    var qualityRaw = dlg.control("quality_edit").value;
    var videoQuality = parseVideoQualityOverride(shell, qualityRaw);
    if (videoQuality === false) {
        return null;
    }
    var framerateRaw = dlg.control("framerate_edit").value;
    var videoFramerate = parseVideoFramerateOverride(shell, framerateRaw);
    if (videoFramerate === false) {
        return null;
    }
    saveHandbrakeSettings(
        shell,
        fso,
        fso.GetFileName(presetPath),
        String(maxPictureSide),
        trimStr(qualityRaw),
        trimStr(framerateRaw)
    );
    return {
        presetPath: presetPath,
        maxPictureSide: maxPictureSide,
        videoQuality: videoQuality,
        videoFramerate: videoFramerate
    };
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
 * Read UTF-8 preset JSON: pick default preset (Default true) or first in PresetList.
 * Returns { presetName, outputExt }.
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
    return {
        presetName: presetName,
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

    if (!cli) {
        popup(shell, "HandBrakeCLI.exe not found under Program Files\\HandBrake.", "HandBrake", 16);
        return;
    }

    var options = pickHandbrakeEncodeOptions(clickData, shell, fso);
    if (!options) {
        return;
    }
    var presetPath = options.presetPath;
    var maxPictureSide = options.maxPictureSide;
    var videoQuality = options.videoQuality;
    var videoFramerate = options.videoFramerate;
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
            "Could not read preset JSON:\n" + e.message + "\n\n" + presetPath,
            "HandBrake",
            16
        );
        return;
    }
    var presetName = active.presetName;
    var outputExt = active.outputExt;
    var logMsg =
        "HandBrake: using preset \"" +
        presetName +
        "\" (max picture side " +
        maxPictureSide +
        ", output " +
        outputExt;
    if (videoQuality != null) {
        logMsg += ", quality -q " + videoQuality;
    }
    if (videoFramerate != null) {
        logMsg += ", framerate -r " + videoFramerate;
    }
    logMsg += ")";
    DOpus.Output(logMsg);

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
            (videoQuality != null ? " -q " + videoQuality : "") +
            (videoFramerate != null ? " -r " + videoFramerate : "") +
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
