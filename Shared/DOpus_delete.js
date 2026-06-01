// Shared delete helpers for Directory Opus JScript tools.
// Load from OnInit: eval(DOpus_delete.js) — assigns global DOpusDeleteLib.

function dopusDelete_isEphemeralTempPath(pathStr, shell) {
    var lower = String(pathStr).toLowerCase();
    var temp = String(shell.ExpandEnvironmentStrings("%TEMP%")).toLowerCase();
    if (temp && lower.indexOf(temp) === 0) {
        return true;
    }
    if (lower.indexOf("dopus_hb_") >= 0 && lower.lastIndexOf(".txt") === lower.length - 4) {
        return true;
    }
    if (lower.indexOf("dopus_where_") >= 0 && lower.lastIndexOf(".txt") === lower.length - 4) {
        return true;
    }
    if (lower.indexOf("dopus_frame_") >= 0 && lower.lastIndexOf(".txt") === lower.length - 4) {
        return true;
    }
    if (lower.indexOf("gdl-") >= 0 && (lower.lastIndexOf(".txt") === lower.length - 4 || lower.lastIndexOf(".ps1") === lower.length - 4)) {
        return true;
    }
    return false;
}

DOpusDeleteLib = {
    isEphemeralTempPath: dopusDelete_isEphemeralTempPath,

    permanentDeleteFile: function (fso, filePath) {
        if (!fso.FileExists(filePath)) {
            return true;
        }
        try {
            fso.DeleteFile(filePath, true);
            return true;
        } catch (e) {
            return false;
        }
    },

    recycleDeleteFile: function (shell, fso, filePath) {
        if (!fso.FileExists(filePath)) {
            return true;
        }
        var psPath = String(filePath).replace(/'/g, "''");
        var cmd =
            'powershell -NoProfile -Command "Add-Type -AssemblyName Microsoft.VisualBasic; ' +
            "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile('" +
            psPath +
            "', 'OnlyErrorDialogs', 'SendToRecycleBin')\"";
        try {
            shell.Run(cmd, 0, true);
        } catch (e) {}
        if (!fso.FileExists(filePath)) {
            return true;
        }
        return this.permanentDeleteFile(fso, filePath);
    },

    deleteFile: function (shell, fso, filePath) {
        if (!fso.FileExists(filePath)) {
            return true;
        }
        if (dopusDelete_isEphemeralTempPath(filePath, shell)) {
            return this.permanentDeleteFile(fso, filePath);
        }
        return this.recycleDeleteFile(shell, fso, filePath);
    }
};
