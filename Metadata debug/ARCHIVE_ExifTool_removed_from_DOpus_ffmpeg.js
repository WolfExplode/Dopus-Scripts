/**
 * ARCHIVE — removed from DOpus_ffmpeg.js (ExifTool metadata restore after ffmpeg).
 * Re-integrate by: paste functions below after thumbMimeForImageExt(), restore EXIFTOOL_EXE + calls.
 *
 * Call sites that were removed (search main script to re-add):
 * - thumbEmbedCoverCore: after MoveFile tmp→mediaPath → thumbRestoreMetadataExifTool(shell, fso, bakPath, mediaPath);
 * - runSplitOrCombineCover (strip path): after MoveFile stripTmp→mediaPath → thumbRestoreMetadataExifTool(shell, fso, stripBak, mediaPath);
 * - runAudioToMono: after successful replace → restoreMetadataExifTool(shell, fso, bakPath, vidPath);
 * - runSplitAvCopy combine branch: after muxDone → restoreMetadataExifTool(shell, fso, bakPath, vidPath);
 * - runSplitAvCopy split branch: after MoveFile vidTmp→vidPath → restoreMetadataExifTool(shell, fso, bakPath, vidPath);
 *   and if audioOk → restoreMetadataExifTool(shell, fso, bakPath, audOut);
 * - OnClick convert success: restoreMetadataExifTool(shell, fso, item.realpath + "", outPath);
 *
 * Standalone test script (same folder): exif_copy_mp3_SourceToTarget.ps1
 */

var EXIFTOOL_EXE = "C:\\Users\\WXP\\Desktop\\Tools\\ExifTool-13.55\\exiftool-13.55_64\\exiftool.exe";

function writeExifToolArgFile(fso, filePath, args) {
    var stream = new ActiveXObject("ADODB.Stream");
    stream.Type = 2;
    stream.Charset = "UTF-8";
    stream.Open();
    stream.WriteText(args.join("\r\n"));
    stream.SaveToFile(filePath, 2);
    stream.Close();
}

function exifToolRunFromArgFile(shell, fso, label, args) {
    var tmpDir = fso.GetSpecialFolder(2);
    var argPath = fso.BuildPath(tmpDir, "opus_exif_" + Math.floor(Math.random() * 1000000000) + ".args");
    try {
        writeExifToolArgFile(fso, argPath, args);
        var cmd = '"' + EXIFTOOL_EXE + '" -@ "' + argPath + '"';
        DOpus.Output("ExifTool (" + label + "): " + cmd + "  (" + args.length + " args, UTF-8)");
        return shell.Run(cmd, 0, true);
    } catch (ex) {
        DOpus.Output("ExifTool: " + ex.message);
        return -1;
    } finally {
        try {
            if (fso.FileExists(argPath)) {
                fso.DeleteFile(argPath);
            }
        } catch (eDel) { /* ignore */ }
    }
}

function restoreMetadataExifToolMp3Supplement(shell, fso, tagSourcePath, targetPath) {
    function runSupp(label, extraArgs) {
        var args = ["-m", "-charset", "filename=UTF8", "-TagsFromFile", tagSourcePath].concat(extraArgs).concat(["-overwrite_original", targetPath]);
        exifToolRunFromArgFile(shell, fso, label, args);
    }
    runSupp("restore MP3 supplement TXXX/WXXX/PRIV/WOAR*", [
        "-TXXX:All",
        "-WXXX:All",
        "-PRIV:All",
        "-WOAR",
        "-WOAS",
        "-WORS",
        "-WCOM"
    ]);
    runSupp("restore MP3 supplement Comment+Comment-xxx", ["-Comment", "-Comment-xxx"]);
    runSupp("restore MP3 supplement APE (DOpus)", ["-APE:All"]);
}

function restoreMetadataExifTool(shell, fso, tagSourcePath, targetPath) {
    if (!EXIFTOOL_EXE || EXIFTOOL_EXE.length < 4) {
        return;
    }
    if (!fso.FileExists(EXIFTOOL_EXE)) {
        DOpus.Output("ExifTool: skip (not found): " + EXIFTOOL_EXE);
        return;
    }
    if (!fso.FileExists(tagSourcePath) || !fso.FileExists(targetPath)) {
        DOpus.Output("ExifTool: skip (source or target missing)");
        return;
    }
    var ext = fileExtLower(targetPath);

    function runPass(label, extraArgs) {
        var args = ["-m", "-charset", "filename=UTF8", "-TagsFromFile", tagSourcePath].concat(extraArgs).concat(["-overwrite_original", targetPath]);
        return exifToolRunFromArgFile(shell, fso, label, args);
    }

    function mp3SupplementIfNeeded() {
        if (ext == ".mp3") {
            restoreMetadataExifToolMp3Supplement(shell, fso, tagSourcePath, targetPath);
        }
    }

    var code;
    if (ext == ".mp3") {
        code = runPass("restore 1 MP3 backup all:all", ["-all:all", "-unsafe"]);
        if (code !== 0) {
            code = runPass("restore 1 MP3 ID3:All fallback", ["-ID3:All"]);
        }
        if (code === 0) {
            mp3SupplementIfNeeded();
            return;
        }
    }
    code = runPass("restore all+unsafe", ["-all:all", "-unsafe"]);
    if (code === 0) {
        mp3SupplementIfNeeded();
        return;
    }
    if (ext == ".mp3") {
        code = runPass("restore MP3 ID3:All retry", ["-ID3:All"]);
        if (code === 0) {
            mp3SupplementIfNeeded();
            return;
        }
    }

    if (ext == ".mp4" || ext == ".m4v" || ext == ".mov" || ext == ".m4a" || ext == ".aac") {
        code = runPass("restore QuickTime/XMP", ["-XMP:All", "-ItemList:All", "-Keys:All", "-UserData:All", "-QuickTime:All"]);
        if (code === 0) {
            return;
        }
    }

    if (ext == ".mkv" || ext == ".mka") {
        code = runPass("restore Matroska", ["-Matroska:All"]);
        if (code === 0) {
            return;
        }
    }

    if (ext == ".flac") {
        code = runPass("restore FLAC", ["-FLAC:All"]);
        if (code === 0) {
            return;
        }
    }
    if (ext == ".ogg" || ext == ".opus") {
        code = runPass("restore Vorbis", ["-Vorbis:All"]);
        if (code === 0) {
            return;
        }
    }

    code = runPass("restore last XMP+EXIF+IPTC", ["-XMP:All", "-EXIF:All", "-IPTC:All"]);
    if (code === 0) {
        mp3SupplementIfNeeded();
    } else {
        DOpus.Output("ExifTool: still exit " + code + " after fallbacks — run exiftool manually with -v3 on these paths to see errors.");
    }
}

function thumbRestoreMetadataExifTool(shell, fso, tagSourcePath, targetPath) {
    restoreMetadataExifTool(shell, fso, tagSourcePath, targetPath);
}
