// Video/Audio Converter with XML Dialog
var SETTINGS_FILE = null;  // Set on first use via %APPDATA%

function getSettingsPath(shell) {
    if (!SETTINGS_FILE) {
        SETTINGS_FILE = shell.ExpandEnvironmentStrings("%APPDATA%") + "\\DOpus_ffmpeg_settings.ini";
    }
    return SETTINGS_FILE;
}

function loadLastSettings(shell, fso) {
    var path = getSettingsPath(shell);
    var out = { mode: 0, formatName: "", quality: "23", lastAction: "convert" };
    try {
        if (fso.FileExists(path)) {
            var stream = fso.OpenTextFile(path, 1, false);
            var content = stream.ReadAll();
            stream.Close();
            var lines = content.split("\n");
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].replace(/\r$/, "");
                var eq = line.indexOf("=");
                if (eq > 0) {
                    var key = line.substring(0, eq);
                    var val = line.substring(eq + 1);
                    if (key == "mode") out.mode = parseInt(val, 10) || 0;
                    else if (key == "formatName") out.formatName = val;
                    else if (key == "quality") out.quality = val;
                    else if (key == "lastAction") out.lastAction = val;
                }
            }
        }
    } catch (e) { /* use defaults */ }
    if (!out.lastAction) out.lastAction = "convert";
    return out;
}

function saveLastSettings(shell, fso, mode, formatName, quality, lastAction) {
    try {
        if (lastAction === undefined || lastAction === null || lastAction === "") {
            lastAction = "convert";
        }
        var path = getSettingsPath(shell);
        var stream = fso.OpenTextFile(path, 2, true);  // ForWriting, Create
        stream.WriteLine("mode=" + mode);
        stream.WriteLine("formatName=" + formatName);
        stream.WriteLine("quality=" + (quality || "23"));
        stream.WriteLine("lastAction=" + lastAction);
        stream.Close();
    } catch (e) { /* ignore */ }
}

/** Remember which command was used (convert / cover / mono / splitav / splitch) while keeping convert presets. */
function saveLastActionOnly(shell, fso, action) {
    var last = loadLastSettings(shell, fso);
    saveLastSettings(shell, fso, last.mode, last.formatName, last.quality, action);
}

function fileExtLower(name) {
    var p = (name + "").lastIndexOf(".");
    if (p < 0) return "";
    return (name + "").substring(p).toLowerCase();
}

var THUMB_IMAGE_EXT = {
    ".jpg": 1, ".jpeg": 1, ".jfif": 1, ".pjpeg": 1, ".pjp": 1,
    ".png": 1, ".apng": 1,
    ".webp": 1, ".bmp": 1, ".dib": 1, ".gif": 1,
    ".tif": 1, ".tiff": 1,
    ".heic": 1, ".heif": 1, ".avif": 1, ".jxl": 1
};
var THUMB_VIDEO_EXT = {
    ".mp4": 1, ".m4v": 1, ".mov": 1, ".qt": 1,
    ".mkv": 1, ".webm": 1, ".avi": 1, ".wmv": 1, ".asf": 1,
    ".mpg": 1, ".mpeg": 1, ".mpe": 1, ".m1v": 1, ".m2v": 1, ".mpv": 1,
    ".vob": 1,
    ".ts": 1, ".mts": 1, ".m2t": 1, ".m2ts": 1,
    ".3gp": 1, ".3g2": 1,
    ".flv": 1, ".f4v": 1,
    ".ogv": 1, ".ogm": 1,
    ".dv": 1,
    ".mxf": 1
};
var THUMB_AUDIO_EXT = {
    ".mp3": 1, ".mp2": 1, ".mpa": 1,
    ".m4a": 1, ".m4b": 1, ".m4p": 1, ".aac": 1, ".adts": 1,
    ".flac": 1, ".wav": 1, ".aiff": 1, ".aif": 1, ".aifc": 1, ".caf": 1,
    ".ogg": 1, ".oga": 1, ".opus": 1,
    ".mka": 1,
    ".wma": 1,
    ".ac3": 1, ".eac3": 1, ".dts": 1,
    ".amr": 1, ".awb": 1,
    ".au": 1, ".snd": 1,
    ".ape": 1, ".tta": 1, ".wv": 1,
    ".weba": 1
};

function isThumbImageName(name) {
    return THUMB_IMAGE_EXT[fileExtLower(name)] == 1;
}
function isThumbVideoName(name) {
    return THUMB_VIDEO_EXT[fileExtLower(name)] == 1;
}
function isThumbAudioName(name) {
    return THUMB_AUDIO_EXT[fileExtLower(name)] == 1;
}

function mimeTypeForImageExt(ext) {
    if (ext == ".jpg" || ext == ".jpeg" || ext == ".jfif" || ext == ".pjpeg" || ext == ".pjp") return "image/jpeg";
    if (ext == ".png" || ext == ".apng") return "image/png";
    if (ext == ".webp") return "image/webp";
    if (ext == ".gif") return "image/gif";
    if (ext == ".bmp" || ext == ".dib") return "image/bmp";
    if (ext == ".tif" || ext == ".tiff") return "image/tiff";
    if (ext == ".heic" || ext == ".heif") return "image/heif";
    if (ext == ".avif") return "image/avif";
    if (ext == ".jxl") return "image/jxl";
    return "image/jpeg";
}

// Log only (no modal dialogs): shell.Popup / DOpus.dlg.message are not used.
function thumbInfo(shell, text, title) {
    DOpus.Output("[" + title + "] " + text);
}
function thumbErr(shell, text, title) {
    DOpus.Output("[" + title + " ERROR] " + text);
}

/** Next unused path folder\\stem.jpg, stem_1.jpg, … */
function uniqueJpgPathNextToMedia(fso, folder, stem) {
    var outPath = folder + "\\" + stem + ".jpg";
    var counter = 1;
    while (fso.FileExists(outPath)) {
        outPath = folder + "\\" + stem + "_" + counter + ".jpg";
        counter++;
    }
    return outPath;
}

/**
 * Try embedded picture streams: preferSecondVideoFirst true = typical video+poster (try 0:v:1 then 0:v:0);
 * false = typical audio/cover-only (try 0:v:0 then 0:v:1). Writes JPEG to outPath.
 */
function extractCoverViaVideoStreams(shell, fso, mediaPath, outPath, preferSecondVideoFirst) {
    var order = preferSecondVideoFirst ? ["0:v:1", "0:v:0"] : ["0:v:0", "0:v:1"];
    var i;
    for (i = 0; i < order.length; i++) {
        if (fso.FileExists(outPath)) {
            try {
                fso.DeleteFile(outPath);
            } catch (eDel) { /* ignore */ }
        }
        var exec = 'ffmpeg.exe -y -i "' + mediaPath + '" -map ' + order[i] + ' -frames:v 1 -q:v 2 "' + outPath + '"';
        DOpus.Output("Extract thumbnail: " + exec);
        try {
            var exitCode = shell.Run(exec, 0, true);
            if (exitCode === 0 && fso.FileExists(outPath)) {
                return true;
            }
        } catch (ex) {
            DOpus.Output("Extract thumbnail stream attempt error: " + ex.message);
        }
    }
    return false;
}

/**
 * Matroska cover is often an attachment; dump then decode to JPEG so .jpg matches file contents (PNG/WebP attachments).
 * Returns true if outPath exists.
 */
function extractCoverMatroskaAttachment(shell, fso, mediaPath, folder, stem, outPath) {
    var rawPath = folder + "\\" + stem + ".__opus_cover_raw";
    if (fso.FileExists(rawPath)) {
        try {
            fso.DeleteFile(rawPath);
        } catch (e0) { /* ignore */ }
    }
    var dumpExec = 'ffmpeg.exe -y -dump_attachment:t:0 "' + rawPath + '" -i "' + mediaPath + '"';
    DOpus.Output("Extract thumbnail (MKV/MKA attachment): " + dumpExec);
    var dumpExit;
    try {
        dumpExit = shell.Run(dumpExec, 0, true);
    } catch (exD) {
        DOpus.Output("Extract thumbnail: dump attachment error: " + exD.message);
        return false;
    }
    if (!fso.FileExists(rawPath)) {
        DOpus.Output("Extract thumbnail: no Matroska attachment t:0 (exit " + dumpExit + "); will try video streams if any.");
        return false;
    }
    if (fso.FileExists(outPath)) {
        try {
            fso.DeleteFile(outPath);
        } catch (e1) { /* ignore */ }
    }
    var convExec = 'ffmpeg.exe -y -i "' + rawPath + '" -frames:v 1 -q:v 2 "' + outPath + '"';
    DOpus.Output("Extract thumbnail (normalize to JPEG): " + convExec);
    var convExit;
    try {
        convExit = shell.Run(convExec, 0, true);
    } catch (exC) {
        convExit = -1;
        DOpus.Output("Extract thumbnail: JPEG normalize error: " + exC.message);
    }
    try {
        fso.DeleteFile(rawPath);
    } catch (eR) { /* ignore */ }
    if (convExit !== 0 || !fso.FileExists(outPath)) {
        DOpus.Output("Extract thumbnail: could not decode attachment to JPEG (exit " + convExit + "); will try video streams if any.");
        return false;
    }
    return true;
}

/** Try to write embedded cover from media to outPath (.jpg). */
function tryExtractCoverToPath(shell, fso, mediaPath, mediaName, folder, stem, ext, outPath) {
    if (ext == ".mkv" || ext == ".mka") {
        if (extractCoverMatroskaAttachment(shell, fso, mediaPath, folder, stem, outPath)) {
            return true;
        }
        return extractCoverViaVideoStreams(shell, fso, mediaPath, outPath, false);
    }
    if (isThumbVideoName(mediaName)) {
        return extractCoverViaVideoStreams(shell, fso, mediaPath, outPath, true);
    }
    return extractCoverViaVideoStreams(shell, fso, mediaPath, outPath, false);
}

/** True if this filename is a supported video or audio container for embed cover/poster. */
function isThumbMediaName(name) {
    return isThumbVideoName(name) || isThumbAudioName(name);
}

/**
 * ffmpeg command to embed imgPath as cover/poster into mediaPath -> tmpPath.
 * isVideoExt: from THUMB_VIDEO_EXT (mp4, mov, …), not audio-only extensions.
 */
function ffmpegSetThumbnailExec(mediaPath, imgPath, imgPathForMime, ext, tmpPath, isVideoExt) {
    if (ext == ".mkv" || ext == ".mka") {
        var mime = mimeTypeForImageExt(imgPathForMime);
        return 'ffmpeg.exe -y -i "' + mediaPath + '" -map_metadata 0 -map_chapters 0 -map 0 -map -0:t -c copy -attach "' + imgPath + '" -metadata:s:t mimetype=' + mime + ' "' + tmpPath + '"';
    }
    if (isVideoExt) {
        return 'ffmpeg.exe -y -i "' + mediaPath + '" -i "' + imgPath + '" -map_metadata 0 -map_chapters 0 -map 0:v:0 -map 0:a? -map 0:s? -map 0:d? -map 0:t? -map 1 -c copy -c:v:1 mjpeg -disposition:v:1 attached_pic "' + tmpPath + '"';
    }
    if (ext == ".m4a" || ext == ".m4b" || ext == ".m4p" || ext == ".aac") {
        return 'ffmpeg.exe -y -i "' + mediaPath + '" -i "' + imgPath + '" -map_metadata 0 -map_chapters 0 -map 0:a? -map 1:0 -c copy -c:v:0 mjpeg -disposition:v:0 attached_pic "' + tmpPath + '"';
    }
    if (ext == ".mp3") {
        return 'ffmpeg.exe -y -i "' + mediaPath + '" -i "' + imgPath + '" -map_metadata 0 -map 0:a? -map 1:0 -c copy "' + tmpPath + '"';
    }
    return 'ffmpeg.exe -y -i "' + mediaPath + '" -i "' + imgPath + '" -map_metadata 0 -map 0:a? -map 1:0 -c copy "' + tmpPath + '"';
}

/**
 * Remux without extra video/cover streams or attachments (best-effort). Matroska tries video+attachments drop then audio-only.
 * Returns true if stripTmp was created.
 */
function tryStripCoverToTmp(shell, fso, mediaPath, ext, stripTmp, isVideoExt) {
    function runAttempt(cmd) {
        if (fso.FileExists(stripTmp)) {
            try {
                fso.DeleteFile(stripTmp);
            } catch (e0) { /* ignore */ }
        }
        DOpus.Output("Strip cover: " + cmd);
        try {
            var ex = shell.Run(cmd, 0, true);
            return ex === 0 && fso.FileExists(stripTmp);
        } catch (exRun) {
            return false;
        }
    }
    var cmdVideo = 'ffmpeg.exe -y -i "' + mediaPath + '" -map_metadata 0 -map_chapters 0 -map 0:v:0 -map 0:a? -map 0:s? -map 0:d? -map 0:t? -c copy "' + stripTmp + '"';
    var cmdMkvAudioOnly = 'ffmpeg.exe -y -i "' + mediaPath + '" -map_metadata 0 -map_chapters 0 -map 0:a? -map 0:s? -c copy "' + stripTmp + '"';
    var cmdM4a = 'ffmpeg.exe -y -i "' + mediaPath + '" -map_metadata 0 -map_chapters 0 -map 0:a? -c copy "' + stripTmp + '"';
    var cmdMp3 = 'ffmpeg.exe -y -i "' + mediaPath + '" -map_metadata 0 -map 0:a? -c copy "' + stripTmp + '"';
    var cmdAudioGeneric = 'ffmpeg.exe -y -i "' + mediaPath + '" -map_metadata 0 -map 0:a? -c copy "' + stripTmp + '"';

    if (ext == ".mkv" || ext == ".mka") {
        if (runAttempt(cmdVideo)) {
            return true;
        }
        return runAttempt(cmdMkvAudioOnly);
    }
    if (isVideoExt) {
        return runAttempt(cmdVideo);
    }
    if (ext == ".m4a" || ext == ".m4b" || ext == ".m4p" || ext == ".aac") {
        return runAttempt(cmdM4a);
    }
    if (ext == ".mp3") {
        return runAttempt(cmdMp3);
    }
    return runAttempt(cmdAudioGeneric);
}

/**
 * Embed imgPath into mediaItem in place (tmp/bak replace).
 * @return {{ ok: boolean, err: string }}
 */
function thumbEmbedCoverCore(shell, fso, mediaItem, imgPath, imgPathForMime) {
    var mediaName = mediaItem.name + "";
    var mediaPath = mediaItem.realpath + "";
    var folder = mediaItem.path + "";
    var stem = mediaItem.name_stem + "";
    var ext = fileExtLower(mediaName);
    var isVideoExt = isThumbVideoName(mediaName);
    var tmpPath = folder + "\\" + stem + ".__opus_thumb_tmp" + ext;
    var bakPath = folder + "\\" + stem + ".__opus_thumb_orig" + ext;

    if (fso.FileExists(tmpPath)) {
        try {
            fso.DeleteFile(tmpPath);
        } catch (e0) { /* ignore */ }
    }
    if (fso.FileExists(bakPath)) {
        try {
            fso.DeleteFile(bakPath);
        } catch (e1) { /* ignore */ }
    }

    var exec = ffmpegSetThumbnailExec(mediaPath, imgPath, imgPathForMime, ext, tmpPath, isVideoExt);
    DOpus.Output("Embed cover: " + exec);

    try {
        var exitCode = shell.Run(exec, 0, true);
        if (exitCode != 0) {
            return { ok: false, err: "ffmpeg failed (exit code " + exitCode + "). See DOpus Script Output." };
        }
        if (!fso.FileExists(tmpPath)) {
            return { ok: false, err: "ffmpeg finished but the output file was not created." };
        }

        try {
            fso.MoveFile(mediaPath, bakPath);
        } catch (eRen) {
            return { ok: false, err: "Could not rename the original file (it may be open in another program).\n\nNew file left at:\n" + tmpPath };
        }
        try {
            fso.MoveFile(tmpPath, mediaPath);
        } catch (eMv) {
            try {
                fso.MoveFile(bakPath, mediaPath);
            } catch (eRest) { /* ignore */ }
            return { ok: false, err: "Could not replace the media file; the original was restored." };
        }
        try {
            fso.DeleteFile(bakPath);
        } catch (eDel) { /* leave backup if locked */ }

        return { ok: true, err: "" };
    } catch (ex) {
        return { ok: false, err: "Error: " + ex.message };
    }
}

/**
 * Like split/combine AV: 1 image + 1 media → embed cover and delete the image file; 1 media → extract .jpg and remux media without embedded cover.
 */
function runSplitOrCombineCover(clickData, fso, shell) {
    var logTitle = "Split/combine cover";
    var sel = clickData.func.sourcetab.selected_files;
    var imgItems = [];
    var mediaItems = [];
    var badNames = [];
    var en = new Enumerator(sel);
    for (; !en.atEnd(); en.moveNext()) {
        var it = en.item();
        var n = it.name + "";
        if (isThumbImageName(n)) {
            imgItems.push(it);
        } else if (isThumbMediaName(n)) {
            mediaItems.push(it);
        } else {
            badNames.push(n);
        }
    }

    if (badNames.length > 0) {
        thumbErr(shell, "Unsupported file type(s): " + badNames.join(", "), logTitle);
        return;
    }

    if (imgItems.length > 0) {
        if (sel.count != 2 || imgItems.length != 1 || mediaItems.length != 1) {
            thumbErr(shell, "To combine cover: select exactly one image and one video or audio file (nothing else).", logTitle);
            return;
        }
        var imgItem = imgItems[0];
        var mediaItem = mediaItems[0];
        var imgPath = imgItem.realpath + "";
        var imgPathForMime = fileExtLower(imgItem.name + "");
        var res = thumbEmbedCoverCore(shell, fso, mediaItem, imgPath, imgPathForMime);
        if (!res.ok) {
            thumbErr(shell, res.err, logTitle);
            return;
        }
        try {
            fso.DeleteFile(imgPath);
        } catch (exDel) {
            DOpus.Output("[" + logTitle + "] Embedded OK but could not delete image file (in use?): " + imgPath + " — " + exDel.message);
        }
        thumbInfo(shell, "Cover embedded in media (same as combine AV). Image file removed if possible:\n" + (mediaItem.realpath + ""), logTitle);
        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf) { /* ignore */ }
        return;
    }

    if (mediaItems.length != 1 || sel.count != 1) {
        thumbErr(shell, "To split cover: select exactly one video or audio file (nothing else).", logTitle);
        return;
    }

    var mediaItem = mediaItems[0];
    var mediaName = mediaItem.name + "";
    var mediaPath = mediaItem.realpath + "";
    var folder = mediaItem.path + "";
    var stem = mediaItem.name_stem + "";
    var ext = fileExtLower(mediaName);
    var isVideoExt = isThumbVideoName(mediaName);
    var outPath = uniqueJpgPathNextToMedia(fso, folder, stem);
    var stripTmp = folder + "\\" + stem + ".__opus_strip_tmp" + ext;
    var stripBak = folder + "\\" + stem + ".__opus_strip_orig" + ext;

    try {
        if (!tryExtractCoverToPath(shell, fso, mediaPath, mediaName, folder, stem, ext, outPath)) {
            thumbErr(shell, "No embedded cover to split, or extract failed. See DOpus Script Output.\n\n" + mediaName, logTitle);
            return;
        }

        if (fso.FileExists(stripTmp)) {
            try {
                fso.DeleteFile(stripTmp);
            } catch (eSt0) { /* ignore */ }
        }
        if (fso.FileExists(stripBak)) {
            try {
                fso.DeleteFile(stripBak);
            } catch (eSt1) { /* ignore */ }
        }

        if (!tryStripCoverToTmp(shell, fso, mediaPath, ext, stripTmp, isVideoExt)) {
            thumbInfo(shell, "Cover saved to:\n" + outPath + "\n\nCould not remove embedded cover from the media file (see Script Output). Original media unchanged.", logTitle);
            try {
                clickData.func.command.RunCommand("Go REFRESH");
            } catch (eRf0) { /* ignore */ }
            return;
        }

        try {
            fso.MoveFile(mediaPath, stripBak);
        } catch (eRen) {
            thumbErr(shell, "Cover saved to:\n" + outPath + "\n\nCould not rename media to strip cover (file in use?). Deleted temp strip output.", logTitle);
            try {
                fso.DeleteFile(stripTmp);
            } catch (eCl) { /* ignore */ }
            return;
        }
        try {
            fso.MoveFile(stripTmp, mediaPath);
        } catch (eMv) {
            try {
                fso.MoveFile(stripBak, mediaPath);
            } catch (eRest) { /* ignore */ }
            try {
                fso.DeleteFile(stripTmp);
            } catch (eCl2) { /* ignore */ }
            thumbErr(shell, "Could not replace media after strip; restored original. Cover file:\n" + outPath, logTitle);
            return;
        }
        try {
            fso.DeleteFile(stripBak);
        } catch (eDelB) { /* ignore */ }

        thumbInfo(shell, "Split cover: media without embedded cover:\n" + mediaPath + "\nImage:\n" + outPath, logTitle);
        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf1) { /* ignore */ }
    } catch (ex) {
        thumbErr(shell, "Error: " + ex.message, logTitle);
    }
}

/** ffmpeg audio codec + options for mono remux (video `-c copy`); must match container. */
function monoAudioEncodeArgsForExt(ext) {
    if (ext == ".webm") {
        return "libopus -ac 1 -b:a 128k";
    }
    if (ext == ".avi") {
        return "libmp3lame -ac 1 -b:a 192k";
    }
    if (ext == ".wmv") {
        return "wmav2 -ac 1 -b:a 128k";
    }
    if (ext == ".ogv" || ext == ".ogm") {
        return "libvorbis -ac 1 -b:a 192k";
    }
    if (ext == ".flv" || ext == ".f4v") {
        return "aac -ac 1 -b:a 192k";
    }
    if (ext == ".mpg" || ext == ".mpeg" || ext == ".mpe" || ext == ".m1v" || ext == ".vob") {
        return "mp2 -ac 1 -b:a 192k";
    }
    return "aac -ac 1 -b:a 192k";
}

/** Re-encode all audio to mono, copy video and other streams; same extension/path in place. */
function runAudioToMono(clickData, fso, shell) {
    var sel = clickData.func.sourcetab.selected_files;
    if (sel.count < 1) {
        thumbErr(shell, "Select one or more video files.", "Audio to mono");
        return;
    }
    var list = [];
    var en = new Enumerator(sel);
    for (; !en.atEnd(); en.moveNext()) {
        var it = en.item();
        if (!isThumbVideoName(it.name + "")) {
            thumbErr(shell, "Not a supported video file:\n\n" + it.name, "Audio to mono");
            return;
        }
        list.push(it);
    }

    var ok = 0;
    var fail = 0;

    for (var i = 0; i < list.length; i++) {
        var vidItem = list[i];
        var vidPath = vidItem.realpath + "";
        var folder = vidItem.path + "";
        var ext = fileExtLower(vidItem.name + "");
        var stem = vidItem.name_stem + "";
        var tmpPath = folder + "\\" + stem + ".__opus_mono_tmp" + ext;
        var bakPath = folder + "\\" + stem + ".__opus_mono_orig" + ext;
        var aEnc = monoAudioEncodeArgsForExt(ext);

        if (fso.FileExists(tmpPath)) {
            try {
                fso.DeleteFile(tmpPath);
            } catch (eT0) { /* ignore */ }
        }
        if (fso.FileExists(bakPath)) {
            try {
                fso.DeleteFile(bakPath);
            } catch (eT1) { /* ignore */ }
        }

        var exec = 'ffmpeg.exe -y -i "' + vidPath + '" -map_metadata 0 -map_chapters 0 -map 0 -c copy -c:a ' + aEnc + ' "' + tmpPath + '"';
        DOpus.Output("Audio to mono: " + exec);

        try {
            var exitCode = shell.Run(exec, 0, true);
            if (exitCode != 0) {
                DOpus.Output("Audio to mono failed (exit " + exitCode + "): " + vidItem.name);
                fail++;
                continue;
            }
            if (!fso.FileExists(tmpPath)) {
                DOpus.Output("Audio to mono: output missing after ffmpeg: " + vidItem.name);
                fail++;
                continue;
            }

            try {
                fso.MoveFile(vidPath, bakPath);
            } catch (eRen) {
                DOpus.Output("Audio to mono: could not rename original (in use?): " + vidItem.name + " — left temp: " + tmpPath);
                fail++;
                continue;
            }
            try {
                fso.MoveFile(tmpPath, vidPath);
            } catch (eMv) {
                try {
                    fso.MoveFile(bakPath, vidPath);
                } catch (eRest) { /* ignore */ }
                DOpus.Output("Audio to mono: could not replace file, restored original: " + vidItem.name);
                fail++;
                continue;
            }
            try {
                fso.DeleteFile(bakPath);
            } catch (eDel) { /* leave backup if locked */ }
            ok++;
        } catch (ex) {
            DOpus.Output("Audio to mono error on " + vidItem.name + ": " + ex.message);
            fail++;
        }
    }

    if (fail > 0 && ok === 0) {
        thumbErr(shell, "All " + fail + " file(s) failed. See DOpus Script Output.", "Audio to mono");
    } else if (fail > 0) {
        thumbInfo(shell, "Finished with errors. OK: " + ok + ", Failed: " + fail + ". Details in Script Output.", "Audio to mono");
    } else {
        thumbInfo(shell, "Audio to mono finished. Files updated: " + ok, "Audio to mono");
    }
    try {
        clickData.func.command.RunCommand("Go REFRESH");
    } catch (eRf) { /* ignore */ }
}

/** Channel count for first audio stream (0:a:0), or -1 if missing / ffprobe failed. */
function probeAudioChannelCount(shell, fso, mediaPath) {
    var tmp = shell.ExpandEnvironmentStrings("%TEMP%") + "\\DOpus_ffmpeg_chcnt.txt";
    if (fso.FileExists(tmp)) {
        try {
            fso.DeleteFile(tmp);
        } catch (e0) { /* ignore */ }
    }
    var cmd = 'cmd /c ffprobe.exe -v error -select_streams a:0 -show_entries stream=channels -of csv=p=0 "';
    cmd += mediaPath + '" 1> "' + tmp + '"';
    var code;
    try {
        code = shell.Run(cmd, 0, true);
    } catch (ex) {
        return -1;
    }
    if (code != 0 || !fso.FileExists(tmp)) {
        try {
            if (fso.FileExists(tmp)) {
                fso.DeleteFile(tmp);
            }
        } catch (e1) { /* ignore */ }
        return -1;
    }
    var line = "";
    try {
        var ts = fso.OpenTextFile(tmp, 1);
        line = ts.ReadAll().replace(/^\s+|\s+$/g, "").replace(/\r|\n/g, "");
        ts.Close();
    } catch (eR) {
        line = "";
    }
    try {
        fso.DeleteFile(tmp);
    } catch (e2) { /* ignore */ }
    var n = parseInt(line, 10);
    if (isNaN(n) || n < 1) {
        return -1;
    }
    return n;
}

/**
 * First audio stream only: one mono 16-bit WAV per channel (stem.ch01.wav …).
 * Filter script is written under %TEMP% (short path) so long/Unicode media paths do not exceed MAX_PATH.
 */
function runExtractAllAudioChannels(clickData, fso, shell) {
    var logTitle = "Extract audio channels";
    var sel = clickData.func.sourcetab.selected_files;
    if (sel.count < 1) {
        thumbErr(shell, "Select one or more video or audio files.", logTitle);
        return;
    }
    var list = [];
    var en = new Enumerator(sel);
    for (; !en.atEnd(); en.moveNext()) {
        var it = en.item();
        var n = it.name + "";
        if (!isThumbVideoName(n) && !isThumbAudioName(n)) {
            thumbErr(shell, "Not a supported video or audio file:\n\n" + n, logTitle);
            return;
        }
        list.push(it);
    }

    var ok = 0;
    var fail = 0;

    for (var i = 0; i < list.length; i++) {
        var item = list[i];
        var mediaPath = item.realpath + "";
        var folder = item.path + "";
        var stem = item.name_stem + "";

        var ch = probeAudioChannelCount(shell, fso, mediaPath);
        if (ch < 1) {
            DOpus.Output(logTitle + ": no audio stream or could not read channel count: " + item.name);
            fail++;
            continue;
        }

        var filtPath = shell.ExpandEnvironmentStrings("%TEMP%") + "\\DOpus_ffmpeg_chfilt_" + i + "_" + (new Date()).getTime() + ".txt";
        if (fso.FileExists(filtPath)) {
            try {
                fso.DeleteFile(filtPath);
            } catch (eF0) { /* ignore */ }
        }

        var parts = [];
        var c;
        for (c = 0; c < ch; c++) {
            parts.push("[0:a:0]pan=mono|c0=c" + c + "[ch" + c + "]");
        }
        try {
            var fw = fso.OpenTextFile(filtPath, 2, true);
            fw.WriteLine(parts.join(";"));
            fw.Close();
        } catch (eW) {
            DOpus.Output(logTitle + ": could not write filter file: " + item.name + " — " + eW.message);
            fail++;
            continue;
        }

        var exec = 'ffmpeg.exe -y -i "' + mediaPath + '" -filter_complex_script "' + filtPath + '" -c:a pcm_s16le';
        for (c = 0; c < ch; c++) {
            var idx = c + 1;
            var suffix = idx < 10 ? "0" + idx : String(idx);
            var outPath = folder + "\\" + stem + ".ch" + suffix + ".wav";
            var counter = 1;
            while (fso.FileExists(outPath)) {
                outPath = folder + "\\" + stem + ".ch" + suffix + "_" + counter + ".wav";
                counter++;
            }
            exec += ' -map "[ch' + c + ']" "' + outPath + '"';
        }

        DOpus.Output(logTitle + " (" + ch + " ch): " + exec);
        try {
            var exitCode = shell.Run(exec, 0, true);
            if (exitCode != 0) {
                DOpus.Output(logTitle + " failed (exit " + exitCode + "): " + item.name);
                fail++;
            } else {
                ok++;
            }
        } catch (ex) {
            DOpus.Output(logTitle + " error on " + item.name + ": " + ex.message);
            fail++;
        }
        try {
            if (fso.FileExists(filtPath)) {
                fso.DeleteFile(filtPath);
            }
        } catch (eF1) { /* ignore */ }
    }

    if (fail > 0 && ok === 0) {
        thumbErr(shell, "All " + fail + " file(s) failed. See DOpus Script Output.", logTitle);
    } else if (fail > 0) {
        thumbInfo(shell, "Finished with errors. OK: " + ok + ", Failed: " + fail + ". Details in Script Output.", logTitle);
    } else {
        thumbInfo(shell, logTitle + " finished. Files processed: " + ok, logTitle);
    }
    try {
        clickData.func.command.RunCommand("Go REFRESH");
    } catch (eRf) { /* ignore */ }
}

/** Split: demux with -c copy; original path becomes video-only, audio → stem.audio.mka. Combine: one video + one audio → remux to video’s path (-c copy), delete audio. */
function runSplitAvCopy(clickData, fso, shell) {
    var logTitle = "Split/combine AV";
    var sel = clickData.func.sourcetab.selected_files;
    if (sel.count < 1) {
        thumbErr(shell, "Select one or more video files, or one video plus one audio file.", logTitle);
        return;
    }

    var vidItems = [];
    var audItems = [];
    var badNames = [];
    var en = new Enumerator(sel);
    for (; !en.atEnd(); en.moveNext()) {
        var it = en.item();
        var n = it.name + "";
        if (isThumbVideoName(n)) {
            vidItems.push(it);
        } else if (isThumbAudioName(n)) {
            audItems.push(it);
        } else {
            badNames.push(n);
        }
    }

    if (badNames.length > 0) {
        thumbErr(shell, "Unsupported file type(s): " + badNames.join(", "), logTitle);
        return;
    }

    if (audItems.length > 0) {
        if (sel.count != 2 || vidItems.length != 1 || audItems.length != 1) {
            thumbErr(shell, "To combine: select exactly one video and one audio (nothing else). To split: select only video file(s).", logTitle);
            return;
        }
        var vidItem = vidItems[0];
        var audItem = audItems[0];
        var vidPath = vidItem.realpath + "";
        var audPath = audItem.realpath + "";
        var folder = vidItem.path + "";
        var ext = fileExtLower(vidItem.name + "");
        var stem = vidItem.name_stem + "";
        var muxTmp = folder + "\\" + stem + ".__opus_mux_tmp" + ext;
        var bakPath = folder + "\\" + stem + ".__opus_mux_orig" + ext;
        if (fso.FileExists(muxTmp)) {
            try {
                fso.DeleteFile(muxTmp);
            } catch (eMt0) { /* ignore */ }
        }
        if (fso.FileExists(bakPath)) {
            try {
                fso.DeleteFile(bakPath);
            } catch (eBk0) { /* ignore */ }
        }
        var execMux = 'ffmpeg.exe -y -i "' + vidPath + '" -i "' + audPath + '" -map_metadata 0 -map_chapters 0 -map 0:v:0 -map 1:a:0 -c copy -shortest "' + muxTmp + '"';
        DOpus.Output(logTitle + " (combine): " + execMux);
        try {
            var muxExit = shell.Run(execMux, 0, true);
            if (muxExit != 0 || !fso.FileExists(muxTmp)) {
                thumbErr(shell, "Combine failed (ffmpeg exit " + muxExit + "). Container/codec pair may be incompatible with stream copy. See Script Output.", logTitle);
            } else {
                try {
                    if (fso.FileExists(audPath)) {
                        fso.DeleteFile(audPath);
                    }
                } catch (exDelA) {
                    DOpus.Output(logTitle + " (combine): could not delete audio source: " + audPath + " — " + exDelA.message);
                }
                var muxDone = false;
                try {
                    fso.MoveFile(vidPath, bakPath);
                } catch (exRen) {
                    DOpus.Output(logTitle + " (combine): could not rename video aside: " + vidPath + " — " + exRen.message);
                    try {
                        fso.DeleteFile(muxTmp);
                    } catch (eCl) { /* ignore */ }
                    thumbErr(shell, "Combine output was discarded (could not replace video). See Script Output.", logTitle);
                }
                if (fso.FileExists(muxTmp) && fso.FileExists(bakPath)) {
                    try {
                        fso.MoveFile(muxTmp, vidPath);
                        muxDone = true;
                    } catch (exMv) {
                        try {
                            fso.MoveFile(bakPath, vidPath);
                        } catch (eRest) { /* ignore */ }
                        try {
                            fso.DeleteFile(muxTmp);
                        } catch (eCl2) { /* ignore */ }
                        thumbErr(shell, "Combine: could not move mux to final path; restored video-only file.", logTitle);
                    }
                }
                if (muxDone) {
                    try {
                        fso.DeleteFile(bakPath);
                    } catch (eDelB) { /* ignore */ }
                    var audGone = !fso.FileExists(audPath);
                    var sumMsg = "Remuxed to: " + vidPath + " (same name as video input).";
                    if (!audGone) {
                        sumMsg += " Audio file could not be deleted — see Script Output.";
                    }
                    thumbInfo(shell, sumMsg, logTitle);
                }
            }
        } catch (exM) {
            thumbErr(shell, "Combine error: " + exM.message, logTitle);
        }
        try {
            clickData.func.command.RunCommand("Go REFRESH");
        } catch (eRf0) { /* ignore */ }
        return;
    }

    if (vidItems.length < 1) {
        thumbErr(shell, "No video file in selection.", logTitle);
        return;
    }

    var list = vidItems;

    var ok = 0;
    var partial = 0;
    var fail = 0;

    for (var i = 0; i < list.length; i++) {
        var vidItem = list[i];
        var vidPath = vidItem.realpath + "";
        var folder = vidItem.path + "";
        var ext = fileExtLower(vidItem.name + "");
        var stem = vidItem.name_stem + "";

        var vidTmp = folder + "\\" + stem + ".__opus_split_v_tmp" + ext;
        var bakPath = folder + "\\" + stem + ".__opus_split_orig" + ext;
        if (fso.FileExists(vidTmp)) {
            try {
                fso.DeleteFile(vidTmp);
            } catch (eTv0) { /* ignore */ }
        }
        if (fso.FileExists(bakPath)) {
            try {
                fso.DeleteFile(bakPath);
            } catch (eBk0) { /* ignore */ }
        }

        var audOut = folder + "\\" + stem + ".audio.mka";
        var ac = 1;
        while (fso.FileExists(audOut)) {
            audOut = folder + "\\" + stem + ".audio_" + ac + ".mka";
            ac++;
        }

        var execV = 'ffmpeg.exe -y -i "' + vidPath + '" -map_metadata 0 -map_chapters 0 -map 0:v:0 -c copy -an "' + vidTmp + '"';
        var execA = 'ffmpeg.exe -y -i "' + vidPath + '" -map_metadata 0 -map_chapters 0 -map 0:a:0 -c copy -vn "' + audOut + '"';

        DOpus.Output(logTitle + " (split video): " + execV);
        try {
            var ev = shell.Run(execV, 0, true);
            if (ev != 0 || !fso.FileExists(vidTmp)) {
                DOpus.Output(logTitle + ": video demux failed (exit " + ev + "): " + vidItem.name);
                fail++;
                continue;
            }
        } catch (exV) {
            DOpus.Output(logTitle + " video error on " + vidItem.name + ": " + exV.message);
            fail++;
            continue;
        }

        var audioOk = false;
        DOpus.Output(logTitle + " (split audio): " + execA);
        try {
            var ea = shell.Run(execA, 0, true);
            if (ea != 0 || !fso.FileExists(audOut)) {
                DOpus.Output(logTitle + ": audio demux failed or no audio stream (exit " + ea + "): " + vidItem.name);
            } else {
                audioOk = true;
            }
        } catch (exA) {
            DOpus.Output(logTitle + " audio error on " + vidItem.name + ": " + exA.message);
        }

        try {
            fso.MoveFile(vidPath, bakPath);
        } catch (eRen) {
            DOpus.Output(logTitle + ": could not rename original (in use?): " + vidItem.name + " — left temp: " + vidTmp);
            try {
                fso.DeleteFile(vidTmp);
            } catch (eDelT) { /* ignore */ }
            fail++;
            continue;
        }
        try {
            fso.MoveFile(vidTmp, vidPath);
        } catch (eMv) {
            try {
                fso.MoveFile(bakPath, vidPath);
            } catch (eRest) { /* ignore */ }
            DOpus.Output(logTitle + ": could not replace with video-only; restored original: " + vidItem.name);
            fail++;
            continue;
        }
        try {
            fso.DeleteFile(bakPath);
        } catch (eDelB) { /* leave backup if locked */ }

        if (audioOk) {
            ok++;
        } else {
            partial++;
        }
    }

    if (fail > 0 && ok === 0 && partial === 0) {
        thumbErr(shell, "All " + fail + " file(s) failed (video demux or replace). See DOpus Script Output.", logTitle);
    } else {
        var msg = "Split finished (original → video-only + .audio.mka). Full: " + ok + ", Video-only file (no separate audio): " + partial;
        if (fail > 0) {
            msg += ", Failed: " + fail;
        }
        msg += ". Details in Script Output.";
        thumbInfo(shell, msg, logTitle);
    }
    try {
        clickData.func.command.RunCommand("Go REFRESH");
    } catch (eRf) { /* ignore */ }
}

/** Video re-encode args for in-place rotate/flip (must match container). */
function videoEncodeForTransform(ext) {
    var e = (ext + "").toLowerCase();
    if (e == ".webm") {
        return "libvpx-vp9 -crf 30 -b:v 0";
    }
    return "libx264 -crf 18 -preset fast -pix_fmt yuv420p";
}

/**
 * Rotate or flip video in place (re-encode first video stream, copy audio). vfFilter e.g. transpose=1, hflip.
 */
function runVideoTransform(clickData, fso, shell, vfFilter, logTitle) {
    var sel = clickData.func.sourcetab.selected_files;
    if (sel.count < 1) {
        thumbErr(shell, "Select one or more video files.", logTitle);
        return;
    }
    var list = [];
    var en = new Enumerator(sel);
    for (; !en.atEnd(); en.moveNext()) {
        var it = en.item();
        if (!isThumbVideoName(it.name + "")) {
            thumbErr(shell, "Not a supported video file:\n\n" + it.name, logTitle);
            return;
        }
        list.push(it);
    }

    var ok = 0;
    var fail = 0;

    for (var i = 0; i < list.length; i++) {
        var vidItem = list[i];
        var vidPath = vidItem.realpath + "";
        var folder = vidItem.path + "";
        var ext = fileExtLower(vidItem.name + "");
        var stem = vidItem.name_stem + "";
        var tmpPath = folder + "\\" + stem + ".__opus_xform_tmp" + ext;
        var bakPath = folder + "\\" + stem + ".__opus_xform_orig" + ext;
        var vEnc = videoEncodeForTransform(ext);

        if (fso.FileExists(tmpPath)) {
            try {
                fso.DeleteFile(tmpPath);
            } catch (eT0) { /* ignore */ }
        }
        if (fso.FileExists(bakPath)) {
            try {
                fso.DeleteFile(bakPath);
            } catch (eT1) { /* ignore */ }
        }

        var exec = 'ffmpeg.exe -y -i "' + vidPath + '" -vf "' + vfFilter + '" -map_metadata 0 -map_chapters 0 -map 0:v:0 -map "0:a?" -c:v ' + vEnc + ' -c:a copy "' + tmpPath + '"';
        DOpus.Output(logTitle + ": " + exec);

        try {
            var exitCode = shell.Run(exec, 0, true);
            if (exitCode != 0) {
                DOpus.Output(logTitle + " failed (exit " + exitCode + "): " + vidItem.name);
                fail++;
                continue;
            }
            if (!fso.FileExists(tmpPath)) {
                DOpus.Output(logTitle + ": output missing after ffmpeg: " + vidItem.name);
                fail++;
                continue;
            }

            try {
                fso.MoveFile(vidPath, bakPath);
            } catch (eRen) {
                DOpus.Output(logTitle + ": could not rename original (in use?): " + vidItem.name + " — left temp: " + tmpPath);
                fail++;
                continue;
            }
            try {
                fso.MoveFile(tmpPath, vidPath);
            } catch (eMv) {
                try {
                    fso.MoveFile(bakPath, vidPath);
                } catch (eRest) { /* ignore */ }
                DOpus.Output(logTitle + ": could not replace file, restored original: " + vidItem.name);
                fail++;
                continue;
            }
            try {
                fso.DeleteFile(bakPath);
            } catch (eDel) { /* leave backup if locked */ }
            ok++;
        } catch (ex) {
            DOpus.Output(logTitle + " error on " + vidItem.name + ": " + ex.message);
            fail++;
        }
    }

    if (fail > 0 && ok === 0) {
        thumbErr(shell, "All " + fail + " file(s) failed. See DOpus Script Output.", logTitle);
    } else if (fail > 0) {
        thumbInfo(shell, "Finished with errors. OK: " + ok + ", Failed: " + fail + ". Details in Script Output.", logTitle);
    } else {
        thumbInfo(shell, logTitle + " finished. Files updated: " + ok, logTitle);
    }
    try {
        clickData.func.command.RunCommand("Go REFRESH");
    } catch (eRf) { /* ignore */ }
}

/**
 * Run conversion using saved or dialog values (shared by OK and Ctrl+click).
 * modeIndex: 0 = video, 1 = audio. formatName must match a preset name or first preset is used.
 */
function runConvertWithSelectedFiles(clickData, tab, fso, shell, videoFormats, audioFormats, modeIndex, formatName, qualityStr) {
    if (tab.selstats.selfiles == 0) {
        DOpus.Output("[Converter ERROR] No files selected to convert. Select files in the lister, then run the converter again.");
        return;
    }
    var isVideo = (modeIndex == 0);
    var formats = isVideo ? videoFormats : audioFormats;
    var formatIndex = -1;
    var fi;
    for (fi = 0; fi < formats.length; fi++) {
        if (formats[fi].name === formatName) {
            formatIndex = fi;
            break;
        }
    }
    if (formatIndex < 0 || formatIndex >= formats.length) {
        formatIndex = 0;
        if (formatName) {
            DOpus.Output("[Converter] Unknown saved format \"" + formatName + "\"; using first preset.");
        }
    }
    var fmt = formats[formatIndex];
    var qStr = (qualityStr + "").replace(/^\s+|\s+$/g, "");
    if (!qStr) {
        qStr = "23";
    }

    DOpus.Output("Mode index: " + modeIndex);
    DOpus.Output("Format index: " + formatIndex);
    DOpus.Output("Quality: " + qStr);
    DOpus.Output("Mode: " + (isVideo ? "Video" : "Audio"));
    DOpus.Output("Format: " + fmt.name);

    var processed = 0;
    var failed = 0;
    var enumerator = new Enumerator(tab.selected_files);

    for (; !enumerator.atEnd(); enumerator.moveNext()) {
        var item = enumerator.item();
        var outPath = item.path + "\\" + item.name_stem + fmt.ext;

        var counter = 1;
        while (fso.FileExists(outPath)) {
            outPath = item.path + "\\" + item.name_stem + "_" + counter + fmt.ext;
            counter++;
        }

        var exec;
        if (isVideo) {
            var vcodec = fmt.codec;
            if (fmt.crf) {
                vcodec = vcodec.replace(/-crf\s+\d+/, "-crf " + qStr);
            }
            exec = 'ffmpeg.exe -i "' + item.realpath + '" -map_metadata 0 -map_chapters 0 -c:v ' + vcodec + ' -y "' + outPath + '"';
        } else {
            exec = 'ffmpeg.exe -i "' + item.realpath + '" -map_metadata 0 -map_chapters 0 -vn -c:a ' + fmt.codec + ' -y "' + outPath + '"';
        }

        DOpus.Output("Converting: " + item.name + " -> " + fmt.name);

        try {
            var exitCode = shell.Run(exec, 0, true);
            if (exitCode == 0) {
                processed++;
                DOpus.Output("Success: " + outPath);
            } else {
                DOpus.Output("Failed (code " + exitCode + "): " + item.name);
                failed++;
            }
        } catch (e) {
            DOpus.Output("Error: " + e.message);
            failed++;
        }
    }

    var summary = "Conversion finished. Successful: " + processed;
    if (failed > 0) {
        summary += ", Failed: " + failed;
    }
    DOpus.Output("[Converter] " + summary);

    try {
        clickData.func.command.RunCommand("Go REFRESH");
    } catch (eRf) { /* ignore */ }
}

function OnClick(clickData) {
    var tab = clickData.func.sourcetab;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var shell = new ActiveXObject("WScript.Shell");

    // Format definitions (crf: include Quality edit value in -crf for this preset)
    var videoFormats = [
        { name: "MP4 H.264 (Fast)", ext: ".mp4", codec: "libx264 -crf 23 -preset fast -c:a aac -b:a 192k -pix_fmt yuv420p", crf: true },
        { name: "MP4 H.265/HEVC", ext: ".mp4", codec: "libx265 -crf 28 -preset fast -c:a aac -b:a 192k -pix_fmt yuv420p", crf: true },
        { name: "MP4 YouTube Ready", ext: ".mp4", codec: "libx264 -crf 23 -preset slow -c:a aac -b:a 256k -pix_fmt yuv420p -movflags +faststart", crf: true },
        { name: "MOV ProRes 422", ext: ".mov", codec: "prores -profile:v 2 -c:a pcm_s16le", crf: false },
        { name: "MOV ProRes 4444", ext: ".mov", codec: "prores -profile:v 3 -alpha_bits 0 -c:a pcm_s16le", crf: false },
        { name: "MOV H.264", ext: ".mov", codec: "libx264 -crf 23 -preset fast -c:a aac -b:a 192k -pix_fmt yuv420p", crf: true },
        { name: "WebM VP9", ext: ".webm", codec: "libvpx-vp9 -crf 30 -b:v 0 -c:a libopus -b:a 128k", crf: true },
        { name: "AVI Uncompressed", ext: ".avi", codec: "rawvideo -c:a pcm_s16le", crf: false }
    ];

    var audioFormats = [
        { name: "MP3 High Quality (320k)", ext: ".mp3", codec: "libmp3lame -q:a 0 -b:a 320k" },
        { name: "MP3 Standard (192k)", ext: ".mp3", codec: "libmp3lame -q:a 2 -b:a 192k" },
        { name: "MP3 Voice (64k)", ext: ".mp3", codec: "libmp3lame -q:a 6 -b:a 64k" },
        { name: "FLAC Lossless", ext: ".flac", codec: "flac" },
        { name: "WAV PCM 16-bit", ext: ".wav", codec: "pcm_s16le" },
        { name: "WAV PCM 24-bit", ext: ".wav", codec: "pcm_s24le" },
        { name: "AAC M4A", ext: ".m4a", codec: "aac -b:a 256k" },
        { name: "OGG Vorbis", ext: ".ogg", codec: "libvorbis -q:a 6" },
        { name: "OGG Opus", ext: ".ogg", codec: "libopus -b:a 128k" }
    ];

    var qualStr = "";
    try {
        qualStr = String(clickData.func.qualifiers + "");
    } catch (eq) {
        qualStr = "";
    }
    if (qualStr.indexOf("ctrl") >= 0) {
        var lastQuick = loadLastSettings(shell, fso);
        var act = lastQuick.lastAction || "convert";
        if (act != "convert" && act != "cover" && act != "mono" && act != "splitav" && act != "splitch" && act != "rotatecw" && act != "rotateccw" && act != "fliph" && act != "flipv") {
            act = "convert";
        }
        DOpus.Output("ffmpeg: Ctrl+click — last action: " + act + " (no dialog)");
        if (act == "cover") {
            runSplitOrCombineCover(clickData, fso, shell);
            return;
        }
        if (act == "mono") {
            runAudioToMono(clickData, fso, shell);
            return;
        }
        if (act == "splitav") {
            runSplitAvCopy(clickData, fso, shell);
            return;
        }
        if (act == "splitch") {
            runExtractAllAudioChannels(clickData, fso, shell);
            return;
        }
        if (act == "rotatecw") {
            runVideoTransform(clickData, fso, shell, "transpose=1", "Rotate 90° CW");
            return;
        }
        if (act == "rotateccw") {
            runVideoTransform(clickData, fso, shell, "transpose=2", "Rotate 90° CCW");
            return;
        }
        if (act == "fliph") {
            runVideoTransform(clickData, fso, shell, "hflip", "Flip horizontal");
            return;
        }
        if (act == "flipv") {
            runVideoTransform(clickData, fso, shell, "vflip", "Flip vertical");
            return;
        }
        runConvertWithSelectedFiles(clickData, tab, fso, shell, videoFormats, audioFormats, lastQuick.mode, lastQuick.formatName, lastQuick.quality);
        return;
    }

    // Create detached dialog
    var dlg = DOpus.dlg;
    dlg.window = clickData.func.sourcetab;
    dlg.template = "DOpus_ffmpeg_Dlg";
    dlg.detach = true;

    // Create dialog first (hidden)
    dlg.Create();

    // Get control references
    var modeCtrl = dlg.control("mode_combo");
    var formatCtrl = dlg.control("format_combo");
    var qualityCtrl = dlg.control("quality_edit");
    var qualityLabelCtrl = dlg.control("quality_label");
    var qualityHintCtrl = dlg.control("quality_hint");

    function qualityApplicable(isVideoMode, fmtIdx) {
        if (!isVideoMode) {
            return false;
        }
        if (fmtIdx < 0 || fmtIdx >= videoFormats.length) {
            return false;
        }
        return videoFormats[fmtIdx].crf === true;
    }

    function syncQualityControlsEnabled() {
        var modeItem = modeCtrl.value;
        var isVideoMode = (modeItem.index == 0);
        var fmtItem = formatCtrl.value;
        var fmtIdx = fmtItem ? fmtItem.index : 0;
        var on = qualityApplicable(isVideoMode, fmtIdx);
        qualityCtrl.enabled = on;
        qualityLabelCtrl.enabled = on;
        qualityHintCtrl.enabled = on;
    }

    // Function to populate format dropdown based on mode
    function populateFormats(isVideo) {
        var formats = isVideo ? videoFormats : audioFormats;

        // Clear existing items
        formatCtrl.RemoveItem(-1);

        // Add new items
        for (var i = 0; i < formats.length; i++) {
            formatCtrl.AddItem(formats[i].name, formats[i].name);
        }

        // Select first item (index 0) unless formatName provided
        if (formats.length > 0) {
            formatCtrl.SelectItem(0);
        }
    }

    // Restore last used settings
    var last = loadLastSettings(shell, fso);
    var modeIdx = (last.mode === 1) ? 1 : 0;
    modeCtrl.SelectItem(modeIdx);
    populateFormats(modeIdx === 0);
    if (last.formatName) {
        try {
            var fmtItem = formatCtrl.GetItemByName(last.formatName);
            if (fmtItem) formatCtrl.SelectItem(fmtItem);
        } catch (e) { /* keep default selection */ }
    }
    qualityCtrl.value = last.quality || "23";
    syncQualityControlsEnabled();

    // Show the fully initialized dialog
    dlg.Show();

    // Variable to track if user clicked OK or Cancel
    var dialogResult = 0;
    /** True after rotate/flip, Split/combine cover, Audio→mono, Split/combine AV, or Extract channels — skip Convert (dlg.result may not be string "0"). */
    var dialogClosedAfterTool = false;

    // Message loop to handle events
    while (true) {
        var msg = dlg.GetMsg();

        // Exit if dialog closed
        if (!msg.result) {
            dialogResult = dlg.result;
            break;
        }

        if (msg.event == "click" && msg.control == "split_cover_btn") {
            saveLastActionOnly(shell, fso, "cover");
            runSplitOrCombineCover(clickData, fso, shell);
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }

        if (msg.event == "click" && msg.control == "rotate90_cw_btn") {
            saveLastActionOnly(shell, fso, "rotatecw");
            runVideoTransform(clickData, fso, shell, "transpose=1", "Rotate 90° CW");
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }
        if (msg.event == "click" && msg.control == "rotate90_ccw_btn") {
            saveLastActionOnly(shell, fso, "rotateccw");
            runVideoTransform(clickData, fso, shell, "transpose=2", "Rotate 90° CCW");
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }
        if (msg.event == "click" && msg.control == "flip_h_btn") {
            saveLastActionOnly(shell, fso, "fliph");
            runVideoTransform(clickData, fso, shell, "hflip", "Flip horizontal");
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }
        if (msg.event == "click" && msg.control == "flip_v_btn") {
            saveLastActionOnly(shell, fso, "flipv");
            runVideoTransform(clickData, fso, shell, "vflip", "Flip vertical");
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }

        if (msg.event == "click" && msg.control == "audio_mono_btn") {
            saveLastActionOnly(shell, fso, "mono");
            runAudioToMono(clickData, fso, shell);
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }

        if (msg.event == "click" && msg.control == "split_av_btn") {
            saveLastActionOnly(shell, fso, "splitav");
            runSplitAvCopy(clickData, fso, shell);
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }

        if (msg.event == "click" && msg.control == "extract_ch_btn") {
            saveLastActionOnly(shell, fso, "splitch");
            runExtractAllAudioChannels(clickData, fso, shell);
            dialogClosedAfterTool = true;
            dlg.EndDlg("0");
            dialogResult = dlg.result;
            break;
        }

        // Handle selection change events
        if (msg.event == "selchange") {
            if (msg.control == "mode_combo") {
                var modeItem2 = modeCtrl.value;
                var isVideo2 = (modeItem2.index == 0);
                populateFormats(isVideo2);
                syncQualityControlsEnabled();
            } else if (msg.control == "format_combo") {
                syncQualityControlsEnabled();
            }
        }
    }

    // Check if user clicked OK (close="1") or Cancel (close="2")
    // dlg.result will be "1" for OK, "2" for Cancel, or "0" for window close / tool buttons
    if (dialogClosedAfterTool) {
        return;
    }
    if (dialogResult == "2") {
        DOpus.Output("Dialog cancelled");
        return;
    }
    if (dialogResult == "0" || dialogResult === 0 || dialogResult == "") {
        return;
    }

    if (tab.selstats.selfiles == 0) {
        DOpus.Output("[Converter ERROR] No files selected to convert. Select files in the lister, then run the converter again.");
        return;
    }

    // Get final values
    var modeItem = modeCtrl.value;
    var formatItem = formatCtrl.value;
    var quality = qualityCtrl.value;

    var modeIndex = modeItem.index;

    saveLastSettings(shell, fso, modeIndex, formatItem.name, quality, "convert");

    runConvertWithSelectedFiles(clickData, tab, fso, shell, videoFormats, audioFormats, modeIndex, formatItem.name, quality);
}
