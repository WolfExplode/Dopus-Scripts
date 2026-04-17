// yt-dlp Downloader — reads URL from clipboard and downloads audio or video
var SETTINGS_FILE = null;

// JScript (ES3) has no String.trim()
function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

// FSO TextStream.WriteLine throws "Invalid procedure call or argument" if the string contains \0
function safeWriteStr(s) {
    return String(s).replace(/\x00/g, "");
}

// PowerShell single-quoted literal: only ' is escaped as ''
function escapePsSingleQuoted(s) {
    return safeWriteStr(s).replace(/'/g, "''");
}

// Literal text before %(…) in -o template; % -> %%, " -> ' so -o "…" stays valid
function escapeYtdlpOutputPrefix(s) {
    var t = trimStr(s);
    if (!t) return "";
    t = t.replace(/%/g, "%%");
    t = t.replace(/"/g, "'");
    t = t.replace(/[\r\n]/g, " ");
    return t;
}

// PowerShell: compare yt-dlp --version to GitHub latest; pip upgrade only if needed
function writeYtDlpUpdateBlock(ps1) {
    var lines = [
        "$ErrorActionPreference = 'Continue'",
        "try {",
        "    $localVer = $null",
        "    $yv = & yt-dlp --version 2>&1",
        "    if ($LASTEXITCODE -eq 0 -and $yv) {",
        "        $localVer = ($yv | Select-Object -First 1).ToString().Trim()",
        "    }",
        "    $latestVer = $null",
        "    try {",
        "        $rel = Invoke-RestMethod -Uri 'https://api.github.com/repos/yt-dlp/yt-dlp/releases/latest' -UseBasicParsing -TimeoutSec 15",
        "        $latestVer = ($rel.tag_name -replace '^v','').Trim()",
        "    } catch {",
        "        Write-Host 'Could not check latest yt-dlp version on GitHub.'",
        "    }",
        "    $doPip = $false",
        "    if (-not $localVer) {",
        "        $doPip = $true",
        "        Write-Host 'yt-dlp not found or version unreadable; running pip upgrade...'",
        "    }",
        "    elseif ($latestVer -and $localVer -ne $latestVer) {",
        "        $doPip = $true",
        "        Write-Host ('yt-dlp update: local ' + $localVer + ' -> latest ' + $latestVer)",
        "    }",
        "    if ($doPip) {",
        "        python -m pip install --upgrade yt-dlp",
        "        if ($LASTEXITCODE -ne 0) { py -m pip install --upgrade yt-dlp }",
        "        if ($LASTEXITCODE -ne 0) { Write-Host 'pip upgrade failed; install Python/pip or update yt-dlp manually (e.g. yt-dlp -U).' }",
        "    } elseif ($latestVer) {",
        "        Write-Host ('yt-dlp is up to date (' + $localVer + ').')",
        "    } else {",
        "        Write-Host 'Skipping pip update (could not fetch latest release from GitHub).'",
        "    }",
        "} catch {",
        "    Write-Host ('Update check error: ' + $_.Exception.Message)",
        "}",
        "Write-Host ''"
    ];
    for (var i = 0; i < lines.length; i++) {
        ps1.WriteLine(lines[i]);
    }
}

// PowerShell: move console to bottom-right of primary monitor work area (Win32 + Forms)
function writeConsoleWindowBottomRight(ps1) {
    var lines = [
        "try {",
        "    Add-Type @\"",
        "using System;",
        "using System.Runtime.InteropServices;",
        "public struct YtdlpRect { public int Left; public int Top; public int Right; public int Bottom; }",
        "public static class YtdlpConsoleWin {",
        "    [DllImport(\"kernel32.dll\")] public static extern IntPtr GetConsoleWindow();",
        "    [DllImport(\"user32.dll\")] public static extern bool GetWindowRect(IntPtr hWnd, ref YtdlpRect lpRect);",
        "    [DllImport(\"user32.dll\")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);",
        "}",
        "\"@",
        "    Add-Type -AssemblyName System.Windows.Forms",
        "    $hwnd = [YtdlpConsoleWin]::GetConsoleWindow()",
        "    if ($hwnd -ne [IntPtr]::Zero) {",
        "        $r = New-Object YtdlpRect",
        "        [void][YtdlpConsoleWin]::GetWindowRect($hwnd, [ref]$r)",
        "        $winW = $r.Right - $r.Left",
        "        $winH = $r.Bottom - $r.Top",
        "        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea",
        "        $x = [Math]::Max($wa.Left, $wa.Right - $winW)",
        "        $y = [Math]::Max($wa.Top, $wa.Bottom - $winH)",
        "        $flags = [uint32]0x0001 -bor 0x0004 -bor 0x0040",
        "        [void][YtdlpConsoleWin]::SetWindowPos($hwnd, [IntPtr]::Zero, $x, $y, 0, 0, $flags)",
        "    }",
        "} catch { }",
        ""
    ];
    for (var i = 0; i < lines.length; i++) {
        ps1.WriteLine(lines[i]);
    }
}

function getSettingsPath(shell) {
    if (!SETTINGS_FILE) {
        SETTINGS_FILE = shell.ExpandEnvironmentStrings("%APPDATA%") + "\\DOpus_ytdlp_settings.ini";
    }
    return SETTINGS_FILE;
}

function loadSettings(shell, fso) {
    var out = { mode: 0, cookies: 0, metadata: 0, dateprefix: 1, fileprefix: "", overwrite: 0, update: 0, keepps: 0 };
    try {
        var path = getSettingsPath(shell);
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
                    if (key === "mode") out.mode = parseInt(val, 10) || 0;
                    else if (key === "cookies") out.cookies = parseInt(val, 10) || 0;
                    else if (key === "metadata") {
                        var mv = parseInt(val, 10);
                        out.metadata = (mv === 0) ? 0 : 1;
                    }
                    else if (key === "dateprefix") {
                        var dpv = parseInt(val, 10);
                        out.dateprefix = (dpv === 0) ? 0 : 1;
                    }
                    else if (key === "fileprefix") out.fileprefix = val;
                    else if (key === "overwrite") out.overwrite = parseInt(val, 10) || 0;
                    else if (key === "update") out.update = parseInt(val, 10) || 0;
                    else if (key === "keepps") out.keepps = parseInt(val, 10) || 0;
                }
            }
        }
    } catch (e) { /* use defaults */ }
    return out;
}

function saveSettings(shell, fso, mode, cookies, metadata, dateprefix, fileprefix, overwrite, update, keepps) {
    try {
        var path = getSettingsPath(shell);
        var stream = fso.OpenTextFile(path, 2, true);
        stream.WriteLine("mode=" + mode);
        stream.WriteLine("cookies=" + cookies);
        stream.WriteLine("metadata=" + metadata);
        stream.WriteLine("dateprefix=" + dateprefix);
        stream.WriteLine("fileprefix=" + fileprefix.replace(/[\r\n]/g, " "));
        stream.WriteLine("overwrite=" + overwrite);
        stream.WriteLine("update=" + update);
        stream.WriteLine("keepps=" + keepps);
        stream.Close();
    } catch (e) { /* ignore */ }
}

function OnClick(clickData) {
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");

    var destPath = safeWriteStr(clickData.func.sourcetab.path + "");

    // Read clipboard (DOpus API is GetClip("text"), not GetClipText)
    var url = "";
    try {
        var clip = DOpus.GetClip("text");
        if (clip)
            url = trimStr(clip);
    } catch (e) {
        url = "";
    }

    var qualStr = "";
    try {
        qualStr = String(clickData.func.qualifiers + "");
    } catch (eq) {
        qualStr = "";
    }
    var skipUi = qualStr.indexOf("ctrl") >= 0;

    var finalUrl;
    var filePrefixRaw;
    var isAudio;
    var useCookies;
    var includeMetadata;
    var datePrefix;
    var allowOverwrite;
    var doUpdate;
    var keepPsOpen;

    if (skipUi) {
        var savedQuick = loadSettings(shell, fso);
        finalUrl = safeWriteStr(trimStr(url));
        filePrefixRaw = String(savedQuick.fileprefix || "");
        isAudio = (savedQuick.mode !== 1);
        useCookies = (savedQuick.cookies === 1);
        includeMetadata = (savedQuick.metadata === 1);
        datePrefix = (savedQuick.dateprefix === 1);
        allowOverwrite = (savedQuick.overwrite === 1);
        doUpdate = (savedQuick.update === 1);
        keepPsOpen = (savedQuick.keepps === 1);

        if (!finalUrl) {
            shell.Popup("No URL in clipboard. Ctrl+click uses saved settings and the clipboard URL.", 0, "yt-dlp", 48);
            return;
        }
        DOpus.Output("yt-dlp: Ctrl+click — saved settings, no dialog");
    } else {
        var dlg = DOpus.dlg;
        dlg.window = clickData.func.sourcetab;
        dlg.template = "YtDlpDlg";
        dlg.detach = true;
        dlg.Create();

        dlg.control("url_edit").value = url;

        var saved = loadSettings(shell, fso);
        dlg.control("prefix_edit").value = saved.fileprefix || "";
        if (saved.mode === 1) {
            dlg.control("video_radio").value = true;
        } else {
            dlg.control("audio_radio").value = true;
        }
        dlg.control("metadata_check").value = (saved.metadata === 1);
        dlg.control("dateprefix_check").value = (saved.dateprefix === 1);
        dlg.control("cookies_check").value = (saved.cookies === 1);
        dlg.control("overwrite_check").value = (saved.overwrite === 1);
        dlg.control("update_check").value = (saved.update === 1);
        dlg.control("keepps_check").value = (saved.keepps === 1);

        dlg.Show();

        var dialogResult = 0;
        while (true) {
            var msg = dlg.GetMsg();
            if (!msg.result) {
                dialogResult = dlg.result;
                break;
            }
        }

        if (dialogResult == "0" || dialogResult == "2") {
            DOpus.Output("yt-dlp: cancelled");
            return;
        }

        finalUrl = safeWriteStr(trimStr(dlg.control("url_edit").value));
        filePrefixRaw = String(dlg.control("prefix_edit").value);
        isAudio = dlg.control("audio_radio").value;
        useCookies = dlg.control("cookies_check").value;
        includeMetadata = dlg.control("metadata_check").value;
        datePrefix = dlg.control("dateprefix_check").value;
        allowOverwrite = dlg.control("overwrite_check").value;
        doUpdate = dlg.control("update_check").value;
        keepPsOpen = dlg.control("keepps_check").value;

        if (!finalUrl) {
            shell.Popup("No URL provided.", 0, "yt-dlp", 48);
            return;
        }

        saveSettings(shell, fso, isAudio ? 0 : 1, useCookies ? 1 : 0, includeMetadata ? 1 : 0, datePrefix ? 1 : 0, trimStr(filePrefixRaw), allowOverwrite ? 1 : 0, doUpdate ? 1 : 0, keepPsOpen ? 1 : 0);
    }

    var filePrefixEsc = escapeYtdlpOutputPrefix(filePrefixRaw);

    var cookiesFromBrowser = " --cookies-from-browser firefox";
    var overwriteArg = allowOverwrite ? " --force-overwrites" : " --no-overwrites";
    // Lets yt-dlp download EJS solver scripts (needed for YouTube + Deno / JS challenges; see yt-dlp wiki EJS)
    var ejsArg = " --remote-components ejs:github";
    var metaAudio = " --extract-audio --audio-format best --add-metadata --embed-thumbnail --embed-subs --parse-metadata \":(?P<chapters>)\"";
    var metaVideo = " --add-metadata --embed-thumbnail --write-auto-subs --embed-subs";
    var innerCore = datePrefix
        ? "[%(upload_date>%m-%d-%Y)s] %(title)s.%(ext)s"
        : "%(title)s.%(ext)s";
    var inner = filePrefixEsc ? (filePrefixEsc + innerCore) : innerCore;
    var outTemplate = '"' + inner + '"';

    // Build yt-dlp argument string
    // Written into a .ps1 file so % format specifiers are never touched by cmd.exe
    // Plain mode: no metadata/embed flags; with metadata: previous full options
    // URL is appended with PS single-quoted literal so & ? etc. never break parsing; avoids NUL/ANSI WriteLine issues
    var ytArgsBody;
    if (isAudio) {
        ytArgsBody = "-o " + outTemplate
               + ' -f bestaudio'
               + overwriteArg
               + ejsArg
               + (includeMetadata ? metaAudio : "");
    } else {
        ytArgsBody = "-o " + outTemplate
               + overwriteArg
               + ejsArg
               + (includeMetadata ? metaVideo : "");
    }

    var ytArgsFirst = ytArgsBody + (useCookies ? cookiesFromBrowser : "");

    // Write a temp PowerShell script to avoid cmd.exe expanding % characters
    var tempPs1 = shell.ExpandEnvironmentStrings("%TEMP%") + "\\yt-dlp-run.ps1";
    try {
        // unicode=true: UTF-16LE + BOM so non-ANSI paths / pasted text do not break WriteLine
        var ps1 = fso.CreateTextFile(tempPs1, true, true);
        ps1.WriteLine("Set-Location '" + escapePsSingleQuoted(destPath) + "'");
        writeConsoleWindowBottomRight(ps1);
        if (doUpdate) {
            writeYtDlpUpdateBlock(ps1);
        }
        ps1.WriteLine("yt-dlp " + ytArgsFirst + " '" + escapePsSingleQuoted(finalUrl) + "'");
        if (useCookies) {
            ps1.WriteLine("$code = $LASTEXITCODE");
            ps1.WriteLine("if ($code -ne 0) {");
            ps1.WriteLine("    Write-Host ''");
            ps1.WriteLine("    Write-Host \"yt-dlp failed (exit $code); retrying without browser cookies...\" -ForegroundColor Yellow");
            ps1.WriteLine("    yt-dlp " + ytArgsBody + " '" + escapePsSingleQuoted(finalUrl) + "'");
            ps1.WriteLine("}");
        }
        ps1.Close();
    } catch (e) {
        var errMsg = "Failed to write temp script.";
        try {
            if (e && e.message) errMsg += " " + e.message;
        } catch (x) { /* ignore */ }
        shell.Popup(errMsg, 0, "yt-dlp Error", 16);
        return;
    }

    DOpus.Output("yt-dlp | URL: " + finalUrl);
    DOpus.Output("yt-dlp | Dest: " + destPath);
    DOpus.Output("yt-dlp | Mode: " + (isAudio ? "Audio" : "Video") + " | Metadata: " + includeMetadata + " | Date prefix: " + datePrefix + " | File prefix: " + (trimStr(filePrefixRaw) ? trimStr(filePrefixRaw) : "(none)") + " | Cookies: " + useCookies + " | Overwrite: " + allowOverwrite + " | Update check: " + doUpdate + " | Keep PS: " + keepPsOpen + " | UI: " + (skipUi ? "skipped (Ctrl)" : "dialog"));

    var psCmd = keepPsOpen
        ? 'powershell -NoExit -ExecutionPolicy Bypass -File "' + tempPs1 + '"'
        : 'powershell -ExecutionPolicy Bypass -File "' + tempPs1 + '"';
    shell.Run(psCmd, 1, false);
}
