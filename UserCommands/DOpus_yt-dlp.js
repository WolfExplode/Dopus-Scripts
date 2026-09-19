// yt-dlp Downloader — reads URL from clipboard and downloads audio or video
var SETTINGS_FILE = null;

// JScript (ES3) has no String.trim()
function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

// Make a string safe to place on a single PowerShell line
function oneLine(s) {
    return safeWriteStr(s).replace(/[\r\n\t]+/g, " ");
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

// Split a command line into tokens, honouring double-quoted runs
function tokenizeCommandLine(s) {
    var out = [];
    var cur = "";
    var inQuote = false;
    var started = false;
    for (var i = 0; i < s.length; i++) {
        var c = s.charAt(i);
        if (c === '"') {
            inQuote = !inQuote;
            started = true;
        } else if (!inQuote && (c === " " || c === "\t" || c === "\r" || c === "\n")) {
            if (started) { out.push(cur); cur = ""; started = false; }
        } else {
            cur += c;
            started = true;
        }
    }
    if (started) out.push(cur);
    return out;
}

// Quote one token for the PowerShell command line; bare flags are left alone
function psQuoteArg(tok) {
    if (/^--?[A-Za-z][A-Za-z0-9-]*$/.test(tok)) return tok;
    return "'" + escapePsSingleQuoted(tok) + "'";
}

// Accepts a full "yt-dlp ... URL" command line (e.g. copied from the YouTube Clipper browser
// extension) and splits it into the URL plus the remaining arguments. Returns null when the
// text is not such a command, so a bare URL still works exactly as before.
// -o/--output is deliberately dropped: this script builds its own template, and a template
// copied from a cmd.exe-targeted command has its % signs doubled, which would be wrong here.
function parseYtdlpCommand(text) {
    var t = trimStr(text);
    if (!/^yt-dlp(\.exe)?[ \t]/i.test(t)) return null;

    var toks = tokenizeCommandLine(t);
    var url = "";
    var extras = [];

    for (var i = 1; i < toks.length; i++) {
        var tok = toks[i];
        if (tok === "-o" || tok === "--output") {
            i++;
            continue;
        }
        if (/^https?:\/\//i.test(tok)) {
            url = tok;
            continue;
        }
        extras.push(psQuoteArg(tok));
    }

    return { url: url, args: extras.join(" ") };
}

// PowerShell: compare yt-dlp --version to GitHub latest; pip upgrade only if needed
function writeYtDlpUpdateBlock(ps1) {
    var lines = [
        "$ErrorActionPreference = \"Continue\"",
        "try {",
        "    $localVer = $null",
        "    $yv = & $_yt --version 2>&1",
        "    if ($LASTEXITCODE -eq 0 -and $yv) {",
        "        $localVer = ($yv | Select-Object -First 1).ToString().Trim()",
        "    }",
        "    $latestVer = $null",
        "    try {",
        "        $rel = Invoke-RestMethod -Uri \"https://api.github.com/repos/yt-dlp/yt-dlp/releases/latest\" -UseBasicParsing -TimeoutSec 15",
        "        $latestVer = ($rel.tag_name -replace \"^v\",\"\").Trim()",
        "    } catch {",
        "        Write-Host \"Could not check latest yt-dlp version on GitHub.\"",
        "    }",
        "    $doPip = $false",
        "    if (-not $localVer) {",
        "        $doPip = $true",
        "        Write-Host \"yt-dlp not found or version unreadable; running pip upgrade...\"",
        "    }",
        "    elseif ($latestVer -and $localVer -ne $latestVer) {",
        "        $doPip = $true",
        "        Write-Host (\"yt-dlp update: local \" + $localVer + \" -> latest \" + $latestVer)",
        "    }",
        "    if ($doPip) {",
        "        if ($_ytpy) {",
        "            Write-Host (\"Upgrading yt-dlp for \" + $_ytpy)",
        "            & $_ytpy -m pip install --upgrade yt-dlp",
        "        } else {",
        "            python -m pip install --upgrade yt-dlp",
        "            if ($LASTEXITCODE -ne 0) { py -m pip install --upgrade yt-dlp }",
        "        }",
        "        if ($LASTEXITCODE -ne 0) { Write-Host \"pip upgrade failed; install Python/pip or update yt-dlp manually (e.g. yt-dlp -U).\" }",
        "        $yv2 = & $_yt --version 2>&1",
        "        if ($LASTEXITCODE -eq 0 -and $yv2) { Write-Host (\"yt-dlp now at \" + ($yv2 | Select-Object -First 1).ToString().Trim()) }",
        "    } elseif ($latestVer) {",
        "        Write-Host (\"yt-dlp is up to date (\" + $localVer + \").\")",
        "    } else {",
        "        Write-Host \"Skipping pip update (could not fetch latest release from GitHub).\"",
        "    }",
        "} catch {",
        "    Write-Host (\"Update check error: \" + $_.Exception.Message)",
        "}",
        "Write-Host \"\""
    ];
    for (var i = 0; i < lines.length; i++) {
        ps1.WriteLine(lines[i]);
    }
}

// PowerShell: resolve actual yt-dlp .exe, bypassing pyenv .bat shim so %(...) in -o template
// is not mangled by cmd.exe batch % expansion inside `call pyenv exec %~n0 %*`
function writeYtdlpFinder(ps1) {
    var lines = [
        "$_yt = 'yt-dlp'",
        "try {",
        "    $_ytp = ((pyenv which yt-dlp 2>$null) -join '').Trim()",
        "    if ($_ytp -and (Test-Path $_ytp)) {",
        "        $_yt = $_ytp",
        "    } else {",
        "        $_cmd = Get-Command yt-dlp -ErrorAction SilentlyContinue",
        "        if ($_cmd -and $_cmd.Source -and (Test-Path $_cmd.Source)) { $_yt = $_cmd.Source }",
        "    }",
        "} catch { }",
        "",
        "$_ytpy = $null",
        "try {",
        "    if ($_yt -ne 'yt-dlp' -and $_yt -like '*.exe') {",
        "        $_p = Join-Path (Split-Path (Split-Path $_yt -Parent) -Parent) 'python.exe'",
        "        if (Test-Path $_p) { $_ytpy = $_p }",
        "    }",
        "} catch { }",
        ""
    ];
    for (var i = 0; i < lines.length; i++) { ps1.WriteLine(lines[i]); }
}

// PowerShell: move console to bottom-right of primary monitor work area (Win32 + Forms)
function writeConsoleWindowBottomRight(ps1) {
    var lines = [
        "try {",
        "Add-Type @\"",
        "using System;",
        "using System.Runtime.InteropServices;",
        "public struct YtdlpRect { public int Left; public int Top; public int Right; public int Bottom; }",
        "public static class YtdlpConsoleWin {",
        "    [DllImport(\"kernel32.dll\")] public static extern IntPtr GetConsoleWindow();",
        "    [DllImport(\"user32.dll\")] public static extern bool GetWindowRect(IntPtr hWnd, ref YtdlpRect lpRect);",
        "    [DllImport(\"user32.dll\")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);",
        "}",
        "\"@",
        "Add-Type -AssemblyName System.Windows.Forms",
        "$hwnd = [YtdlpConsoleWin]::GetConsoleWindow()",
        "if ($hwnd -ne [IntPtr]::Zero) {",
        "    $r = New-Object YtdlpRect",
        "    [void][YtdlpConsoleWin]::GetWindowRect($hwnd, [ref]$r)",
        "    $winW = $r.Right - $r.Left",
        "    $winH = $r.Bottom - $r.Top",
        "    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea",
        "    $x = [Math]::Max($wa.Left, $wa.Right - $winW)",
        "    $y = [Math]::Max($wa.Top, $wa.Bottom - $winH)",
        "    $flags = [uint32]0x0001 -bor 0x0004 -bor 0x0040",
        "    [void][YtdlpConsoleWin]::SetWindowPos($hwnd, [IntPtr]::Zero, $x, $y, 0, 0, $flags)",
        "}",
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
    var out = { mode: 0, cookies: 0, metadata: 0, dateprefix: 1, fileprefix: "", overwrite: 0, update: 0, keepps: 0, mp4container: 0, impersonate: 0, nocookies: 0 };
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
                    else if (key === "mp4container") {
                        var mp4v = parseInt(val, 10);
                        out.mp4container = (mp4v === 0) ? 0 : 1;
                    }
                    else if (key === "impersonate") out.impersonate = parseInt(val, 10) || 0;
                    else if (key === "nocookies") out.nocookies = parseInt(val, 10) || 0;
                }
            }
        }
    } catch (e) { /* use defaults */ }
    return out;
}

function saveSettings(shell, fso, mode, cookies, metadata, dateprefix, fileprefix, overwrite, update, keepps, mp4container, impersonate, nocookies) {
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
        stream.WriteLine("mp4container=" + mp4container);
        stream.WriteLine("impersonate=" + impersonate);
        stream.WriteLine("nocookies=" + nocookies);
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

    // The clipboard may hold a whole yt-dlp command rather than a bare URL
    var clipCmd = parseYtdlpCommand(url);
    var clipArgs = "";
    if (clipCmd) {
        url = clipCmd.url;
        clipArgs = clipCmd.args;
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
    var mp4Container;
    var useImpersonate;
    var noCookies;
    var extraArgs;

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
        mp4Container = (savedQuick.mp4container === 1);
        useImpersonate = (savedQuick.impersonate === 1);
        noCookies = (savedQuick.nocookies === 1);
        extraArgs = clipArgs;

        if (!finalUrl) {
            shell.Popup("No URL in clipboard. Ctrl+click uses saved settings and the clipboard URL or yt-dlp command.", 0, "yt-dlp", 48);
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
        dlg.control("mp4_check").value = (saved.mp4container === 1);
        dlg.control("impersonate_check").value = (saved.impersonate === 1);
        dlg.control("nocookies_check").value = (saved.nocookies === 1);
        dlg.control("args_edit").value = clipArgs;

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
        mp4Container = dlg.control("mp4_check").value;
        useImpersonate = dlg.control("impersonate_check").value;
        noCookies = dlg.control("nocookies_check").value;
        extraArgs = String(dlg.control("args_edit").value);

        if (!finalUrl) {
            shell.Popup("No URL provided.", 0, "yt-dlp", 48);
            return;
        }

        saveSettings(shell, fso, isAudio ? 0 : 1, useCookies ? 1 : 0, includeMetadata ? 1 : 0, datePrefix ? 1 : 0, trimStr(filePrefixRaw), allowOverwrite ? 1 : 0, doUpdate ? 1 : 0, keepPsOpen ? 1 : 0, mp4Container ? 1 : 0, useImpersonate ? 1 : 0, noCookies ? 1 : 0);
    }

    var filePrefixEsc = escapeYtdlpOutputPrefix(filePrefixRaw);

    // Extra yt-dlp arguments, typed in the dialog or parsed out of a clipboard yt-dlp command.
    // Inserted verbatim into the PowerShell command line, after the flags built above, so they
    // win where yt-dlp lets a later option override an earlier one.
    var extraArgsClean = oneLine(trimStr(extraArgs || ""));
    var extraArgsPart = extraArgsClean ? (" " + extraArgsClean) : "";
    // One video can yield several clips, so keep the range in the name or they overwrite.
    var hasSections = /--download-sections/i.test(extraArgsClean);

    var cookiesFromBrowser = " --cookies-from-browser firefox";
    var overwriteArg = allowOverwrite ? " --force-overwrites" : " --no-overwrites";
    // Lets yt-dlp download EJS solver scripts (needed for YouTube + Deno / JS challenges; see yt-dlp wiki EJS)
    var ejsArg = " --remote-components ejs:github";
    // Spoofs Chrome's TLS/HTTP fingerprint (requires curl_cffi); helps dodge bot-detection blocks
    var impersonateArg = useImpersonate ? " --impersonate chrome" : "";
    // Explicitly disables cookie use, overriding any cookies configured in yt-dlp's own config file
    var noCookiesArg = noCookies ? " --no-cookies" : "";
    // Subtitles are always fetched for the whole video, even when only a section is kept.
    // Embedding that full-length track stretches the Matroska container duration to the source
    // length, and its timestamps are absolute so they no longer line up with the clip.
    // Section downloads are forced onto ffmpeg's single-stream downloader, so a 4K source costs
    // far more wall time than a clip is worth. Cap at 1080p when cutting; full downloads are
    // left alone. Placed before extraArgsPart so a -S/-f in the extra args still wins.
    var sectionResArg = hasSections ? " -S 'res:1080'" : "";
    var subsAudio = hasSections ? "" : " --embed-subs";
    var subsVideo = hasSections ? "" : " --write-auto-subs --embed-subs";
    var metaAudio = " --extract-audio --audio-format best --add-metadata --embed-thumbnail" + subsAudio + " --parse-metadata \":(?P<chapters>)\"";
    var metaVideo = " --add-metadata --embed-thumbnail" + subsVideo;
    var videoMp4Args = (mp4Container && !isAudio) ? " --merge-output-format mp4 --remux-video mp4" : "";
    var sectionSuffix = hasSections ? " %(section_start)s-%(section_end)s" : "";
    var innerCore = datePrefix
        ? "[%(upload_date>%m-%d-%Y)s] %(title)s" + sectionSuffix + ".%(ext)s"
        : "%(title)s" + sectionSuffix + ".%(ext)s";
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
               + impersonateArg
               + noCookiesArg
               + (includeMetadata ? metaAudio : "")
               + extraArgsPart;
    } else {
        ytArgsBody = "-o " + outTemplate
               + overwriteArg
               + ejsArg
               + impersonateArg
               + noCookiesArg
               + videoMp4Args
               + sectionResArg
               + (includeMetadata ? metaVideo : "")
               + extraArgsPart;
    }

    var ytArgsFirst = ytArgsBody + ((useCookies && !noCookies) ? cookiesFromBrowser : "");

    // Unique suffix so concurrent invocations never clobber each other's temp files.
    var runId = String((new Date()).getTime()) + "_" + String(Math.floor(Math.random() * 1000000));

    // Write a temp PowerShell script to avoid cmd.exe expanding % characters
    var tempPs1 = shell.ExpandEnvironmentStrings("%TEMP%") + "\\yt-dlp-run-" + runId + ".ps1";
    try {
        // unicode=true: UTF-16LE + BOM so non-ANSI paths / pasted text do not break WriteLine
        var ps1 = fso.CreateTextFile(tempPs1, true, true);
        ps1.WriteLine("$_dest = '" + escapePsSingleQuoted(oneLine(destPath)) + "'");
        ps1.WriteLine("if (-not (Test-Path -LiteralPath $_dest)) {");
        ps1.WriteLine("    Write-Host \"ERROR: Destination folder does not exist: $_dest\" -ForegroundColor Red");
        ps1.WriteLine("    Read-Host 'Press Enter to close'");
        ps1.WriteLine("    exit 1");
        ps1.WriteLine("}");
        ps1.WriteLine("Set-Location -LiteralPath $_dest");
        writeConsoleWindowBottomRight(ps1);
        writeYtdlpFinder(ps1);
        if (doUpdate) {
            writeYtDlpUpdateBlock(ps1);
        }
        // Write URL to a temp file and use --batch-file so URLs with ? & etc. never hit a cmd.exe
        // argument (the pyenv .bat shim uses `call` internally, which interprets /? as a help flag).
        var tempUrl = shell.ExpandEnvironmentStrings("%TEMP%") + "\\yt-dlp-url-" + runId + ".txt";
        ps1.WriteLine("$_urlf = '" + escapePsSingleQuoted(tempUrl) + "'");
        ps1.WriteLine("[System.IO.File]::WriteAllText($_urlf, '" + escapePsSingleQuoted(oneLine(finalUrl)) + "', [System.Text.Encoding]::UTF8)");
        ps1.WriteLine("& $_yt " + ytArgsFirst + " --batch-file $_urlf");
        if (useCookies && !noCookies) {
            ps1.WriteLine("$code = $LASTEXITCODE");
            ps1.WriteLine("if ($code -ne 0) {");
            ps1.WriteLine("    Write-Host ''");
            ps1.WriteLine("    Write-Host \"yt-dlp failed (exit $code); retrying without browser cookies...\" -ForegroundColor Yellow");
            ps1.WriteLine("    & $_yt " + ytArgsBody + " --batch-file $_urlf");
            ps1.WriteLine("}");
        }
        ps1.WriteLine("Remove-Item $_urlf -Force -ErrorAction SilentlyContinue");
        if (!keepPsOpen) {
            // Delete this run's own script so uniquely-named temps don't accumulate.
            ps1.WriteLine("Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue");
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
    DOpus.Output("yt-dlp | Mode: " + (isAudio ? "Audio" : "Video") + (isAudio ? "" : (" | MP4 container: " + mp4Container)) + " | Metadata: " + includeMetadata + " | Date prefix: " + datePrefix + " | File prefix: " + (trimStr(filePrefixRaw) ? trimStr(filePrefixRaw) : "(none)") + " | Cookies: " + useCookies + " | Overwrite: " + allowOverwrite + " | Update check: " + doUpdate + " | Keep PS: " + keepPsOpen + " | Impersonate Chrome: " + useImpersonate + " | No cookies: " + noCookies + " | UI: " + (skipUi ? "skipped (Ctrl)" : "dialog") + (hasSections ? " | Section: 1080p cap, no subs" : ""));
    if (extraArgsClean) DOpus.Output("yt-dlp | Extra args: " + extraArgsClean);

    var psCmd = keepPsOpen
        ? 'powershell -NoExit -ExecutionPolicy Bypass -File "' + tempPs1 + '"'
        : 'powershell -ExecutionPolicy Bypass -File "' + tempPs1 + '"';
    shell.Run(psCmd, 1, false);
}
