// gallery-dl — Pinterest profiles; cache + settings under %APPDATA%\DOpus_gallery_dl_pinterest\
var GALLERY_DL_EXE = "C:\\Users\\WXP\\AppData\\Local\\Programs\\Python\\Python310\\Scripts\\gallery-dl.exe";

// gallery-dl appends its own "pinterest\username\boardname" under this root.
var PINTEREST_OUTPUT_ROOT = "C:\\Users\\WXP\\Documents\\Pureref\\gallery-dl";

var PIN_URLS = [
    "https://www.pinterest.com/allyfire1281/",
    "https://www.pinterest.com/FedTheBeast/",
    "https://www.pinterest.com/fireally31/"
];
var PROFILE_LABELS = ["allyfire1281", "FedTheBeast", "fireally31"];

// ── helpers ──────────────────────────────────────────────────────────────────

function trimStr(s) { return String(s).replace(/^\s+|\s+$/g, ""); }

function showError(shell, text, title) { shell.Popup(text, 0, title, 16); }
function showInfo(shell, text, title)  { shell.Popup(text, 0, title, 64); }

function psQuote(s) { return String(s).replace(/'/g, "''"); }

function getCacheDir(shell) {
    return shell.ExpandEnvironmentStrings("%APPDATA%") + "\\DOpus_gallery_dl_pinterest";
}
function ensureCacheDir(fso, shell) {
    var d = getCacheDir(shell);
    if (!fso.FolderExists(d)) fso.CreateFolder(d);
}
function boardCachePath(shell, slug) {
    return getCacheDir(shell) + "\\boards_" + slug + ".txt";
}
function settingsPath(shell) {
    return getCacheDir(shell) + "\\settings.ini";
}

// Extract username slug from a profile URL for use in cache filenames.
function slugFromUrl(url) {
    var s = trimStr(url).replace(/\/+$/, "").split("/");
    return (s[s.length - 1] || "user").replace(/[^\w\-]/g, "_");
}

// ── settings ─────────────────────────────────────────────────────────────────

function loadSettings(shell, fso) {
    var out = { profile: 0, boardUrl: "", cookies: 1, keepps: 0 };
    try {
        var p = settingsPath(shell);
        if (!fso.FileExists(p)) return out;
        var f = fso.OpenTextFile(p, 1, false);
        var lines = f.ReadAll().split("\n");
        f.Close();
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i].replace(/\r$/, "");
            var eq = line.indexOf("=");
            if (eq < 1) continue;
            var k = line.substring(0, eq), v = line.substring(eq + 1);
            if      (k === "profile")  out.profile  = parseInt(v, 10) || 0;
            else if (k === "boardUrl") out.boardUrl = v;
            else if (k === "cookies")  out.cookies  = (parseInt(v, 10) === 0) ? 0 : 1;
            else if (k === "keepps")   out.keepps   = (parseInt(v, 10) === 0) ? 0 : 1;
        }
    } catch (e) { /* use defaults */ }
    if (out.profile < 0 || out.profile > 2) out.profile = 0;
    return out;
}

function saveSettings(shell, fso, profile, boardUrl, cookies, keepps) {
    try {
        ensureCacheDir(fso, shell);
        var f = fso.CreateTextFile(settingsPath(shell), true, false);
        f.WriteLine("profile=" + profile);
        f.WriteLine("boardUrl=" + String(boardUrl || "").replace(/[\r\n]/g, ""));
        f.WriteLine("cookies=" + cookies);
        f.WriteLine("keepps=" + keepps);
        f.Close();
    } catch (e) { /* ignore */ }
}

// ── board cache ───────────────────────────────────────────────────────────────

function isBoardUrl(line) {
    line = trimStr(line);
    if (!line || line.charAt(0) === "#") return false;
    if (line.toLowerCase().indexOf("pinterest.") < 0) return false;
    if (line.indexOf("/pin/") >= 0) return false;
    var m = line.match(/pinterest\.[^/]+\/[^/]+\/([^/?#]+)\/?$/);
    if (!m) return false;
    var seg = m[1].toLowerCase();
    return seg !== "pins" && seg !== "_created" && seg !== "_saved" && seg !== "search" && seg !== "ideas";
}

function boardLabel(url) {
    var m = trimStr(url).match(/\/([^/?#]+)\/?$/);
    if (!m) return url;
    try { return decodeURIComponent(m[1].replace(/\+/g, " ")); } catch (e) { return m[1]; }
}

function readBoardCache(fso, path) {
    if (!fso.FileExists(path)) return [];
    var f = fso.OpenTextFile(path, 1, false);
    var lines = f.ReadAll().split(/\r?\n/);
    f.Close();
    var out = [], seen = {};
    for (var i = 0; i < lines.length; i++) {
        var line = trimStr(lines[i]);
        if (isBoardUrl(line) && !seen[line]) { seen[line] = true; out.push(line); }
    }
    return out;
}

// Runs gallery-dl -g on the profile URL, captures output, filters board URLs,
// saves to cache. Returns count saved (0 on failure).
function refreshBoardCache(shell, fso, profileUrl, useCookies) {
    if (!fso.FileExists(GALLERY_DL_EXE)) {
        showError(shell, "gallery-dl.exe not found at:\n" + GALLERY_DL_EXE, "gallery-dl Error");
        return 0;
    }

    var slug    = slugFromUrl(profileUrl);
    var tempOut = shell.ExpandEnvironmentStrings("%TEMP%") + "\\gdl-boards-" + slug + ".txt";
    var tempPs1 = shell.ExpandEnvironmentStrings("%TEMP%") + "\\gdl-list-boards.ps1";

    try {
        var ps1 = fso.CreateTextFile(tempPs1, true, false);
        ps1.WriteLine("$ErrorActionPreference = 'Continue'");
        ps1.WriteLine("$gdArgs = @('-g'");
        if (useCookies) {
            ps1.WriteLine("    '--cookies-from-browser', 'firefox'");
        }
        ps1.WriteLine("    '" + psQuote(profileUrl) + "')");
        ps1.WriteLine("& '" + psQuote(GALLERY_DL_EXE) + "' @gdArgs 2>&1 |");
        ps1.WriteLine("    ForEach-Object { $_.ToString() } |");
        ps1.WriteLine("    Set-Content -LiteralPath '" + psQuote(tempOut) + "' -Encoding utf8");
        ps1.Close();
    } catch (e) {
        showError(shell, "Failed to write list script: " + e.message, "gallery-dl Error");
        return 0;
    }

    shell.Run('powershell -NoProfile -ExecutionPolicy Bypass -File "' + tempPs1 + '"', 0, true);

    var boards = readBoardCache(fso, tempOut);
    if (boards.length === 0) {
        showError(shell, "No board URLs found for:\n" + profileUrl + "\n\nCheck Firefox cookies, network, and gallery-dl.\nPrevious cache was not overwritten.", "gallery-dl");
        return 0;
    }

    ensureCacheDir(fso, shell);
    var cf = fso.CreateTextFile(boardCachePath(shell, slug), true, false);
    for (var i = 0; i < boards.length; i++) cf.WriteLine(boards[i]);
    cf.Close();

    try { if (fso.FileExists(tempOut)) fso.DeleteFile(tempOut); } catch (e2) { /* ignore */ }
    return boards.length;
}

// ── dialog helpers ────────────────────────────────────────────────────────────

function profileIdx(dlg) {
    if (dlg.control("pin_radio2").value) return 1;
    if (dlg.control("pin_radio3").value) return 2;
    return 0;
}

// Returns [{label, value}, …] for the board combo for a given profile.
function boardRows(profileIdx, shell, fso) {
    var rows = [{ label: "(All boards — entire profile)", value: PIN_URLS[profileIdx] }];
    var boards = readBoardCache(fso, boardCachePath(shell, slugFromUrl(PIN_URLS[profileIdx])));
    for (var i = 0; i < boards.length; i++) {
        rows.push({ label: boardLabel(boards[i]), value: boards[i] });
    }
    return rows;
}

// DOpus combo .value returns the selected index, not the item data.
function comboUrl(raw, rows) {
    var n = parseInt(String(raw), 10);
    if (!isNaN(n) && n >= 0 && n < rows.length) return rows[n].value;
    return rows[0].value;
}

function populateCombo(dlg, pidx, shell, fso, selectUrl) {
    var combo = dlg.control("board_combo");
    var rows = boardRows(pidx, shell, fso);
    combo.RemoveItem(-1);
    for (var r = 0; r < rows.length; r++) combo.AddItem(rows[r].label, rows[r].value);

    var sel = 0;
    if (selectUrl) {
        for (var j = 0; j < rows.length; j++) {
            if (trimStr(rows[j].value) === trimStr(selectUrl)) { sel = j; break; }
        }
    }
    combo.SelectItem(sel);
}

// ── main ──────────────────────────────────────────────────────────────────────

function OnClick(clickData) {
    var shell = new ActiveXObject("WScript.Shell");
    var fso   = new ActiveXObject("Scripting.FileSystemObject");

    // Build dialog
    var dlg = DOpus.dlg;
    dlg.window   = clickData.func.sourcetab;
    dlg.template = "GalleryDlPinterestDlg";
    dlg.detach   = true;
    dlg.Create();

    var saved = loadSettings(shell, fso);
    if (saved.profile === 1) dlg.control("pin_radio2").value = true;
    else if (saved.profile === 2) dlg.control("pin_radio3").value = true;
    else dlg.control("pin_radio1").value = true;
    dlg.control("cookies_check").value = saved.cookies === 1;
    dlg.control("keepps_check").value  = saved.keepps === 1;
    populateCombo(dlg, saved.profile, shell, fso, saved.boardUrl);

    dlg.Show();

    // Message loop
    var dialogResult = 0;
    while (true) {
        var msg = dlg.GetMsg();
        if (!msg.result) { dialogResult = dlg.result; break; }

        if (msg.event === "click" && msg.control === "refresh_btn") {
            var pidx = profileIdx(dlg);
            var nb = refreshBoardCache(shell, fso, PIN_URLS[pidx], dlg.control("cookies_check").value);
            if (nb > 0) {
                populateCombo(dlg, pidx, shell, fso);
                showInfo(shell, "Saved " + nb + " board(s).\n\n" + boardCachePath(shell, slugFromUrl(PIN_URLS[pidx])), "gallery-dl");
            }
        } else if (msg.event === "click" && msg.control.indexOf("pin_radio") === 0) {
            populateCombo(dlg, profileIdx(dlg), shell, fso);
        }
    }

    // Read final values
    var idx          = profileIdx(dlg);
    var rows         = boardRows(idx, shell, fso);
    var url          = comboUrl(dlg.control("board_combo").value, rows);
    var useFirefox   = dlg.control("cookies_check").value;
    var keepPsOpen   = dlg.control("keepps_check").value;

    saveSettings(shell, fso, idx, url, useFirefox ? 1 : 0, keepPsOpen ? 1 : 0);

    if (dialogResult == "0" || dialogResult == "2") {
        DOpus.Output("gallery-dl Pinterest: cancelled");
        return;
    }

    if (!fso.FileExists(GALLERY_DL_EXE)) {
        showError(shell, "gallery-dl.exe not found at:\n" + GALLERY_DL_EXE, "gallery-dl Error");
        return;
    }

    // Write and run download script.
    // Pass -d PINTEREST_OUTPUT_ROOT only; gallery-dl appends pinterest\username\board\ itself.
    var tempPs1 = shell.ExpandEnvironmentStrings("%TEMP%") + "\\gallery-dl-pinterest-run.ps1";
    try {
        var ps1 = fso.CreateTextFile(tempPs1, true, false);
        ps1.WriteLine("$gdArgs = @('-d', '" + psQuote(PINTEREST_OUTPUT_ROOT) + "'");
        if (useFirefox) {
            ps1.WriteLine("    '--cookies-from-browser', 'firefox'");
        }
        ps1.WriteLine("    '" + psQuote(url) + "')");
        ps1.WriteLine("& '" + psQuote(GALLERY_DL_EXE) + "' @gdArgs");
        ps1.Close();
    } catch (e) {
        showError(shell, "Failed to write temp script: " + e.message, "gallery-dl Error");
        return;
    }

    DOpus.Output("gallery-dl Pinterest | Profile: " + PROFILE_LABELS[idx]);
    DOpus.Output("gallery-dl Pinterest | URL: " + url);
    DOpus.Output("gallery-dl Pinterest | Root: " + PINTEREST_OUTPUT_ROOT);
    DOpus.Output("gallery-dl Pinterest | Firefox cookies: " + (useFirefox ? "on" : "off"));

    var psCmd = keepPsOpen
        ? 'powershell -NoExit -ExecutionPolicy Bypass -File "' + tempPs1 + '"'
        : 'powershell -ExecutionPolicy Bypass -File "' + tempPs1 + '"';
    shell.Run(psCmd, 1, false);
}
