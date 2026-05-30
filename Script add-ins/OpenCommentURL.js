// Ctrl+double-click a file or folder to open the URL in its Comment metadata

function OnInit(initData) {
    initData.name = "Open Comment URL";
    initData.desc = "Ctrl+double-click opens the URL stored in a file's Comment metadata.";
    initData.version = "1.0";
    initData.copyright = "";
    initData.default_enable = true;
    return false;
}

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function getFileComment(item) {
    try {
        var meta = item.metadata;
        if (!meta) return "";
        if (String(meta + "") === "none") return "";
        var other = meta.other;
        if (other && other.usercomment)
            return trimStr(other.usercomment);
    } catch (e) { }
    return "";
}

function extractUrl(text) {
    if (!text) return "";
    var t = trimStr(text);
    if (/^https?:\/\//i.test(t) && t.indexOf(" ") < 0)
        return t;
    var m = t.match(/https?:\/\/[^\s<>"')\]]+/i);
    return m ? m[0] : "";
}

function openCommentUrl(item) {
    var name = String(item.name || item.realpath || "");
    var comment = getFileComment(item);
    if (!comment) {
        DOpus.Output("Comment URL: no comment on " + name);
        return;
    }

    var url = extractUrl(comment);
    if (!url) {
        DOpus.Output("Comment URL: no URL in comment on " + name);
        return;
    }

    var shell = new ActiveXObject("WScript.Shell");
    shell.Run(url, 1, false);
    DOpus.Output("Opened in browser: " + url);
}

function OnDoubleClick(data) {
    if (String(data.qualifiers).indexOf("ctrl") < 0)
        return false;
    if (String(data.mouse) !== "left")
        return false;
    if (data.early)
        return false;

    openCommentUrl(data.item);
    return true;
}
