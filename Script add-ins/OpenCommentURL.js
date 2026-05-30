// Ctrl+double-click a file or folder to open the URL in its Comment metadata

function OnInit(initData) {
    initData.name = "Open Comment URL";
    initData.desc = "Ctrl+double-click opens the URL stored in a file's Comment metadata.";
    initData.version = "1.1";
    initData.copyright = "";
    initData.default_enable = true;
    return false;
}

function trimStr(s) {
    return String(s).replace(/^\s+|\s+$/g, "");
}

function metaComment(meta, group, field) {
    try {
        var obj = meta[group];
        if (!obj) return "";
        var val = obj[field];
        if (val == null || val === undefined) return "";
        return trimStr(val);
    } catch (e) { }
    return "";
}

function getFileComment(item) {
    var path = String(item.realpath || item.path || "");
    var meta = null;
    try {
        if (path)
            meta = DOpus.FSUtil.GetMetadata(path);
    } catch (e) { }
    if (!meta) {
        try { meta = item.metadata; } catch (e2) { }
    }

    if (meta && String(meta + "") !== "none") {
        // DOpus extended comment (NTFS ADS, descript.ion, or unified view)
        var comment = metaComment(meta, "other", "usercomment");
        if (comment) return comment;

        // DOpus metadata editor stores Comment inside the file when the format supports it
        comment = metaComment(meta, "doc", "comments");
        if (comment) return comment;
        comment = metaComment(meta, "image", "imagedesc");
        if (comment) return comment;
        comment = metaComment(meta, "audio", "mp3comment");
        if (comment) return comment;
    }

    // Windows Properties > Details > Comments (System.Comment shell property)
    try {
        var shellComment = item.shellprop("System.Comment");
        if (shellComment)
            return trimStr(shellComment);
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
