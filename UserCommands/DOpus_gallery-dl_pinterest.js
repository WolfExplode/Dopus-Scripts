// gallery-dl Pinterest: launches GalleryDl\GalleryDlTool.py (Python / PySide6) from this repo.
//
// The window picks a profile and board (board lists are cached; Refresh updates them)
// and streams gallery-dl's output while it scrapes.
// Cache + settings: %APPDATA%\DOpus_gallery_dl_pinterest\ (unchanged from the old dialog).

/** Optional full path to GalleryDlTool.py if auto-detect fails. */
var GALLERYDL_PY = "";

function quoteArg(s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
}

function resolveToolPy(fso) {
    if (GALLERYDL_PY && fso.FileExists(GALLERYDL_PY)) {
        return GALLERYDL_PY;
    }
    try {
        if (typeof Script !== "undefined" && Script && Script.file) {
            var sibling = fso.GetAbsolutePathName(
                fso.BuildPath(fso.GetParentFolderName(Script.file), "..\\GalleryDl\\GalleryDlTool.py")
            );
            if (fso.FileExists(sibling)) {
                return sibling;
            }
        }
    } catch (e) {}
    var fallback = "C:\\Users\\WXP\\Documents\\GitHub\\Dopus-Scripts\\GalleryDl\\GalleryDlTool.py";
    return fso.FileExists(fallback) ? fallback : "";
}

function OnClick(clickData) {
    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var toolPy = resolveToolPy(fso);
    if (!toolPy) {
        shell.Popup("GalleryDlTool.py not found.\n\nSet GALLERYDL_PY in DOpus_gallery-dl_pinterest.js.", 0,
            "gallery-dl", 16);
        return;
    }
    var exec = quoteArg("pythonw") + " " + quoteArg(toolPy);
    DOpus.Output("gallery-dl Pinterest: " + exec);
    shell.Run(exec, 0, false);
}
