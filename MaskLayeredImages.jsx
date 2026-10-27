// Replace image frames labeled "Layered Image: <suffix>"
// with "<baseName><suffix><extension>"

(function () {

    function logln(msg) { log += msg + "\n"; }
    function trimString(str) {
        return str.replace(/^\s+|\s+$/g, "");
    }

    var doc = app.activeDocument;
    var links = doc.links;
    var log = "";

    logln("Active document: " + doc.name);
    logln("Total links: " + links.length);
    logln("--------------------------------------------");

    for (var i = 0; i < links.length; i++) {

        var link = links[i];
        var graphic = link.parent;
        var frame = graphic.parent;

        if (!frame || !frame.label) {
            continue;
        }

        var label = frame.label;

        logln("Checking link #" + i + ": " + link.name);
        logln("  -> Frame type: " + frame.constructor.name);
        logln("  -> Label: '" + label + "'");

        var prefix = "Layered Image: ";
        if (label.indexOf(prefix) !== 0) {
            logln("  -> Not a Layered Image item. Skipping.");
            continue;
        }

        // Extract suffix from label (without using String.trim())
        var rawSuffix = label.substring(prefix.length);
        var suffix = trimString(rawSuffix);
        logln("  -> Suffix from label: '" + suffix + "'");

        var originalPath = link.filePath;
        logln("  -> Original link path: " + originalPath);

        var fileObj = File(originalPath);
        var fileName = fileObj.name;
        var dotIndex = fileName.lastIndexOf(".");

        if (dotIndex < 0) {
            logln("  -> ERROR: File has no extension. Skipping.");
            continue;
        }

        var baseName = fileName.substring(0, dotIndex);
        var extension = fileName.substring(dotIndex);

        logln("  -> Base name: " + baseName);
        logln("  -> Extension: " + extension);

        // Skip if already suffixed
        if (baseName.indexOf(suffix, baseName.length - suffix.length) !== -1) {
            logln("  -> Already using suffix. Skipping.");
            continue;
        }

        var newName = baseName + suffix + extension;
        var newPath = fileObj.path + "/" + newName;

        logln("  -> New filename: " + newName);
        logln("  -> New path: " + newPath);

        try {
            frame.place(File(newPath));
            logln("  -> Placed new image into frame.");
        } catch (e) {
            logln("  -> ERROR placing new image: " + e);
        }

        logln("--------------------------------------------");
    }

    // alert(log);

})();
