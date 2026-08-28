#target indesign

(function () {

    // --- Prompt user for export mode ---
    var dlg = new Window("dialog", "Export with fullId");
    dlg.orientation = "column";

    dlg.add("statictext", undefined, "Export as:");
    var rbPages = dlg.add("radiobutton", undefined, "Pages");
    var rbSpreads = dlg.add("radiobutton", undefined, "Spreads");
    rbPages.value = true;

    dlg.add("statictext", undefined, "Enter PNG export resolution (DPI):");
    var resInput = dlg.add("edittext", undefined, "300");
    resInput.characters = 6;

    dlg.add("statictext", undefined, "Suffix:");
    var suffixInput = dlg.add("edittext", undefined, "");
    suffixInput.characters = 18;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    btnGroup.add("button", undefined, "OK");
    btnGroup.add("button", undefined, "Cancel");

    if (dlg.show() !== 1) return;

    var exportSpreads = rbSpreads.value;
    var resolution = parseInt(resInput.text, 10);
    var suffix = suffixInput.text;
    if (isNaN(resolution) || resolution < 10) {
        alert("Invalid resolution.");
        return;
    }


    // --- Ask for output folder using a fast file dialog ---
    var tempFile = File.saveDialog("Choose export location (any filename)");
    if (!tempFile) return;
    var outFolder = tempFile.parent;

    var doc = app.activeDocument;

    // --- Helper: find fullId frame on a page ---
    function getFullIdFromPage(page) {
        try {
            var frames = page.textFrames;
            for (var i = 0; i < frames.length; i++) {
                if (frames[i].label === "fullId") {
                    var txt = frames[i].contents.replace(/^\s+|\s+$/g, "");
                    if (txt.length > 0) return txt;
                }
            }
        } catch (e) {}
        return null;
    }

    function expandSuffix(suffix, info) {
        if (!suffix) return "";

        var out = suffix;
        out = out.replace(/\{fullId\}/gi, info.fullId || "");
        out = out.replace(/\{page\}/gi, info.pageNum || "");
        out = out.replace(/\{spread\}/gi, info.spreadNum || "");
        out = out.replace(/\{name\}/gi, info.name || "");
        out = out.replace(/-0\^#/g, "-" + pad3(info.pageNum || info.spreadNum || 0));

        return out;
    }

    // --- Helper: export a page ---
    function exportPage(page, filename) {
        var file = File(outFolder + "/" + filename + ".png");
        //app.pngExportPreferences.pngQuality = PNGQualityEnum.PNG_24;
        app.pngExportPreferences.transparentBackground = true;
        app.pngExportPreferences.antiAlias = true;
        app.pngExportPreferences.exportResolution = resolution;
        app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.RGB;
        app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
        app.pngExportPreferences.pageString = page.name;
        doc.exportFile(ExportFormat.PNG_FORMAT, file, false);
    }

    function exportSpread(spread, filename) {
        var file = File(outFolder + "/" + filename + ".png");

        //app.pngExportPreferences.pngQuality = PNGQualityEnum.PNG_24;
        app.pngExportPreferences.transparentBackground = true;
        app.pngExportPreferences.antiAlias = true;
        app.pngExportPreferences.exportResolution = resolution;
        app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.RGB;

        // Important: enable spread export
        app.pngExportPreferences.exportingSpread = true;
        app.pngExportPreferences.pngExportRange = PNGExportRangeEnum.EXPORT_RANGE;
        app.pngExportPreferences.pageString = spread.pages[0].name + "-" + spread.pages[-1].name;

        doc.exportFile(ExportFormat.PNG_FORMAT, file, false);

        // Reset spread mode so page exports later don’t break
        app.pngExportPreferences.exportingSpread = false;
    }

    function pad3(n) {
        n = parseInt(n, 10);
        if (n < 10) return "00" + n;
        if (n < 100) return "0" + n;
        return "" + n;
    }

    // --- Export Pages Mode ---
    if (!exportSpreads) {
        for (var i = 0; i < doc.pages.length; i++) {
            var page = doc.pages[i];

            var id = getFullIdFromPage(page);
            if (!id) {
                var num = pad3(i + 1);
                id = "page_" + num;
            }

            var filename = id + suffix;
            exportPage(page, filename);
        }

        alert("Export complete (Pages).");
        return;
    }

    // --- Export Spreads Mode ---
    for (var s = 0; s < doc.spreads.length; s++) {
        var spread = doc.spreads[s];

        var id = null;

        // Search all pages in the spread for fullId
        for (var p = 0; p < spread.pages.length; p++) {
            id = getFullIdFromPage(spread.pages[p]);
            if (id) break;
        }

        if (!id) {
            var num = pad3(s + 1);
            id = "spread_" + num;
        }

        var filename = id + suffix;
        exportSpread(spread, filename);
    }

    alert("Export complete (Spreads).");

})();