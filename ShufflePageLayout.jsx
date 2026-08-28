(function () {
    if (!app.documents.length) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;

    // 1. Enable shuffling on every spread
    for (var i = 0; i < doc.spreads.length; i++) {
        try {
            doc.spreads[i].allowPageShuffle = true;
        } catch (e) {}
    }

    // 2. Enable shuffling on every page (important!)
    for (var p = 0; p < doc.pages.length; p++) {
        try {
            doc.pages[p].allowPageShuffle = true;
        } catch (e) {}
    }

    // 3. Force InDesign to rebuild the entire spread structure
    // by temporarily adding and removing a page.
    var tmp = doc.pages.add(LocationOptions.AT_END);
    tmp.remove();

    // 4. Optional: force spreads to recompute again
    // by reordering the last page
    try {
        var last = doc.pages[-1];
        last.move(LocationOptions.AT_END);
    } catch (e) {}

    alert("All spreads normalized. Pages now shuffle correctly.");
})();