(function () {
    if (!app.documents.length) {
        alert("No document open.");
        return;
    }

    var doc = app.activeDocument;
    var removed = 0;

    // ============================
    // CONFIGURABLE NAMES
    // ============================
    var ODD_NAME  = "odd page header";
    var EVEN_NAME = "even page header";

    function norm(s) {
        if (!s) return "";
        s = String(s).toLowerCase();
        s = s.replace(/^\s+/, "").replace(/\s+$/, ""); // manual trim
        return s;
    }

    var ODD_NORM  = norm(ODD_NAME);
    var EVEN_NORM = norm(EVEN_NAME);

    var items = doc.allPageItems;

    for (var i = items.length - 1; i >= 0; i--) {
        var it = items[i];
        if (!it || !it.isValid) continue;

        var nm = it.name ? norm(it.name) : "";
        var isOddHeader  = (nm === ODD_NORM);
        var isEvenHeader = (nm === EVEN_NORM);

        if (!isOddHeader && !isEvenHeader) continue;

        var pg = it.parentPage;
        if (!pg || !pg.isValid) continue;

        var pageNum = pg.documentOffset + 1; // 1-based

        try {
            // Odd page → remove even header
            if (pageNum % 2 === 1 && isEvenHeader) {
                it.remove();
                removed++;
            }
            // Even page → remove odd header
            else if (pageNum % 2 === 0 && isOddHeader) {
                it.remove();
                removed++;
            }
        } catch (e) {
            // ignore if it vanishes mid-loop
        }
    }

    // alert("Removed " + removed + " inner header(s).");
})();