try {
    if (app.documents.length === 0) {
        throw new Error("No document is open.");
    }

    var doc = app.activeDocument;

    function getPS(name) {
        var ps = doc.paragraphStyles.itemByName(name);
        return ps.isValid ? ps : null;
    }

    function hideMarkupCharacters(textObj, count) {
        var cs = doc.characterStyles.itemByName("Markup");
        if (!cs.isValid) return;

        for (var i = 0; i < count && i < textObj.characters.length; i++) {
            textObj.characters[i].appliedCharacterStyle = cs;
        }
    }

    function forceAutosize(tf) {
        var prefs = tf.textFramePreferences;
        var original = prefs.autoSizingType;
        prefs.autoSizingType = AutoSizingTypeEnum.OFF;
        prefs.autoSizingType = original;
    }

    // Load styles
    var H1 = getPS("Heading 1");
    var H2 = getPS("Heading 2");
    var H3 = getPS("Heading 3");
    var A1 = getPS("Action 1");
    var A2 = getPS("Action 2");
    var ListBullet = getPS("List Bullet");

    // Counters
    var countBR = 0, countH1 = 0, countH2 = 0, countH3 = 0, countA1 = 0, countA2 = 0, countListBullet = 0;

    // Reset preferences
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;

    // 1. Convert <br> to paragraph breaks
    app.findTextPreferences.findWhat = "<br>";
    app.changeTextPreferences.changeTo = "^p";
    var brChanges = doc.changeText();
    countBR = brChanges.length;

    // Helper process function to avoid repeated null-pointer crashes
    function applyStyleToPattern(pattern, styleObj, hideCount) {
        if (!styleObj) return 0; // Skip safely if style is missing from document
        app.findGrepPreferences = NothingEnum.nothing;
        app.findGrepPreferences.findWhat = pattern;
        var found = doc.findGrep();
        for (var i = 0; i < found.length; i++) {
            hideMarkupCharacters(found[i], hideCount);
            found[i].appliedParagraphStyle = styleObj;
        }
        return found.length;
    }

    // 2-7. Execute Replacements
    countH1 = applyStyleToPattern("^#\\s+(.+)$", H1, 1);
    countH2 = applyStyleToPattern("^##\\s+(.+)$", H2, 2);
    countH3 = applyStyleToPattern("^###\\s+(.+)$", H3, 3);
    countA1 = applyStyleToPattern("^@\\s*(.+)$", A1, 1);
    countA2 = applyStyleToPattern("^@@\\s*(.+)$", A2, 2);
    countListBullet = applyStyleToPattern("^-\\s*(.+)$", ListBullet, 1);

    // 8. Auto-size text frames
    var frames = doc.textFrames;
    for (var k = 0; k < frames.length; k++) {
        forceAutosize(frames[k]);
    }

    // Reset preferences
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;

    // Remove or comment out alert() when called within a master batch loop
    // alert("Markdown Processed");

} catch (e) {
    // Re-throw so master script can capture it properly in doScript
    throw e; 
}
