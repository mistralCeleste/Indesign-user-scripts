// Remove text frames that contain no real text (letters or digits),
// even if they are nested inside groups or share a story.

(function () {
    var doc = app.activeDocument;
    var removed = 0;
    var failed = [];

    // Detects whether a character is "real text"
    function isRealChar(ch) {
        var s = ch.contents;
        return /[\d\w\[\]\.\-\!\#\$\%\&\*\(\)\=\+]/.test(s);
    }

    // Check if this specific frame contains any real characters
    function frameHasRealText(tf) {
        try {
            var story = tf.parentStory;
            if (!story) return false;

            var gb = tf.geometricBounds; // [y1, x1, y2, x2]

            for (var i = 0; i < story.characters.length; i++) {
                var ch = story.characters[i];

                // Character baseline position
                var baseline = ch.baseline;

                // Horizontal position (approximate)
                var hpos = ch.horizontalOffset;

                // Check if the character lies inside the frame
                if (baseline >= gb[0] && baseline <= gb[2] &&
                    hpos >= gb[1] && hpos <= gb[3]) {

                    if (isRealChar(ch)) return true;
                }
            }
        } catch (e) {
            // If anything fails, assume no real text
        }
        return false;
    }

    // Iterate all page items, including nested ones
    var items = doc.allPageItems;
    for (var i = items.length - 1; i >= 0; i--) {
        var it = items[i];

        // TEXT FRAME CHECK
        if (it instanceof TextFrame) {
            try {
                if (!frameHasRealText(it)) {
                    if (it.locked) it.locked = false;
                    if (it.itemLayer && it.itemLayer.locked)
                        it.itemLayer.locked = false;

                    it.remove();
                    removed++;
                }
            } catch (e) {
                failed.push("Failed on frame index " + i + ": " + e);
            }
            continue;
        }
    }

    var msg = "Removed " + removed + " text frame(s) without real text.";
    if (failed.length > 0) {
        msg += "\n\nFailures:\n" + failed.join("\n");
    }

    // alert(msg);
})();