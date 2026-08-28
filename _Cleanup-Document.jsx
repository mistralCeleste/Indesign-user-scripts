(function () {
    var base = File($.fileName).parent;
    var ranScripts = [];

    var steps = [
        "FindChange_MarkDown.jsx",
        "FindRemoveEmptyTextFrames.jsx", // Removed redundant 2nd run
        "FixPageHeadersInnerOuter.jsx"
    ];

    for (var i = 0; i < steps.length; i++) {
        var scriptFile = File(base + "/" + steps[i]);

        if (scriptFile.exists) {
            try {
                // Execute the child script
                app.doScript(
                    scriptFile,
                    ScriptLanguage.JAVASCRIPT,
                    undefined,
                    UndoModes.ENTIRE_SCRIPT,
                    "Run " + steps[i]
                );

                ranScripts.push(steps[i]);

            } catch (e) {
                alert("Execution error in: " + steps[i] + "\nLine: " + e.line + "\n" + e.message);
                break; // Stop execution on actual failure to prevent cascading bugs
            }
        } else {
            alert("Missing script file:\n" + scriptFile.fsName);
        }
    }

    alert("Batch Processing Complete!\n\nScripts Ran:\n- " + ranScripts.join("\n- "));
})();
