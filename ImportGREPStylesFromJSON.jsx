// ImportGREPStyles.jsx
// A script to import GREP styles from a text file into InDesign

#target indesign

// Check if a document is open
if (app.documents.length === 0) {
    alert("Please open a document before running this script.");
} else {
    main();
}

function main() {
    var doc = app.activeDocument;
    
    // Ask for the file to import
    var file = File.openDialog("Select GREP Styles file to import", "Text files:*.txt");
    
    if (!file) return; // User canceled
    
    file.encoding = "UTF-8";
    file.open("r");
    
    // Read the header line and discard it
    var headerLine = file.readln();
    
    var importCount = 0;
    var updatedStyles = {};
    var missingStyles = [];
    var missingCharStyles = [];
    
    // Process the imported data
    app.doScript(function() {
        while (!file.eof) {
            var line = file.readln();
            
            // Skip empty lines (without using trim)
            if (line === "" || line.replace(/\s/g, "") === "") continue;
            
            // Parse the tab-delimited line
            var parts = line.split("\t");
            if (parts.length < 3) continue; // Need at least paragraph style, char style, and grep expression
            
            var paragraphStyleName = parts[0];
            var charStyleName = parts[1];
            var grepExpression = parts[2].replace(/\\t/g, "\t").replace(/\\n/g, "\n");
            
            // Find the paragraph style in the current document
            var paragraphStyle = findParagraphStyle(doc, paragraphStyleName);
            
            if (!paragraphStyle) {
                // Check if already in the missing styles array
                var alreadyMissing = false;
                for (var i = 0; i < missingStyles.length; i++) {
                    if (missingStyles[i] === paragraphStyleName) {
                        alreadyMissing = true;
                        break;
                    }
                }
                if (!alreadyMissing) {
                    missingStyles.push(paragraphStyleName);
                }
                continue;
            }
            
            // Find the character style
            var charStyle = findCharacterStyle(doc, charStyleName);
            
            if (!charStyle) {
                // Check if already in the missing char styles array
                var alreadyMissingChar = false;
                for (var i = 0; i < missingCharStyles.length; i++) {
                    if (missingCharStyles[i] === charStyleName) {
                        alreadyMissingChar = true;
                        break;
                    }
                }
                if (!alreadyMissingChar) {
                    missingCharStyles.push(charStyleName);
                }
                continue;
            }
            
            // Create a new GREP style
            var grepStyle = paragraphStyle.nestedGrepStyles.add();
            grepStyle.appliedCharacterStyle = charStyle;
            grepStyle.grepExpression = grepExpression;
            
            // Apply optional properties if supported in this version of InDesign
            if (parts.length > 3 && parts[3] !== "N/A") {
                try {
                    grepStyle.includeFootnotes = (parts[3] === "Yes");
                } catch(e) {
                    // Property not available in this InDesign version - skip it
                }
            }
            
            if (parts.length > 4 && parts[4] !== "N/A" && parts[4] !== "") {
                try {
                    grepStyle.fontStyle = parts[4];
                } catch(e) {
                    // Property not available in this InDesign version - skip it
                }
            }
            
            importCount++;
            
            // Track updated styles
            updatedStyles[paragraphStyleName] = true;
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Import GREP Styles");
    
    file.close();
    
    // Convert the updatedStyles object to an array
    var updatedStylesList = [];
    for (var styleName in updatedStyles) {
        if (updatedStyles.hasOwnProperty(styleName)) {
            updatedStylesList.push(styleName);
        }
    }
    
    // Show results
    var resultMessage = "Imported " + importCount + " GREP styles.\n\n";
    
    if (updatedStylesList.length > 0) {
        resultMessage += "Updated paragraph styles:\n" + updatedStylesList.join("\n") + "\n\n";
    }
    
    if (missingStyles.length > 0) {
        resultMessage += "Warning: The following paragraph styles were not found in this document:\n" + missingStyles.join("\n") + "\n\n";
    }
    
    if (missingCharStyles.length > 0) {
        resultMessage += "Warning: The following character styles were not found in this document:\n" + missingCharStyles.join("\n");
    }
    
    alert(resultMessage);
}

function findParagraphStyle(doc, name) {
    // First try to find an exact match
    try {
        return doc.paragraphStyles.itemByName(name);
    } catch (e) {
        // If that fails, search through all paragraph styles
        for (var i = 0; i < doc.allParagraphStyles.length; i++) {
            if (doc.allParagraphStyles[i].name === name) {
                return doc.allParagraphStyles[i];
            }
        }
    }
    return null;
}

function findCharacterStyle(doc, name) {
    // First try to find an exact match
    try {
        return doc.characterStyles.itemByName(name);
    } catch (e) {
        // If that fails, search through all character styles
        for (var i = 0; i < doc.allCharacterStyles.length; i++) {
            if (doc.allCharacterStyles[i].name === name) {
                return doc.allCharacterStyles[i];
            }
        }
    }
    return null;
}