// ExportGREPStyles.jsx
// A script to export InDesign GREP styles to a file

#target indesign

// Check if a document is open
if (app.documents.length === 0) {
    alert("Please open a document before running this script.");
} else {
    main();
}

function main() {
    var doc = app.activeDocument;
    
    // Get all paragraph styles that have GREP styles
    var paragraphStyles = doc.allParagraphStyles;
    var grepStylesData = [];
    
    for (var i = 0; i < paragraphStyles.length; i++) {
        var style = paragraphStyles[i];
        
        // Skip the [No paragraph style]
        if (style.name === "[No paragraph style]") continue;
        
        // Check if the paragraph style has GREP styles
        if (style.nestedGrepStyles.length > 0) {
            for (var j = 0; j < style.nestedGrepStyles.length; j++) {
                var grepStyle = style.nestedGrepStyles[j];
                
                // Create a simple data string for each GREP style
                var styleEntry = style.name + "\t" + 
                                 grepStyle.appliedCharacterStyle.name + "\t" + 
                                 grepStyle.grepExpression.replace(/\t/g, "\\t").replace(/\n/g, "\\n");
                
                // Add optional properties if they exist
                var includeFootnotes = "N/A";
                try {
                    if (grepStyle.includeFootnotes !== undefined) {
                        includeFootnotes = grepStyle.includeFootnotes ? "Yes" : "No";
                    }
                } catch(e) {
                    // Property not available in this InDesign version - skip it
                }
                
                var fontStyle = "N/A";
                try {
                    if (grepStyle.fontStyle !== undefined) {
                        fontStyle = grepStyle.fontStyle;
                    }
                } catch(e) {
                    // Property not available in this InDesign version - skip it
                }
                
                styleEntry += "\t" + includeFootnotes + "\t" + fontStyle;
                grepStylesData.push(styleEntry);
            }
        }
    }
    
    // Save the data to a file
    var file = File.saveDialog("Save GREP Styles as", "Text files:*.txt");
    
    if (file) {
        file.encoding = "UTF-8";
        file.open("w");
        
        // Write header row
        file.writeln("Paragraph Style\tCharacter Style\tGREP Expression\tInclude Footnotes\tFont Style");
        
        // Write each style row
        for (var i = 0; i < grepStylesData.length; i++) {
            file.writeln(grepStylesData[i]);
        }
        
        file.close();
        
        alert("Exported " + grepStylesData.length + " GREP styles to " + file.fsName);
    }
}