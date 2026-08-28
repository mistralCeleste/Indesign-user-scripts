// Find and Change Script for Adobe InDesign
try {
    // Ensure a document is open
    if (app.documents.length === 0) {
        throw new Error("No document is open. Please open a document and try again.");
    }

    // Define the text to find and replace
    var findText = "<br>"; // Replace with the text you want to find
    var replaceText = "^p"; // Replace with the text you want to replace it with

    // Access the active document
    var doc = app.activeDocument;

    // Clear previous find/change preferences
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;

    // Set the find/change preferences
    app.findTextPreferences.findWhat = findText;
    app.changeTextPreferences.changeTo = replaceText;

    // Perform the find and change operation
    var changesMade = doc.changeText();

    // Notify the user of the results
    alert(changesMade.length + " occurrences of '" + findText + "' were replaced with '" + replaceText + "'.");

    // Clear preferences after the operation
    app.findTextPreferences = NothingEnum.NOTHING;
    app.changeTextPreferences = NothingEnum.NOTHING;

} catch (e) {
    alert("Error: " + e.message);
}
