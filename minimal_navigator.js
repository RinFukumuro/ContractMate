// Minimal Word Document Navigator

try {
    // Get command line arguments
    var args = WScript.Arguments;
    if (args.length < 2) {
        WScript.Echo("Usage: cscript minimal_navigator.js [Word document path] [paragraph number]");
        WScript.Quit(1);
    }
    
    var filePath = args(0);
    var paragraphNumber = parseInt(args(1), 10);
    
    WScript.Echo("Opening document: " + filePath);
    WScript.Echo("Target paragraph: " + paragraphNumber);
    
    // Create Word application
    var word = new ActiveXObject("Word.Application");
    word.Visible = true;
    
    // Open document
    var doc = word.Documents.Open(filePath);
    WScript.Echo("Document opened successfully");
    
    // Get paragraph count
    var paragraphCount = doc.Paragraphs.Count;
    WScript.Echo("Total paragraphs: " + paragraphCount);
    
    // Display first few paragraphs
    var displayCount = Math.min(paragraphCount, 10);
    WScript.Echo("Displaying first " + displayCount + " paragraphs:");
    
    for (var i = 1; i <= displayCount; i++) {
        var para = doc.Paragraphs(i);
        var text = para.Range.Text.replace(/[\r\n]+/g, " ");
        WScript.Echo(i + ": " + text);
    }
    
    // Navigate to specified paragraph if valid
    if (paragraphNumber > 0 && paragraphNumber <= paragraphCount) {
        var para = doc.Paragraphs(paragraphNumber);
        para.Range.Select();
        word.Activate();
        WScript.Echo("Navigated to paragraph " + paragraphNumber);
    } else {
        WScript.Echo("Invalid paragraph number: " + paragraphNumber);
        WScript.Echo("Valid range is 1 to " + paragraphCount);
    }
    
} catch (error) {
    WScript.Echo("Error: " + error.description);
}