// Simple Word Document Navigator

try {
    // Get command line arguments
    var args = WScript.Arguments;
    if (args.length < 2) {
        WScript.Echo("Usage: cscript simple_navigator.js [Word document path] [paragraph number]");
        WScript.Quit();
    }
    
    var filePath = args(0);
    var paragraphNumber = parseInt(args(1), 10);
    
    // Create Word application
    var word = new ActiveXObject("Word.Application");
    word.Visible = true;
    
    // Open document
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var absPath = fso.GetAbsolutePathName(filePath);
    var doc = word.Documents.Open(absPath);
    
    WScript.Echo("Document opened successfully: " + absPath);
    
    // Get paragraph count
    var paragraphCount = doc.Paragraphs.Count;
    WScript.Echo("Total paragraphs: " + paragraphCount);
    
    // Check if paragraph number is valid
    if (paragraphNumber < 1 || paragraphNumber > paragraphCount) {
        WScript.Echo("Invalid paragraph number: " + paragraphNumber);
        WScript.Echo("Valid range is 1 to " + paragraphCount);
    } else {
        // Navigate to paragraph
        var para = doc.Paragraphs(paragraphNumber);
        para.Range.Select();
        word.Activate();
        
        WScript.Echo("Navigated to paragraph " + paragraphNumber);
        
        // Display paragraph info
        WScript.Echo("Style: " + para.Range.Style.NameLocal);
        WScript.Echo("Text: " + para.Range.Text);
    }
} catch (error) {
    WScript.Echo("Error: " + error.description);
}