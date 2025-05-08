// Simple Text File Analyzer
// This script reads a text file and displays line numbers and content

try {
    // Get command line arguments
    var args = WScript.Arguments;
    if (args.length < 1) {
        WScript.Echo("Usage: cscript text_analyzer.js [text file path]");
        WScript.Quit(1);
    }
    
    var filePath = args(0);
    WScript.Echo("Analyzing file: " + filePath);
    
    // Create FileSystemObject
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    
    // Check if file exists
    if (!fso.FileExists(filePath)) {
        WScript.Echo("Error: File does not exist: " + filePath);
        WScript.Quit(1);
    }
    
    // Open the file for reading
    var file = fso.OpenTextFile(filePath, 1); // 1 = ForReading
    
    // Read and display each line with line number
    var lineNumber = 0;
    WScript.Echo("File content with line numbers:");
    
    while (!file.AtEndOfStream) {
        lineNumber++;
        var line = file.ReadLine();
        WScript.Echo(lineNumber + ": " + line);
    }
    
    // Close the file
    file.Close();
    
    WScript.Echo("Total lines: " + lineNumber);
    
    // Analyze document structure (simple version)
    WScript.Echo("\nDocument Structure Analysis:");
    
    // Reopen the file
    file = fso.OpenTextFile(filePath, 1);
    lineNumber = 0;
    
    // Arrays to store section headings
    var sections = [];
    
    while (!file.AtEndOfStream) {
        lineNumber++;
        var line = file.ReadLine();
        
        // Identify section headings (lines starting with ## in markdown)
        if (line.indexOf("## ") === 0) {
            sections.push({
                lineNumber: lineNumber,
                title: line.substring(3).trim()
            });
        }
    }
    
    // Close the file
    file.Close();
    
    // Display sections
    WScript.Echo("\nIdentified Sections:");
    for (var i = 0; i < sections.length; i++) {
        var section = sections[i];
        WScript.Echo("Line " + section.lineNumber + ": " + section.title);
    }
    
    WScript.Echo("\nTotal sections found: " + sections.length);
    
} catch (error) {
    WScript.Echo("Error: " + error.description);
}