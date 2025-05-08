// Stable Word Document Navigator
// Windows Script Host (WSH) + ActiveXObject

// Main function to avoid global variables
function main() {
    try {
        // Get command line arguments
        var args = WScript.Arguments;
        if (args.length < 1) {
            WScript.Echo("Usage: cscript stable_navigator.js [Word document path] [paragraph number (optional)]");
            WScript.Quit(1);
        }
        
        var filePath = args(0);
        var targetParagraph = args.length > 1 ? parseInt(args(1), 10) : 1;
        
        WScript.Echo("Opening document: " + filePath);
        
        // Create Word application
        var word = new ActiveXObject("Word.Application");
        word.Visible = true;
        
        // Open document
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var absPath = fso.GetAbsolutePathName(filePath);
        var doc = word.Documents.Open(absPath);
        
        WScript.Echo("Document opened successfully");
        
        // Get paragraph count
        var paragraphCount = doc.Paragraphs.Count;
        WScript.Echo("Total paragraphs: " + paragraphCount);
        
        // Analyze document structure
        analyzeDocumentStructure(doc);
        
        // Navigate to specified paragraph if valid
        if (targetParagraph > 0 && targetParagraph <= paragraphCount) {
            navigateToParagraph(doc, targetParagraph);
            WScript.Echo("Navigated to paragraph " + targetParagraph);
        } else if (targetParagraph > 0) {
            WScript.Echo("Invalid paragraph number: " + targetParagraph);
            WScript.Echo("Valid range is 1 to " + paragraphCount);
        }
        
        // Keep Word open for user interaction
        WScript.Echo("\nWord document is now open. Close this script with Ctrl+C when finished.");
        WScript.Echo("DO NOT close Word manually before closing this script.");
        
        // Wait for user to finish
        while (true) {
            WScript.Sleep(1000); // Sleep for 1 second
        }
        
    } catch (error) {
        WScript.Echo("Error: " + (error.description || error));
        
        // Try to quit Word if an error occurs
        try {
            if (word && word.Quit) {
                word.Quit(0); // 0 = Don't save changes
            }
        } catch (e) {
            // Ignore errors when trying to quit Word
        }
    }
}

// Function to analyze document structure
function analyzeDocumentStructure(doc) {
    try {
        WScript.Echo("\nDocument Structure Analysis:");
        
        // Get heading styles and outline levels
        var headings = [];
        var paragraphCount = doc.Paragraphs.Count;
        
        for (var i = 1; i <= paragraphCount; i++) {
            var para = doc.Paragraphs(i);
            var paraText = para.Range.Text;
            
            // Clean up text (remove line breaks)
            if (paraText) {
                paraText = paraText.replace(/[\r\n]+/g, " ");
            } else {
                paraText = "";
            }
            
            var style = para.Range.Style.NameLocal;
            var outlineLevel = para.OutlineLevel;
            
            // Store paragraph info
            var paraInfo = {
                index: i,
                text: paraText,
                style: style,
                outlineLevel: outlineLevel
            };
            
            // Display paragraph info
            WScript.Echo("Paragraph " + i + ":");
            WScript.Echo("  Style: " + style);
            WScript.Echo("  Outline Level: " + outlineLevel);
            WScript.Echo("  Text: " + paraText);
            
            // Identify headings (outline level < 9 indicates a heading in Word)
            if (outlineLevel < 9) {
                headings.push(paraInfo);
            }
        }
        
        // Display headings summary
        if (headings.length > 0) {
            WScript.Echo("\nDocument Headings:");
            for (var j = 0; j < headings.length; j++) {
                var heading = headings[j];
                var indent = "";
                for (var k = 1; k < heading.outlineLevel; k++) {
                    indent += "  ";
                }
                WScript.Echo(indent + "Paragraph " + heading.index + ": " + heading.text);
            }
        } else {
            WScript.Echo("\nNo headings found in document.");
        }
        
        return headings;
    } catch (error) {
        WScript.Echo("Error analyzing document structure: " + (error.description || error));
        return [];
    }
}

// Function to navigate to a specific paragraph
function navigateToParagraph(doc, paragraphNumber) {
    try {
        var para = doc.Paragraphs(paragraphNumber);
        para.Range.Select();
        return true;
    } catch (error) {
        WScript.Echo("Error navigating to paragraph: " + (error.description || error));
        return false;
    }
}

// Start the main function
main();