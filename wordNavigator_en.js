// Word Document Navigator Tool
// Windows Script Host (WSH) + ActiveXObject

var WordNavigator = (function () {
    function WordNavigator() {
        try {
            this.wordApp = new ActiveXObject("Word.Application");
            this.wordApp.Visible = true;
        }
        catch (error) {
            WScript.Echo("Failed to start Word: " + error);
            WScript.Quit();
        }
    }
    
    // Open Word document
    WordNavigator.prototype.openDocument = function (filePath) {
        try {
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            var absPath = fso.GetAbsolutePathName(filePath);
            this.document = this.wordApp.Documents.Open(absPath);
            return true;
        }
        catch (error) {
            WScript.Echo("Could not open document: " + error);
            return false;
        }
    };
    
    // Get paragraph structure
    WordNavigator.prototype.getParagraphStructure = function () {
        if (!this.document) {
            WScript.Echo("No document is open.");
            return [];
        }
        
        var structure = [];
        var paragraphCount = this.document.Paragraphs.Count;
        
        for (var i = 1; i <= paragraphCount; i++) {
            var para = this.document.Paragraphs(i);
            var paraText = para.Range.Text;
            if (paraText) {
                paraText = paraText.replace(/\r/g, "").replace(/\n/g, "");
            } else {
                paraText = "";
            }
            var style = para.Range.Style.NameLocal;
            var outlineLevel = para.OutlineLevel;
            
            structure.push({
                index: i,
                text: paraText,
                style: style,
                outlineLevel: outlineLevel
            });
        }
        return structure;
    };
    
    // Navigate to specific paragraph
    WordNavigator.prototype.navigateToParagraph = function (paragraphNumber) {
        if (!this.document) {
            WScript.Echo("No document is open.");
            return false;
        }
        
        try {
            var totalParagraphs = this.document.Paragraphs.Count;
            if (paragraphNumber < 1 || paragraphNumber > totalParagraphs) {
                WScript.Echo("Invalid paragraph number: " + paragraphNumber);
                return false;
            }
            
            var para = this.document.Paragraphs(paragraphNumber);
            para.Range.Select();
            this.wordApp.Activate();
            return true;
        }
        catch (error) {
            WScript.Echo("Could not navigate to paragraph: " + error);
            return false;
        }
    };
    
    // Close document
    WordNavigator.prototype.closeDocument = function (save) {
        if (save === undefined) { save = false; }
        if (this.document) {
            if (save) {
                this.document.Save();
            }
            this.document.Close(save);
            this.document = null;
        }
    };
    
    // Quit Word application
    WordNavigator.prototype.quit = function () {
        if (this.wordApp) {
            this.wordApp.Quit();
            this.wordApp = null;
        }
    };
    
    return WordNavigator;
})();

// Main process
function main() {
    var args = WScript.Arguments;
    if (args.length < 2) {
        WScript.Echo("Usage: cscript wordNavigator.js [Word document path] [paragraph number]");
        WScript.Quit();
    }
    
    var filePath = args(0);
    var targetParagraph = parseInt(args(1), 10);
    
    var navigator = new WordNavigator();
    
    if (navigator.openDocument(filePath)) {
        var structure = navigator.getParagraphStructure();
        
        // Display document structure
        for (var i = 0; i < structure.length; i++) {
            var para = structure[i];
            WScript.Echo("Paragraph: " + para.index + ", Style: " + para.style + 
                       ", Outline Level: " + para.outlineLevel + ", Content: " + para.text);
        }
        
        // Navigate to specified paragraph
        if (!navigator.navigateToParagraph(targetParagraph)) {
            WScript.Echo("Failed to navigate to specified paragraph.");
        }
    }
}

// Execute script
main();