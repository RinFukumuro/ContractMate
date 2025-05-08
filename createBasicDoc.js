// Create Basic Word Document for Testing
// This script creates a simple document with basic styles

try {
    WScript.Echo("Creating basic Word document...");
    
    // Create Word application
    var word = new ActiveXObject("Word.Application");
    word.Visible = false; // Run in background
    
    // Create a new document
    var doc = word.Documents.Add();
    
    // Add title
    doc.Content.Text = "Sample Contract\r\n\r\n";
    
    // Add introduction
    doc.Content.InsertAfter("This Agreement is made between Company A and Company B.\r\n\r\n");
    
    // Add sections with simple formatting
    addHeading(doc, "Article 1: Purpose");
    doc.Content.InsertAfter("The purpose of this Agreement is to establish terms and conditions.\r\n\r\n");
    
    addHeading(doc, "Article 2: Scope of Work");
    doc.Content.InsertAfter("Company B shall perform the following services:\r\n");
    doc.Content.InsertAfter("1. Preparation of legal documents\r\n");
    doc.Content.InsertAfter("2. Review of contracts\r\n");
    doc.Content.InsertAfter("3. Other legal services\r\n\r\n");
    
    addHeading(doc, "Article 3: Term");
    doc.Content.InsertAfter("This Agreement shall be effective for one year.\r\n\r\n");
    
    addHeading(doc, "Article 4: Compensation");
    doc.Content.InsertAfter("Company A shall pay Company B a monthly fee of $5,000.\r\n\r\n");
    
    addHeading(doc, "Article 5: Confidentiality");
    doc.Content.InsertAfter("Both parties shall maintain confidentiality.\r\n\r\n");
    
    // Add signature block
    doc.Content.InsertAfter("IN WITNESS WHEREOF, the parties have executed this Agreement.\r\n\r\n");
    doc.Content.InsertAfter("Company A: ________________________\r\n\r\n");
    doc.Content.InsertAfter("Company B: ________________________\r\n");
    
    // Save the document
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var docPath = fso.GetAbsolutePathName("BasicContract.docx");
    doc.SaveAs(docPath);
    
    // Close document and quit Word
    doc.Close();
    word.Quit();
    
    WScript.Echo("Basic document created successfully: " + docPath);
    WScript.Echo("You can now test the navigator with: cscript stable_navigator.js BasicContract.docx");
    
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

// Function to add a heading (bold and larger font)
function addHeading(doc, title) {
    doc.Content.InsertAfter(title + "\r\n");
    var paraCount = doc.Paragraphs.Count;
    var range = doc.Paragraphs(paraCount).Range;
    range.Font.Bold = true;
    range.Font.Size = 14;
}