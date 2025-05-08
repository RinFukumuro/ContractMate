// Create Sample Word Document for Testing
// This script creates a sample contract document in Word format

try {
    WScript.Echo("Creating sample Word document...");
    
    // Create Word application
    var word = new ActiveXObject("Word.Application");
    word.Visible = false; // Run in background
    
    // Create a new document
    var doc = word.Documents.Add();
    
    // Set document title
    var title = "Sample Contract";
    doc.Content.Text = title + "\r\n";
    doc.Paragraphs(1).Range.Style = "Title";
    
    // Add introduction
    doc.Content.InsertAfter("\r\nThis Agreement is made and entered into as of the date of last signature below, by and between Company A and Company B.\r\n\r\n");
    
    // Add sections with proper heading styles
    addSection(doc, "Article 1: Purpose", "Heading 1", 
        "The purpose of this Agreement is to establish the terms and conditions under which the parties will collaborate.");
    
    addSection(doc, "Article 2: Scope of Work", "Heading 1", 
        "Company B shall perform the following services for Company A:\r\n" +
        "1. Preparation of legal documents\r\n" +
        "2. Review of contracts\r\n" +
        "3. Other legal services as specified by Company A");
    
    addSection(doc, "2.1 Timeline", "Heading 2", 
        "All services shall be performed according to the timeline specified in Appendix A.");
    
    addSection(doc, "2.2 Quality Standards", "Heading 2", 
        "All services shall meet the quality standards specified in Appendix B.");
    
    addSection(doc, "Article 3: Term", "Heading 1", 
        "This Agreement shall be effective for a period of one year from the date of execution. " +
        "It shall automatically renew for successive one-year terms unless either party provides " +
        "written notice of non-renewal at least 30 days prior to the end of the current term.");
    
    addSection(doc, "Article 4: Compensation", "Heading 1", 
        "Company A shall pay Company B a monthly fee of $5,000 (excluding tax) for the services provided under this Agreement.");
    
    addSection(doc, "Article 5: Confidentiality", "Heading 1", 
        "Both parties shall maintain the confidentiality of all information disclosed by the other party " +
        "and shall not disclose such information to any third party without prior written consent.");
    
    addSection(doc, "Article 6: Termination", "Heading 1", 
        "Either party may terminate this Agreement if the other party breaches any material term " +
        "and fails to cure such breach within 30 days after receiving written notice thereof.");
    
    addSection(doc, "Article 7: Miscellaneous", "Heading 1", 
        "Any matters not addressed in this Agreement shall be resolved through good faith discussions between the parties.");
    
    // Add signature block
    doc.Content.InsertAfter("\r\n\r\nIN WITNESS WHEREOF, the parties have executed this Agreement as of the date first written above.\r\n\r\n");
    doc.Content.InsertAfter("Company A: ________________________    Date: ____________\r\n\r\n");
    doc.Content.InsertAfter("Company B: ________________________    Date: ____________\r\n");
    
    // Save the document
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var docPath = fso.GetAbsolutePathName("SampleContract.docx");
    doc.SaveAs(docPath);
    
    // Close document and quit Word
    doc.Close();
    word.Quit();
    
    WScript.Echo("Sample document created successfully: " + docPath);
    WScript.Echo("You can now test the navigator with: cscript stable_navigator.js SampleContract.docx");
    
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

// Function to add a section with heading and content
function addSection(doc, title, headingStyle, content) {
    doc.Content.InsertAfter(title + "\r\n");
    var paraCount = doc.Paragraphs.Count;
    doc.Paragraphs(paraCount).Range.Style = headingStyle;
    doc.Content.InsertAfter(content + "\r\n\r\n");
}