using System;
using Aspose.Words;
using Aspose.Words.Loading;

class InsertIntoProtectedDocument
{
    static void Main()
    {
        // Paths to the files.
        const string protectedDocPath = "ProtectedDocument.docx";   // Password‑protected source document
        const string sourceDocPath   = "SourceDocument.docx";      // Document to insert
        const string outputDocPath   = "ResultDocument.docx";      // Resulting document
        const string password        = "MyPassword";               // Password for the protected document

        // Load the password‑protected document using LoadOptions.
        LoadOptions loadOptions = new LoadOptions(password);
        Document targetDoc = new Document(protectedDocPath, loadOptions);

        // Remove any protection (Unprotect works without needing the password).
        targetDoc.Unprotect();

        // Load the document that will be inserted.
        Document sourceDoc = new Document(sourceDocPath);

        // Move the builder to the bookmark where the insertion should occur.
        DocumentBuilder builder = new DocumentBuilder(targetDoc);
        builder.MoveToBookmark("InsertHere"); // Replace with your bookmark name

        // Insert the source document at the bookmark position.
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the modified document.
        targetDoc.Save(outputDocPath);
    }
}
