using System;
using Aspose.Words;
using Aspose.Words.Saving;

class JoinDocumentsWithBookmark
{
    static void Main()
    {
        // Path to the main document that contains the bookmark.
        string mainDocPath = @"C:\Docs\MainDocument.docx";

        // Path to the document that will be inserted.
        string insertDocPath = @"C:\Docs\DocumentToInsert.docx";

        // Path where the resulting document will be saved.
        string outputPath = @"C:\Docs\ResultDocument.docx";

        // Load the main document (lifecycle rule: load).
        Document mainDoc = new Document(mainDocPath);

        // Load the document to be inserted (lifecycle rule: load).
        Document insertDoc = new Document(insertDocPath);

        // Create a DocumentBuilder for the main document (lifecycle rule: create).
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Move the cursor to the bookmark named "InsertHere".
        // If the bookmark does not exist, MoveToBookmark will throw an exception.
        builder.MoveToBookmark("InsertHere");

        // Insert the document at the bookmark position.
        // KeepSourceFormatting preserves the formatting of the inserted document.
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document (lifecycle rule: save).
        mainDoc.Save(outputPath, SaveFormat.Docx);
    }
}
