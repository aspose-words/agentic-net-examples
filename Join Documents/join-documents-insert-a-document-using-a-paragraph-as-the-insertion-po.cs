using System;
using Aspose.Words;

class JoinDocumentsExample
{
    static void Main()
    {
        // Paths to the documents.
        string mainDocPath = "MainDocument.docx";      // Document that contains the insertion point.
        string insertDocPath = "DocumentToInsert.docx"; // Document to be inserted.
        string resultPath = "JoinedDocument.docx";      // Output file.

        // Load the main document (the one that already has a paragraph).
        Document mainDoc = new Document(mainDocPath);

        // Load the document that will be inserted.
        Document insertDoc = new Document(insertDocPath);

        // Create a DocumentBuilder for the main document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Move the cursor to the desired paragraph.
        // Here we move to the first paragraph (index 0) and its first node (offset 0).
        builder.MoveToParagraph(0, 0);

        // Insert the whole document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted document.
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        mainDoc.Save(resultPath);
    }
}
