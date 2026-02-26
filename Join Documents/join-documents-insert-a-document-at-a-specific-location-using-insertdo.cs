using System;
using Aspose.Words;

class JoinDocumentsExample
{
    static void Main()
    {
        // Paths to the documents.
        string mainDocPath = "MainDocument.docx";          // Document into which we will insert.
        string docToInsertPath = "DocumentToInsert.docx"; // Document that will be inserted.
        string resultPath = "JoinedDocument.docx";         // Output file.

        // Load the main document.
        Document mainDoc = new Document(mainDocPath);

        // Load the document that will be inserted.
        Document insertDoc = new Document(docToInsertPath);

        // Create a DocumentBuilder for the main document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Move the cursor to the desired insertion point.
        // Example: if a bookmark named "InsertHere" exists, insert at the bookmark.
        // Otherwise, insert at the end of the document.
        if (mainDoc.Range.Bookmarks["InsertHere"] != null)
        {
            builder.MoveToBookmark("InsertHere");
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        // Insert the second document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        mainDoc.Save(resultPath);
    }
}
