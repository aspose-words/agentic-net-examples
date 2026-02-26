using System;
using Aspose.Words;

class JoinDocumentsExample
{
    static void Main()
    {
        // Load the source DOCX document that will be inserted.
        Document srcDoc = new Document("SourceDocument.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Use DocumentBuilder to position the cursor at the end of the destination body.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.MoveToDocumentEnd();

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("JoinedDocument.docx");
    }
}
