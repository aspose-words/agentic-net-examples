using System;
using Aspose.Words;

class JoinDocumentsExample
{
    static void Main()
    {
        // Load the source document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Create a new (or load an existing) destination document.
        Document dstDoc = new Document();

        // Use DocumentBuilder to position the cursor where the source will be inserted.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.MoveToDocumentEnd();               // Insert at the end of the destination.
        builder.InsertBreak(BreakType.PageBreak);  // Optional: add a page break before insertion.

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("JoinedDocument.docx");
    }
}
