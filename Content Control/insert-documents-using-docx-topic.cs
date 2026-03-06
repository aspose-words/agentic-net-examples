using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocumentExample
{
    static void Main()
    {
        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Attach a DocumentBuilder to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Move the cursor to the end of the document and insert a page break
        // so the inserted content starts on a new page.
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Load the source document that we want to insert.
        Document srcDoc = new Document("SourceDocument.docx");

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("CombinedDocument.docx");
    }
}
