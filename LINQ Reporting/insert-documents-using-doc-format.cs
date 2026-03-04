using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocExample
{
    static void Main()
    {
        // Load the destination document (could be a blank document or an existing .doc file).
        Document dstDoc = new Document("Destination.doc");

        // Create a DocumentBuilder attached to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Move the cursor to the end of the document and insert a page break before inserting.
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);

        // Load the source document that we want to insert.
        Document srcDoc = new Document("Source.doc");

        // Insert the source document into the destination document.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Prepare save options for the legacy .doc format.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

        // Save the combined document as a .doc file.
        dstDoc.Save("Combined.doc", saveOptions);
    }
}
