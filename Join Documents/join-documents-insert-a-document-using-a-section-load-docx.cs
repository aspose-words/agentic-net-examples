using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document (the one we will append to).
        Document dstDoc = new Document("Destination.docx");

        // Load the source document that we want to insert.
        Document srcDoc = new Document("Source.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Move the builder cursor to the end of the document.
        builder.MoveToDocumentEnd();

        // Insert a page break before the inserted content (optional).
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the source document at the current cursor position,
        // preserving the source formatting.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("Joined.docx");
    }
}
