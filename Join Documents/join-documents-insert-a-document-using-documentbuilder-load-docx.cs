using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Create a new (blank) destination document.
        Document dstDoc = new Document();

        // Attach a DocumentBuilder to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Position the builder at the end of the destination document.
        builder.MoveToDocumentEnd();

        // Optional: insert a page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the source document, keeping its original formatting.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        dstDoc.Save("Combined.docx");
    }
}
