using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocumentExample
{
    static void Main()
    {
        // Load the destination document where the insertion will occur.
        Document destination = new Document("Destination.docx");

        // Create a DocumentBuilder to navigate and edit the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the builder's cursor to the desired insertion point.
        // In this example we move to the end of the document.
        builder.MoveToDocumentEnd();

        // Optional: insert a page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);

        // Load the source document that will be inserted.
        Document source = new Document("Source.docx");

        // Insert the source document at the current cursor position,
        // preserving the source formatting.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting combined document.
        destination.Save("Combined.docx");
    }
}
