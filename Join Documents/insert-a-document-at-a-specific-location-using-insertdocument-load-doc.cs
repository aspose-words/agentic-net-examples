using System;
using Aspose.Words;

namespace InsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document where we want to insert another document.
            Document dstDoc = new Document("Destination.docx");

            // Create a DocumentBuilder to navigate and edit the destination document.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // Move the cursor to the desired insertion point.
            // Example: move to the end of the document and add a page break before insertion.
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            // Load the source document that will be inserted.
            Document srcDoc = new Document("Source.docx");

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the inserted content.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            dstDoc.Save("Result.docx");
        }
    }
}
