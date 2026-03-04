using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsInsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (or create a new blank one).
            Document dstDoc = new Document("Destination.docx");

            // Create a DocumentBuilder attached to the destination document.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // Move the cursor to the end of the document where we want to insert the source.
            builder.MoveToDocumentEnd();

            // Optionally insert a page break before the inserted content.
            builder.InsertBreak(BreakType.PageBreak);

            // Load the source document that will be inserted.
            Document srcDoc = new Document("Source.docx");

            // Insert the source document into the destination.
            // KeepSourceFormatting preserves the original formatting of the source.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            dstDoc.Save("Combined.docx");
        }
    }
}
