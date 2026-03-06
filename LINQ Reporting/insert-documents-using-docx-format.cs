using System;
using Aspose.Words;

namespace AsposeWordsInsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (or create a new blank one)
            Document dstDoc = new Document(); // creates a blank document
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // Move the cursor to the end of the destination document
            builder.MoveToDocumentEnd();

            // Optional: insert a page break before inserting the source document
            builder.InsertBreak(BreakType.PageBreak);

            // Load the source document that we want to insert
            Document srcDoc = new Document("SourceDocument.docx"); // replace with actual path

            // Insert the source document at the current cursor position
            // Keep the source formatting while inserting
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document in DOCX format
            dstDoc.Save("CombinedDocument.docx"); // replace with desired output path
        }
    }
}
