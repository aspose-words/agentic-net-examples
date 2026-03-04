using System;
using Aspose.Words;

namespace AsposeWordsInsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (or create a new blank one).
            Document dstDoc = new Document(); // creates a blank document
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // Optionally add some initial content.
            builder.Writeln("This is the destination document.");

            // Load the source document that we want to insert.
            // Replace the path with the actual location of your source .docx file.
            Document srcDoc = new Document("SourceDocument.docx");

            // Move the cursor to the end of the destination document.
            builder.MoveToDocumentEnd();

            // Insert a page break before the inserted content (optional).
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the source.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            // The file format is inferred from the extension (.docx).
            dstDoc.Save("CombinedDocument.docx");
        }
    }
}
