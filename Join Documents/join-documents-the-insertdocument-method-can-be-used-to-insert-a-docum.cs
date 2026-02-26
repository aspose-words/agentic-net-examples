using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeWordsExamples
{
    class DocumentJoinExample
    {
        static void Main()
        {
            // Load the destination document (the document into which we will insert another document).
            Document destination = new Document("Destination.docx");

            // Load the source document (the document to be inserted).
            Document source = new Document("Source.docx");

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destination);

            // Move the cursor to the end of the destination document.
            builder.MoveToDocumentEnd();

            // Optionally insert a page break before the inserted content.
            builder.InsertBreak(BreakType.PageBreak);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the source document.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document to a new file.
            destination.Save("Joined.docx");
        }
    }
}
