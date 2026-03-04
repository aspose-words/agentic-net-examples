using System;
using Aspose.Words;

namespace AsposeWordsInsertExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (the document into which we will insert another document).
            Document destination = new Document("Destination.docx");

            // Load the source document (the document that will be inserted).
            Document source = new Document("Source.docx");

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destination);

            // Move the cursor to the end of the destination document.
            builder.MoveToDocumentEnd();

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the source document.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            destination.Save("CombinedResult.docx");
        }
    }
}
