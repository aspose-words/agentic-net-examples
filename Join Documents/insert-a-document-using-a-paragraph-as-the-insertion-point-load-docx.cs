using System;
using Aspose.Words;

namespace InsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document (the one that contains the insertion point).
            Document destination = new Document("Destination.docx");

            // Load the source document that will be inserted.
            Document source = new Document("Source.docx");

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destination);

            // Move the cursor to the desired paragraph.
            // Parameters: paragraph index within the current section, node index within the paragraph.
            // Here we move to the first paragraph (index 0) of the first section.
            builder.MoveToParagraph(0, 0);

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the formatting of the source document.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

            // Save the modified document.
            destination.Save("Result.docx");
        }
    }
}
