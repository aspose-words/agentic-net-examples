using System;
using Aspose.Words;

namespace InsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Load the destination document where the insertion will occur.
            Document destination = new Document("Destination.docx");

            // Create a DocumentBuilder to navigate and edit the destination document.
            DocumentBuilder builder = new DocumentBuilder(destination);

            // Move the cursor to a bookmark named "InsertHere".
            // Ensure the destination document contains this bookmark.
            builder.MoveToBookmark("InsertHere");

            // Load the source document that will be inserted.
            Document source = new Document("Source.docx");

            // Insert the source document at the current cursor position,
            // preserving the source formatting.
            builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

            // Save the modified document.
            destination.Save("Result.docx");
        }
    }
}
