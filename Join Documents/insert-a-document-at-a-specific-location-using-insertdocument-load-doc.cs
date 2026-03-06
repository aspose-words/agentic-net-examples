using System;
using Aspose.Words;

class InsertDocumentExample
{
    static void Main()
    {
        // Load the destination document that already contains a bookmark named "InsertHere".
        Document destination = new Document("Destination.docx");

        // Create a DocumentBuilder to work with the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the cursor to the bookmark where the source document will be inserted.
        builder.MoveToBookmark("InsertHere");

        // Load the source document that we want to insert.
        Document source = new Document("Source.docx");

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        destination.Save("Result.docx");
    }
}
