using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the destination document (the one that will receive the insertion).
        Document destination = new Document("Destination.docx");

        // Load the source document that we want to insert.
        Document source = new Document("Source.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the builder's cursor to the insertion point.
        // In this example we use a bookmark named "InsertHere" that must exist in Destination.docx.
        builder.MoveToBookmark("InsertHere");

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to a new file.
        destination.Save("Result.docx");
    }
}
