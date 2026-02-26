using System;
using Aspose.Words;

class JoinDocumentsWithBookmark
{
    static void Main()
    {
        // Load the destination document that contains a bookmark named "InsertHere".
        Document destination = new Document("Destination.docx");

        // Load the source document that we want to insert.
        Document source = new Document("Source.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the cursor to the bookmark where the source document will be inserted.
        builder.MoveToBookmark("InsertHere");

        // Insert the source document at the bookmark position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        destination.Save("CombinedResult.docx");
    }
}
