using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document where the insertion will occur.
        Document destination = new Document("Destination.docx");

        // Load the source document that will be inserted.
        Document source = new Document("Source.docx");

        // Locate the bookmark that marks the insertion point.
        // The bookmark must exist in the destination document.
        Bookmark insertBookmark = destination.Range.Bookmarks["InsertHere"];
        if (insertBookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' not found.");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move the builder's cursor to the start of the bookmark.
        builder.MoveToBookmark("InsertHere");

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the source document.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the modified document.
        destination.Save("Result.docx");
    }
}
