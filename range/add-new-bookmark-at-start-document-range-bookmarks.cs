using System;
using Aspose.Words;

class AddBookmarkAtStart
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert a bookmark named "StartBookmark" at the start position.
        // The bookmark has no content – it starts and ends at the same location.
        builder.StartBookmark("StartBookmark");
        builder.EndBookmark("StartBookmark");

        // Verify that the bookmark was added.
        Console.WriteLine($"Bookmarks count: {doc.Range.Bookmarks.Count}");
        Console.WriteLine($"First bookmark name: {doc.Range.Bookmarks[0].Name}");

        // Save the document to a file.
        doc.Save("BookmarkAtStart.docx");
    }
}
