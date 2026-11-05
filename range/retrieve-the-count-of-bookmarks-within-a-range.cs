using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first bookmark.
        builder.StartBookmark("Bookmark1");
        builder.Writeln("First bookmarked paragraph.");
        builder.EndBookmark("Bookmark1");

        // Insert second bookmark.
        builder.StartBookmark("Bookmark2");
        builder.Writeln("Second bookmarked paragraph.");
        builder.EndBookmark("Bookmark2");

        // Insert a paragraph without a bookmark.
        builder.Writeln("Regular paragraph.");

        // Save the document.
        string filePath = "BookmarksDemo.docx";
        doc.Save(filePath);

        // Load the document.
        Document loadedDoc = new Document(filePath);

        // Retrieve the count of bookmarks within the document's range.
        int bookmarkCount = loadedDoc.Range.Bookmarks.Count;

        // Output the count.
        Console.WriteLine($"Number of bookmarks in the document range: {bookmarkCount}");
    }
}
