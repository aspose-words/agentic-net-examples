using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few bookmarks into the document.
        for (int i = 1; i <= 3; i++)
        {
            string name = $"Bookmark_{i}";
            builder.StartBookmark(name);
            builder.Write($"Text inside {name}.");
            builder.EndBookmark(name);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Retrieve the collection of bookmarks that belong to the whole document range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Get the count of bookmarks.
        int bookmarkCount = bookmarks.Count;

        // Output the count to the console.
        Console.WriteLine($"Number of bookmarks in the document range: {bookmarkCount}");

        // Save the document (optional, demonstrates the lifecycle rule).
        doc.Save("BookmarksCount.docx");
    }
}
