using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three bookmarks with some text inside each.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a paragraph break after each bookmark.
        }

        // Retrieve the collection of bookmarks that exist in the whole document range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Get the count of bookmarks.
        int bookmarkCount = bookmarks.Count;

        // Output the count to the console.
        Console.WriteLine($"Number of bookmarks in the document range: {bookmarkCount}");
    }
}
