using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few bookmarks into the document.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Retrieve the count of bookmarks in the whole‑document range.
        int bookmarkCount = doc.Range.Bookmarks.Count;

        // Output the result.
        Console.WriteLine($"Number of bookmarks in the document range: {bookmarkCount}");
    }
}
