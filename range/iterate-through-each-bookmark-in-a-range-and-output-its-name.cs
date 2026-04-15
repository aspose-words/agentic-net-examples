using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few bookmarks with some text inside each.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Move to a new line after each bookmark.
        }

        // Get the collection of all bookmarks in the whole‑document range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate through each bookmark and output its name.
        foreach (Bookmark bookmark in bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
