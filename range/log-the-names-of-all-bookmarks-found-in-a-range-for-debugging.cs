using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample bookmarks to the document.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Move to a new line after each bookmark.
        }

        // Retrieve the collection of bookmarks from the whole‑document range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Log each bookmark's name to the console for debugging.
        foreach (Bookmark bm in bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bm.Name}");
        }
    }
}
