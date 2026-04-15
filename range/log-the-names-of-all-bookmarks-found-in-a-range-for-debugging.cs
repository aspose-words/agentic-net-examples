using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a few bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            string name = $"Bookmark_{i}";
            builder.StartBookmark(name);
            builder.Write($"Text inside {name}.");
            builder.EndBookmark(name);
            builder.Writeln(); // add a line break after each bookmark.
        }

        // Obtain the range that covers the whole document.
        Aspose.Words.Range range = doc.Range;

        // Retrieve all bookmarks within the range.
        BookmarkCollection bookmarks = range.Bookmarks;

        // Log the name of each bookmark for debugging purposes.
        foreach (Bookmark bookmark in bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bookmark.Name}");
        }
    }
}
