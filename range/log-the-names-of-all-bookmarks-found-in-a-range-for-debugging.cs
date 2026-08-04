using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several bookmarks to the document.
        for (int i = 1; i <= 3; i++)
        {
            string name = $"Bookmark_{i}";
            builder.StartBookmark(name);
            builder.Write($"Text inside {name}.");
            builder.EndBookmark(name);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Retrieve the collection of bookmarks from the document's range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Log each bookmark's name to the console for debugging.
        foreach (Bookmark bookmark in bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bookmark.Name}");
        }
    }
}
