using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few bookmarks with some text inside each.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Obtain the range that covers the whole document.
        Aspose.Words.Range range = doc.Range;

        // Log the name of each bookmark found in the range.
        foreach (Bookmark bookmark in range.Bookmarks)
        {
            Console.WriteLine($"Bookmark name: {bookmark.Name}");
        }

        // Save the document (optional, demonstrates full lifecycle).
        doc.Save("BookmarksDemo.docx");
    }
}
