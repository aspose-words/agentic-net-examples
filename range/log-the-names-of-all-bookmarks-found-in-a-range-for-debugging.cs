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
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Obtain the range that covers the whole document.
        Aspose.Words.Range range = doc.Range; // Fully qualified to avoid conflict with System.Range

        // Log the names of all bookmarks found in this range.
        Console.WriteLine("Bookmarks found in the document range:");
        foreach (Bookmark bookmark in range.Bookmarks)
        {
            Console.WriteLine($"- {bookmark.Name}");
        }

        // Optionally save the document to verify the bookmarks exist.
        doc.Save("Bookmarks.docx");
    }
}
