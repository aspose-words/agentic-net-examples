using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
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

        // Retrieve the number of bookmarks that exist in the document's range.
        int bookmarkCount = doc.Range.Bookmarks.Count;

        // Output the count to the console.
        Console.WriteLine($"Bookmark count in the document range: {bookmarkCount}");

        // Save the document (optional, demonstrates lifecycle handling).
        doc.Save("BookmarksCount.docx");
    }
}
