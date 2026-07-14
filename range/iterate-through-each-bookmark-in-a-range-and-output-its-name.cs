using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few bookmarks to the document.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Move to a new line after each bookmark.
        }

        // Iterate through each bookmark in the document's range and output its name.
        foreach (Bookmark bookmark in doc.Range.Bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
