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
            builder.Writeln(); // start a new paragraph after each bookmark.
        }

        // Retrieve the number of bookmarks that exist in the document's whole range.
        int bookmarkCount = doc.Range.Bookmarks.Count;

        // Output the count to the console.
        Console.WriteLine($"Number of bookmarks in the document: {bookmarkCount}");

        // Save the document (optional, demonstrates the full lifecycle).
        doc.Save("BookmarksCount.docx");
    }
}
