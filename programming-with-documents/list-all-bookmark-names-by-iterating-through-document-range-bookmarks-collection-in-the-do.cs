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
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Retrieve the collection of bookmarks from the document's range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // List all bookmark names to the console.
        foreach (Bookmark bookmark in bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }

        // Optionally save the document (not required for the listing task).
        doc.Save("BookmarksSample.docx");
    }
}
