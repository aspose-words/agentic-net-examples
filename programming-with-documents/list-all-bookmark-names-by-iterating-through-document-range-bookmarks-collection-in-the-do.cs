using System;
using Aspose.Words;

public class ListBookmarksExample
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

        // Save the document (optional, but satisfies the save requirement).
        doc.Save("Bookmarks.docx");

        // Retrieve the collection of bookmarks from the document's range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate through the collection and print each bookmark's name.
        foreach (Bookmark bookmark in bookmarks)
        {
            Console.WriteLine(bookmark.Name);
        }
    }
}
