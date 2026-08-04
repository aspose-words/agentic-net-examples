using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three bookmarks with some text inside each.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Content of {bookmarkName}");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Move to next line.
        }

        // Get the collection of all bookmarks in the whole document range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Iterate over each bookmark and modify its name.
        foreach (Bookmark bookmark in bookmarks)
        {
            // Append "_Modified" to the original bookmark name.
            bookmark.Name = $"{bookmark.Name}_Modified";
        }

        // Save the modified document to disk.
        doc.Save("ModifiedBookmarks.docx");
    }
}
