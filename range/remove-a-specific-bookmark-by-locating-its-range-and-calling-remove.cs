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
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Optional: save the original document for reference.
        doc.Save("Original.docx");

        // Locate the bookmark we want to remove (e.g., "Bookmark_2").
        Bookmark bookmark = doc.Range.Bookmarks["Bookmark_2"];
        if (bookmark != null)
        {
            // Remove the bookmark from the document. The text inside the bookmark remains.
            bookmark.Remove();
        }

        // Save the document after the bookmark has been removed.
        doc.Save("Result.docx");
    }
}
