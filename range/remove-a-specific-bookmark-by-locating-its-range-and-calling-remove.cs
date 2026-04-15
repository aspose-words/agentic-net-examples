using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three bookmarks, each containing some text.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"Bookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Save the original document (optional, shows the state before removal).
        doc.Save("Original.docx");

        // Locate the specific bookmark by name.
        string targetBookmark = "Bookmark_2";
        Bookmark bookmark = doc.Range.Bookmarks[targetBookmark];

        // If the bookmark exists, remove it. The text inside the bookmark remains.
        if (bookmark != null)
        {
            bookmark.Remove();
        }

        // Save the document after the bookmark has been removed.
        doc.Save("Modified.docx");
    }
}
