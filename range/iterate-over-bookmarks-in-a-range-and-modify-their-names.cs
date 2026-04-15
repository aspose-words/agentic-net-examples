using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three bookmarks with some text.
        for (int i = 1; i <= 3; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.Writeln(); // Add a line break after each bookmark.
        }

        // Iterate over all bookmarks in the document's range and modify their names.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        for (int i = 0; i < bookmarks.Count; i++)
        {
            Bookmark bm = bookmarks[i];
            // Append "_Modified" to each bookmark name.
            bm.Name = bm.Name + "_Modified";
            Console.WriteLine($"Bookmark {i} new name: {bm.Name}");
        }

        // Save the modified document.
        string outputPath = "ModifiedBookmarks.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
