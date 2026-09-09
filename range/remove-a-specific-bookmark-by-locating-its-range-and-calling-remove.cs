using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert five bookmarks with text inside each.
        for (int i = 1; i <= 5; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.InsertBreak(BreakType.ParagraphBreak);
        }

        // Locate the specific bookmark by name.
        string targetBookmarkName = "MyBookmark_3";
        Bookmark targetBookmark = doc.Range.Bookmarks[targetBookmarkName];

        // Remove the bookmark (the text remains in the document).
        if (targetBookmark != null)
        {
            targetBookmark.Remove();
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
