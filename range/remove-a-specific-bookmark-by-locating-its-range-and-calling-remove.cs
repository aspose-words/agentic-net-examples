using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several bookmarks with some text inside each.
        for (int i = 1; i <= 5; i++)
        {
            string bookmarkName = $"MyBookmark_{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"Text inside {bookmarkName}.");
            builder.EndBookmark(bookmarkName);
            builder.InsertBreak(BreakType.ParagraphBreak);
        }

        // Locate the bookmark named "MyBookmark_3".
        Bookmark bookmarkToRemove = doc.Range.Bookmarks["MyBookmark_3"];
        if (bookmarkToRemove != null)
        {
            // Remove the bookmark from the document (the text remains).
            bookmarkToRemove.Remove();
        }

        // Save the resulting document.
        string outputPath = "RemovedBookmark.docx";
        doc.Save(outputPath);
    }
}
