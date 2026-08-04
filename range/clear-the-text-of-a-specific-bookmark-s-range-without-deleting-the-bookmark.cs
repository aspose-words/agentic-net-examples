using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark named "SampleBookmark" with some text inside it.
        string bookmarkName = "SampleBookmark";
        builder.StartBookmark(bookmarkName);
        builder.Write("This text will be cleared.");
        builder.EndBookmark(bookmarkName);

        // Save the original document (optional, just for reference).
        string originalPath = Path.Combine(Environment.CurrentDirectory, "Original.docx");
        doc.Save(originalPath);

        // Access the bookmark and clear its text while keeping the bookmark itself.
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];
        if (bookmark != null)
        {
            // Setting the Text property to an empty string removes the enclosed text.
            bookmark.Text = string.Empty;
        }

        // Verify that the bookmark still exists.
        bool bookmarkExists = doc.Range.Bookmarks[bookmarkName] != null;
        Console.WriteLine($"Bookmark exists after clearing text: {bookmarkExists}");

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ClearedBookmark.docx");
        doc.Save(outputPath);
    }
}
