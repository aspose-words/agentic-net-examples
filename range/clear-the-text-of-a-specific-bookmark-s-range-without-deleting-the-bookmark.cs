using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark with some text inside it.
        const string bookmarkName = "MyBookmark";
        builder.StartBookmark(bookmarkName);
        builder.Write("Text inside the bookmark.");
        builder.EndBookmark(bookmarkName);

        // Retrieve the bookmark from the document.
        Bookmark bookmark = doc.Range.Bookmarks[bookmarkName];

        // Clear the text that is enclosed by the bookmark.
        // This does not remove the bookmark itself.
        bookmark.Text = string.Empty;

        // Save the resulting document to verify the bookmark remains.
        const string outputPath = "ClearedBookmark.docx";
        doc.Save(outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Bookmark \"{bookmarkName}\" text cleared. Document saved to {outputPath}.");
    }
}
