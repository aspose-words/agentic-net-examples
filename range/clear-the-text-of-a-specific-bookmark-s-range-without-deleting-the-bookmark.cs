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

        // Clear the text that the bookmark encloses while keeping the bookmark itself.
        bookmark.Text = string.Empty;

        // Save the resulting document to verify the operation.
        const string outputPath = "ClearBookmarkText.docx";
        doc.Save(outputPath);

        // Output the bookmark's text after clearing to the console (should be empty).
        Console.WriteLine($"Bookmark '{bookmark.Name}' text after clearing: '{bookmark.Text}'");
    }
}
