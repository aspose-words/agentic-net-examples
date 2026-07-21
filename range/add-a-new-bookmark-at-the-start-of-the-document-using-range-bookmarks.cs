using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to position the cursor at the very beginning of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentStart();

        // Insert an empty bookmark at the start of the document.
        const string bookmarkName = "StartBookmark";
        builder.StartBookmark(bookmarkName);
        builder.EndBookmark(bookmarkName);

        // Optionally add some text after the bookmark to make the document visible.
        builder.Writeln("Document content after the start bookmark.");

        // Access the bookmarks collection via the document's Range.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;

        // Output basic information about the added bookmark.
        Console.WriteLine($"Bookmarks count: {bookmarks.Count}");
        if (bookmarks.Count > 0)
        {
            Console.WriteLine($"First bookmark name: {bookmarks[0].Name}");
        }

        // Save the document to the local file system.
        const string outputPath = "BookmarkAdded.docx";
        doc.Save(outputPath);
    }
}
