using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder cursor to the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert a bookmark that starts and ends at the document start.
        const string bookmarkName = "StartBookmark";
        builder.StartBookmark(bookmarkName);
        builder.EndBookmark(bookmarkName);

        // Add some sample text after the bookmark so the document is not empty.
        builder.Writeln("This is a sample document with a bookmark at the start.");

        // Verify that the bookmark was added using the Range.Bookmarks collection.
        BookmarkCollection bookmarks = doc.Range.Bookmarks;
        Console.WriteLine($"Total bookmarks: {bookmarks.Count}");
        if (bookmarks.Count > 0)
        {
            Console.WriteLine($"First bookmark name: {bookmarks[0].Name}");
            Console.WriteLine($"First bookmark text: \"{bookmarks[0].Text.Trim()}\"");
        }

        // Save the document to the local file system.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to \"{outputPath}\".");
    }
}
