using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert an empty bookmark at the start of the document.
        builder.StartBookmark("StartBookmark");
        builder.EndBookmark("StartBookmark");

        // Add some sample text after the bookmark.
        builder.Writeln("This is the document content.");

        // Retrieve the bookmark via the document's Range.Bookmarks collection.
        Bookmark bookmark = doc.Range.Bookmarks["StartBookmark"];
        Console.WriteLine($"Bookmark '{bookmark.Name}' added. Text inside bookmark: '{bookmark.Text}'");

        // Save the document to the local file system.
        doc.Save("Output.docx");
    }
}
