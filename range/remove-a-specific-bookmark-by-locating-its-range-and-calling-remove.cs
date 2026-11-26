using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text.
        builder.Writeln("This is a sample document.");

        // Insert a bookmark around a paragraph.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        // Save the document before removing the bookmark.
        doc.Save("Before.docx");

        // Locate the bookmark by name.
        Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];
        if (bookmark != null)
        {
            // Remove the bookmark using its range.
            bookmark.Remove();
        }

        // Save the document after the bookmark has been removed.
        doc.Save("After.docx");
    }
}
