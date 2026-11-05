using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text and a bookmark named "MyBookmark".
        builder.Writeln("Hello world.");
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This is bookmarked text.");
        builder.EndBookmark("MyBookmark");

        // Save the original document (optional, just for reference).
        doc.Save("Original.docx");

        // Locate the bookmark by its name.
        Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];
        if (bookmark != null)
        {
            // Remove the bookmark using its Remove method.
            bookmark.Remove();
        }

        // Save the document after the bookmark has been removed.
        doc.Save("AfterRemoval.docx");
    }
}
