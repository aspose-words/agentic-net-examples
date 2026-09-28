using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build a simple paragraph with a bookmark.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph before the bookmark.");

        // Insert a bookmark named "MyBookmark" around some text.
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This text is inside the bookmark.");
        builder.EndBookmark("MyBookmark");

        builder.Writeln("This is a paragraph after the bookmark.");

        // Save the document before removing the bookmark (optional verification step).
        string beforePath = "DocumentBeforeRemoval.docx";
        doc.Save(beforePath);

        // Locate the specific bookmark by name.
        Bookmark bookmark = doc.Range.Bookmarks["MyBookmark"];
        if (bookmark != null)
        {
            // Remove the bookmark by calling its Remove method.
            bookmark.Remove();
        }

        // Save the document after the bookmark has been removed.
        string afterPath = "DocumentAfterRemoval.docx";
        doc.Save(afterPath);
    }
}
