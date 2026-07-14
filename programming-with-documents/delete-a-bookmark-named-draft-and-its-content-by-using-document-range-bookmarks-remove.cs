using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some regular text.
        builder.Writeln("Document start.");

        // Insert a bookmark named "Draft" with its own content.
        builder.StartBookmark("Draft");
        builder.Writeln("This is the draft content that will be removed.");
        builder.EndBookmark("Draft");

        // Add more text after the bookmark.
        builder.Writeln("Document end.");

        // Save the document before removing the bookmark (optional, for verification).
        string beforePath = "output_before.docx";
        doc.Save(beforePath);

        // Remove the bookmark named "Draft" using the BookmarkCollection.Remove method.
        doc.Range.Bookmarks.Remove("Draft");

        // Save the document after the bookmark has been removed.
        string afterPath = "output_after.docx";
        doc.Save(afterPath);
    }
}
