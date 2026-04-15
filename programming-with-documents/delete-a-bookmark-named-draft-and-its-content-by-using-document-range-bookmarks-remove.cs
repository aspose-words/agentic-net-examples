using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the bookmark.
        builder.Writeln("Paragraph before the bookmark.");

        // Insert a bookmark named "Draft" with some content.
        builder.StartBookmark("Draft");
        builder.Writeln("This is the draft content that will be removed together with the bookmark.");
        builder.EndBookmark("Draft");

        // Add some text after the bookmark.
        builder.Writeln("Paragraph after the bookmark.");

        // Save the document before removal (optional, just to illustrate the original state).
        doc.Save("DocumentBeforeRemoval.docx");

        // Remove the bookmark named "Draft" using the BookmarkCollection.Remove method.
        // This call removes the bookmark definition; the text inside the bookmark remains in the document.
        doc.Range.Bookmarks.Remove("Draft");

        // Save the document after the bookmark has been removed.
        doc.Save("DocumentAfterRemoval.docx");
    }
}
