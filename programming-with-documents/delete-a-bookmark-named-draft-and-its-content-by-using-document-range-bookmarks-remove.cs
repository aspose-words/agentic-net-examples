using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the bookmark.
        builder.Writeln("Paragraph before the bookmark.");

        // Insert a bookmark named "Draft" with some content.
        builder.StartBookmark("Draft");
        builder.Writeln("This is the draft content that will be removed.");
        builder.EndBookmark("Draft");

        // Add some text after the bookmark.
        builder.Writeln("Paragraph after the bookmark.");

        // Save the original document (optional, for reference).
        doc.Save("Original.docx");

        // Remove the bookmark and its content.
        Bookmark draftBookmark = doc.Range.Bookmarks["Draft"];
        if (draftBookmark != null)
        {
            // Clear the bookmark's text (removes its content).
            draftBookmark.Text = string.Empty;

            // Remove the bookmark from the collection.
            doc.Range.Bookmarks.Remove("Draft");
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
