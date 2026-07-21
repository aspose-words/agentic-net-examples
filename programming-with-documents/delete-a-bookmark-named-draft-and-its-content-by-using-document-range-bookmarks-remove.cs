using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a bookmark named "Draft"
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before bookmark.");
        builder.StartBookmark("Draft");
        builder.Writeln("This is draft content that will be removed.");
        builder.EndBookmark("Draft");
        builder.Writeln("Paragraph after bookmark.");

        // Save the original document (optional, for verification)
        doc.Save("Original.docx");

        // Locate the bookmark named "Draft"
        Bookmark draftBookmark = doc.Range.Bookmarks["Draft"];
        if (draftBookmark != null)
        {
            // Delete the content inside the bookmark
            draftBookmark.Text = string.Empty;

            // Remove the bookmark from the collection using Document.Range.Bookmarks.Remove
            doc.Range.Bookmarks.Remove(draftBookmark);
        }

        // Save the modified document
        doc.Save("Modified.docx");
    }
}
