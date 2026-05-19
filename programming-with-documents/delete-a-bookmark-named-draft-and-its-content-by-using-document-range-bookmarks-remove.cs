using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text before the bookmark.
        builder.Writeln("This is before the draft.");

        // Insert the bookmark named "Draft" with some content.
        builder.StartBookmark("Draft");
        builder.Writeln("This is draft content that will be removed.");
        builder.EndBookmark("Draft");

        // Add some text after the bookmark.
        builder.Writeln("This is after the draft.");

        // Retrieve the bookmark.
        Bookmark draftBookmark = doc.Range.Bookmarks["Draft"];
        if (draftBookmark != null)
        {
            // Delete the content inside the bookmark.
            draftBookmark.Text = string.Empty;

            // Remove the bookmark itself using the collection's Remove method.
            doc.Range.Bookmarks.Remove(draftBookmark);
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}
