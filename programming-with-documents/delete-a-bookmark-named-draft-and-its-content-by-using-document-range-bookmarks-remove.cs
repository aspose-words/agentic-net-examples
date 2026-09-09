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

        // Insert some text before the bookmark.
        builder.Writeln("Paragraph before the bookmark.");

        // Create a bookmark named "Draft" and add content inside it.
        builder.StartBookmark("Draft");
        builder.Writeln("This is the draft content that will be removed.");
        builder.EndBookmark("Draft");

        // Insert some text after the bookmark.
        builder.Writeln("Paragraph after the bookmark.");

        // Retrieve the bookmark.
        Bookmark draftBookmark = doc.Range.Bookmarks["Draft"];
        if (draftBookmark != null)
        {
            // Delete the text inside the bookmark.
            draftBookmark.Text = string.Empty;

            // Remove the bookmark from the collection using Document.Range.Bookmarks.Remove.
            doc.Range.Bookmarks.Remove(draftBookmark);
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
        doc.Save(outputPath);
    }
}
