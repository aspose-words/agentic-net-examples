using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content before the bookmark.
        builder.Writeln("Paragraph before the bookmark.");

        // Create a bookmark named "Draft" and put text inside it.
        builder.StartBookmark("Draft");
        builder.Writeln("This is the draft content that will be removed.");
        builder.EndBookmark("Draft");

        // Add some content after the bookmark.
        builder.Writeln("Paragraph after the bookmark.");

        // Remove the bookmark named "Draft" (the bookmark itself is removed; the text remains).
        doc.Range.Bookmarks.Remove("Draft");

        // Save the resulting document.
        const string outputFile = "Result.docx";
        doc.Save(outputFile);
    }
}
