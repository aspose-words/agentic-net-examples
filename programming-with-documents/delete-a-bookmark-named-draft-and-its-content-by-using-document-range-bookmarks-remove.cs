using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content before the bookmark.
        builder.Writeln("This is before the bookmark.");

        // Insert a bookmark named "Draft" with some content.
        builder.StartBookmark("Draft");
        builder.Writeln("This is draft content to be removed.");
        builder.EndBookmark("Draft");

        // Add some content after the bookmark.
        builder.Writeln("This is after the bookmark.");

        // Delete the bookmark "Draft" and its content.
        doc.Range.Bookmarks.Remove("Draft");

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
