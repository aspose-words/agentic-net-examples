using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and the bookmark.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the bookmark at the beginning of the document.
        builder.StartBookmark("StartBookmark");

        // Add a paragraph with sample text.
        builder.Writeln("Hello World!");

        // End the bookmark immediately after the inserted text.
        builder.EndBookmark("StartBookmark");

        // Save the document to a file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the bookmark was added.
        Bookmark bookmark = doc.Range.Bookmarks["StartBookmark"];
        Console.WriteLine($"Bookmark '{bookmark.Name}' has been added.");
    }
}
