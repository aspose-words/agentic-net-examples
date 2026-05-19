using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert content and a bookmark at the very start of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the beginning of the document.
        builder.MoveToDocumentStart();

        // Start a bookmark named "StartBookmark".
        builder.StartBookmark("StartBookmark");

        // Insert some text that will be inside the bookmark.
        builder.Write("This text is inside the start bookmark.");

        // End the bookmark.
        builder.EndBookmark("StartBookmark");

        // Optionally add more content after the bookmark.
        builder.Writeln();
        builder.Write("Additional document content.");

        // Verify that the bookmark was added using the Range.Bookmarks collection.
        Bookmark bookmark = doc.Range.Bookmarks["StartBookmark"];
        if (bookmark != null)
        {
            Console.WriteLine($"Bookmark '{bookmark.Name}' added successfully. Text: \"{bookmark.Text}\"");
        }
        else
        {
            Console.WriteLine("Failed to add the bookmark.");
        }

        // Save the document to the local file system.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
