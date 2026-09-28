using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text so the document has content.
        builder.Writeln("This is the first paragraph.");

        // Move the cursor to the very beginning of the document.
        builder.MoveToDocumentStart();

        // Insert a zero‑length bookmark named "StartBookmark" at the start.
        builder.StartBookmark("StartBookmark");
        builder.EndBookmark("StartBookmark");

        // Verify that the bookmark was added.
        if (doc.Range.Bookmarks["StartBookmark"] != null)
        {
            Console.WriteLine("Bookmark 'StartBookmark' successfully added.");
        }
        else
        {
            Console.WriteLine("Failed to add bookmark 'StartBookmark'.");
        }

        // Save the document to verify the bookmark was added.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
