using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;

public class BookmarkExtractionExample
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words in some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define output folder and file names
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "GeneratedDocument.docx");

        // Create a new document and add bookmarks
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // List of bookmark names we expect to create
        var expectedBookmarks = new List<string> { "Bookmark1", "Bookmark2", "Bookmark3" };

        foreach (var name in expectedBookmarks)
        {
            builder.StartBookmark(name);
            builder.Writeln($"Content inside {name}");
            builder.EndBookmark(name);
        }

        // Save the generated document
        doc.Save(templatePath);

        // Load the document back from disk
        var loadedDoc = new Document(templatePath);

        // Extract bookmark names from the loaded document
        var extractedNames = loadedDoc.Range.Bookmarks
                                         .Select(b => b.Name)
                                         .ToList();

        // Verify that the extracted bookmark names match the expected list
        bool allMatch = expectedBookmarks.SequenceEqual(extractedNames);

        // Output verification result
        Console.WriteLine(allMatch
            ? "All bookmarks match the expected values."
            : "Bookmark names do not match the expected values.");

        // Optionally list the extracted bookmark names
        Console.WriteLine("Extracted bookmark names:");
        foreach (var name in extractedNames)
        {
            Console.WriteLine($"- {name}");
        }
    }
}
