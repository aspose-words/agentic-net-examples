using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare file paths.
        string folder = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(folder);
        string docPath = Path.Combine(folder, "GeneratedDocument.docx");

        // Create a new document and add bookmarks.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // First bookmark.
        builder.Writeln("This is the first paragraph.");
        builder.StartBookmark("FirstBookmark");
        builder.Writeln("Content inside first bookmark.");
        builder.EndBookmark("FirstBookmark");

        // Second bookmark.
        builder.Writeln("This is the second paragraph.");
        builder.StartBookmark("SecondBookmark");
        builder.Writeln("Content inside second bookmark.");
        builder.EndBookmark("SecondBookmark");

        // Save the document.
        doc.Save(docPath);

        // Load the document back (demonstrating load step).
        Document loadedDoc = new(docPath);

        // Extract bookmark names.
        List<string> actualBookmarkNames = loadedDoc.Range.Bookmarks
            .Select(b => b.Name)
            .ToList();

        // Expected bookmark names.
        List<string> expectedBookmarkNames = new() { "FirstBookmark", "SecondBookmark" };

        // Compare (order‑independent).
        bool match = actualBookmarkNames.OrderBy(n => n).SequenceEqual(expectedBookmarkNames.OrderBy(n => n));

        // Output verification result.
        Console.WriteLine($"Bookmark extraction {(match ? "succeeded" : "failed")}.");
        Console.WriteLine("Expected: " + string.Join(", ", expectedBookmarkNames));
        Console.WriteLine("Actual:   " + string.Join(", ", actualBookmarkNames));
    }
}
