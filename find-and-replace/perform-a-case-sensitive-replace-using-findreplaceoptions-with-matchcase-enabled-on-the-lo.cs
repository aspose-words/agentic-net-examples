using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document with known text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a test. world.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document from the file system.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // Configure find‑replace options to be case‑sensitive.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true
        };

        // Replace the exact word "World" with "Universe".
        int replacedCount = loaded.Range.Replace("World", "Universe", options);

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one case‑sensitive replacement.");

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // Optional: indicate success (no interactive input required).
        Console.WriteLine($"Replacements made: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
