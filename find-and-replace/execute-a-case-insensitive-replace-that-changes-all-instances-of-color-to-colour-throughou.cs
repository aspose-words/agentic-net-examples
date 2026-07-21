using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various case forms of the word "color".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("She likes the Color red.");
        builder.Writeln("COLOR is often used in design.");
        builder.Writeln("No matching word here.");

        // Save the source document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace to ignore case.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // case‑insensitive search
        };

        // Perform the replacement: "color" → "colour".
        int replacedCount = loaded.Range.Replace("color", "colour", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loaded.Save(outputPath);
    }
}
