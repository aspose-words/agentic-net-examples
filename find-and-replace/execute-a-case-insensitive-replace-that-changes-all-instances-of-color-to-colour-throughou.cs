using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing various cases of the word "color".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The color of the sky is blue.");
        builder.Writeln("COLOR is often used in design.");
        builder.Writeln("A colorful world."); // This occurrence should remain unchanged because it is part of a larger word.

        // Save the initial document (optional, demonstrates file I/O).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace to ignore case.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false // Case‑insensitive search.
        };

        // Perform the replacement: "color" → "colour".
        int replacedCount = loaded.Range.Replace("color", "colour", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
