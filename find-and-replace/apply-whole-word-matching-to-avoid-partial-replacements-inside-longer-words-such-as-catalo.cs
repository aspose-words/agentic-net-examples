using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class WholeWordReplaceExample
{
    public static void Main()
    {
        // Create a sample document with text that contains the target word both as a whole word and as part of a longer word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The catalog is ready.");
        builder.Writeln("The catalogue is complete.");
        builder.Writeln("Please review the catalog.");

        // Save the source document (optional, just to demonstrate file I/O).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Configure find-and-replace options to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the whole word "catalog" with "list".
        int replacedCount = loadedDoc.Range.Replace("catalog", "list", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole-word replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Output the result count (no interactive prompts).
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }
}
