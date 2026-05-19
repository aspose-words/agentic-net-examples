using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document with varied casing of the word "example".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Aspose example.");
        builder.Writeln("Another Example.");
        builder.Writeln("EXAMPLE test.");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options: case‑insensitive and whole‑word only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,               // Ignore case.
            FindWholeWordsOnly = true        // Replace whole words only.
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("example", "sample", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
