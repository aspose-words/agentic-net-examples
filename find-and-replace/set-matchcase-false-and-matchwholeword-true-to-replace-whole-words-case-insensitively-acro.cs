using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Apple is tasty.");
        builder.Writeln("I like apple pies.");
        builder.Writeln("Pineapple is not an apple.");
        builder.Writeln("APPLE can be written in uppercase.");

        // Save the sample input (optional, demonstrates file handling).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace options:
        // - Case‑insensitive (MatchCase = false)
        // - Replace whole words only (FindWholeWordsOnly = true)
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("apple", "orange", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
