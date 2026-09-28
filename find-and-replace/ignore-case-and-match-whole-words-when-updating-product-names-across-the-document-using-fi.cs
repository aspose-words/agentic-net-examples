using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with product names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our productA is great.");
        builder.Writeln("producta is affordable.");
        builder.Writeln("productB is also good.");
        builder.Writeln("The productA's features are unmatched."); // This should not be replaced because of whole word match.

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find-and-replace.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace options: ignore case and match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // Ignore case.
            FindWholeWordsOnly = true   // Match whole words only.
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("productA", "SuperProduct", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output the result count (optional, not required for interaction).
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }
}
