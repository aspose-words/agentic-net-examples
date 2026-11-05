using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that contains both lower‑case and capitalized forms.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog. The Quick brown fox.");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document that we just created.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options to be case‑sensitive.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true
        };

        // Perform the replacement: only the capitalized "Quick" should be replaced.
        int replacedCount = loaded.Range.Replace("Quick", "Swift", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one case‑sensitive replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optionally, write a simple confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }
}
