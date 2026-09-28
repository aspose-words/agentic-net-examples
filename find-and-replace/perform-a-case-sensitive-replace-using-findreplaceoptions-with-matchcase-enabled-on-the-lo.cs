using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with mixed‑case text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("the quick brown fox jumps over the lazy dog.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Configure case‑sensitive find‑replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = true
        };

        // Replace only the lower‑case occurrence of "the" with "a".
        int replacedCount = loaded.Range.Replace("the", "a", options);

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one case‑sensitive replacement, but none were made.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
