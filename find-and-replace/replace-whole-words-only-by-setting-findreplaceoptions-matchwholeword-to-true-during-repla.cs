using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing both whole‑word and partial matches.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);
        builder.Writeln("The quick brown fox jumps over the lazy dog. The foxes are clever.");

        // Save the sample document locally.
        const string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // Load the document from the file we just created.
        Document doc = new Document(inputPath);

        // Configure find‑and‑replace options to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // In Aspose.Words the property is FindWholeWordsOnly, not MatchWholeWord.
            FindWholeWordsOnly = true
        };

        // Replace the whole‑word occurrence of "fox" with "cat".
        int replacedCount = doc.Range.Replace("fox", "cat", options);

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement, but none occurred.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output the number of replacements performed.
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }
}
