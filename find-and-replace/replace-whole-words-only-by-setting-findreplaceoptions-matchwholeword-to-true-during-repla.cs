using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text containing the target word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("A foxes are clever.");
        builder.Writeln("fox");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Use a regular expression with word boundaries to match whole words only.
        Regex wholeWordFox = new Regex(@"\bfox\b", RegexOptions.IgnoreCase);
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace the whole word "fox" with "cat".
        int replaced = loaded.Range.Replace(wholeWordFox, "cat", options);

        // Verify that at least one replacement occurred.
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
