using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that includes a whole word and a partial match.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The catalogue is ready.");               // Should remain unchanged.
        builder.Writeln("Please see the catalog for details.");   // Should be replaced.

        // Save the initial document (seed input file).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Use a regular expression with word‑boundary anchors to match whole words only.
        Regex wholeWordPattern = new Regex(@"\bcatalog\b", RegexOptions.IgnoreCase);

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(wholeWordPattern, "directory", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
