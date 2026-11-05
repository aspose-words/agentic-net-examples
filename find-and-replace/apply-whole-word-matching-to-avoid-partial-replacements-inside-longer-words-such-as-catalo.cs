using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing the target word and a longer word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The catalog contains many items.");
        builder.Writeln("The catalogue is a different word.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find‑and‑replace processing.
        Document loaded = new Document(inputPath);

        // Use a regular expression with word‑boundary anchors to match whole words only.
        Regex wholeWordPattern = new Regex(@"\bcatalog\b", RegexOptions.IgnoreCase);
        FindReplaceOptions options = new FindReplaceOptions(); // default options

        // Replace the whole word "catalog" with "directory".
        int replacedCount = loaded.Range.Replace(wholeWordPattern, "directory", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
