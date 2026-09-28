using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing a whole word ("catalog") and a longer word ("catalogue").
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The catalog is updated. The catalogue includes many items.");

        // Save the document locally so it can be re‑loaded later.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document from the file.
        Document loaded = new Document(inputPath);

        // Configure find‑replace to match whole words only using a regular expression with word boundaries.
        Regex wholeWordPattern = new Regex(@"\bcatalog\b", RegexOptions.None);
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace the whole word "catalog" with "directory".
        int replacedCount = loaded.Range.Replace(wholeWordPattern, "directory", options);

        // Verify that exactly one replacement occurred (the whole word, not the part of "catalogue").
        if (replacedCount != 1)
            throw new InvalidOperationException($"Expected 1 replacement, but got {replacedCount}.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loaded.Save(outputPath);
    }
}
