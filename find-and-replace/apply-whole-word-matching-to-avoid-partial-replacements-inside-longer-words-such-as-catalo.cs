using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with words that could be partially matched.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The catalog contains many items.");
        builder.Writeln("Please refer to the catalogue for details.");
        builder.Writeln("Our catalog is updated yearly.");

        // Save the original document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the whole word "catalog" with "list".
        int replaced = loaded.Range.Replace("catalog", "list", options);

        // Ensure that at least one replacement occurred.
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Output the result to the console (no user interaction required).
        Console.WriteLine($"Replacements made: {replaced}");
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}
