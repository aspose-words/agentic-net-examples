using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Required package reference

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that includes the word "catalogue"
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The product catalogue contains many items.");
        builder.Writeln("Please refer to the catalogue for details.");
        builder.Writeln("Do not replace the word cataloguing.");

        // Save the source document locally
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing
        Document loaded = new Document(inputPath);

        // Configure find‑replace to match whole words only
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the whole word "catalogue" with "catalog"
        int replacedCount = loaded.Range.Replace("catalogue", "catalog", options);

        // Validate that at least one replacement occurred
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");

        // Save the modified document
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
