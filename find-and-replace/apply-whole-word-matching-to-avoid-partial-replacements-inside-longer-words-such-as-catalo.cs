using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with words that will be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The product catalogue is ready.");
        builder.Writeln("Our catalogue2021 version includes new items.");
        builder.Writeln("Please review the catalogue before purchase.");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find‑replace to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true
        };

        // Replace the word "catalogue" with "catalog".
        int replacedCount = loaded.Range.Replace("catalogue", "catalog", options);

        // Ensure that at least one whole‑word replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one whole‑word replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
