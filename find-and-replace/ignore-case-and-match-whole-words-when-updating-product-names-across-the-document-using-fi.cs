using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with product names in various cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our catalog includes ProductA, productb, and PRODUCTC.");
        builder.Writeln("We also have producta and productb in stock.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find/replace to ignore case and match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // Ignore character case.
            FindWholeWordsOnly = true   // Replace only whole word matches.
        };

        // Perform replacements for each product name.
        int totalReplacements = 0;
        totalReplacements += loaded.Range.Replace("ProductA", "ItemX", options);
        totalReplacements += loaded.Range.Replace("productb", "ItemY", options);
        totalReplacements += loaded.Range.Replace("PRODUCTC", "ItemZ", options);

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
