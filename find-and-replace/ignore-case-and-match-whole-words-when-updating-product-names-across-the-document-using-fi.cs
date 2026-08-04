using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file paths for the input and output documents.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Clean up any previous runs.
        if (File.Exists(inputPath)) File.Delete(inputPath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // -------------------- Create sample document --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our catalog includes ProductA, productb, and PRODUCTC.");
        builder.Writeln("Special offer: producta is now cheaper.");
        doc.Save(inputPath); // Save the source document.

        // -------------------- Load and replace --------------------
        Document loaded = new Document(inputPath);

        // Configure find-replace to ignore case and match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform replacements for each product name.
        int replaced = loaded.Range.Replace("ProductA", "ItemX", options);
        replaced += loaded.Range.Replace("productb", "ItemY", options);
        replaced += loaded.Range.Replace("PRODUCTC", "ItemZ", options);

        // Validate that at least one replacement occurred.
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);
    }
}
