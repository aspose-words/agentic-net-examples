using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the package list, not used directly

public class FindReplaceDemo
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "products_input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "products_output.docx");

        // -----------------------------------------------------------------
        // Create a sample document with various product name occurrences.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Introducing the SuperWidget, the best widget on the market.");
        builder.Writeln("Our customers love the superwidget for its reliability.");
        builder.Writeln("Check out the SuperWidgetPro for advanced features."); // Not a whole-word match.
        builder.Writeln("SUPERWIDGET is now available in bulk.");

        // Save the source document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and perform a case‑insensitive, whole‑word replace.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // Ignore case.
            FindWholeWordsOnly = true   // Replace only whole words.
        };

        int replacedCount = loaded.Range.Replace("SuperWidget", "MegaGadget", options);

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Output the result count (optional, not required for non‑interactive execution).
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }
}
