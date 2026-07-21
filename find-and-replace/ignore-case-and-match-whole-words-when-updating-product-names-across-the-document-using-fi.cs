using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various product name occurrences.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our new Widget is great.");
        builder.Writeln("The widget works well.");
        builder.Writeln("SuperWidget is not a widget.");
        builder.Writeln("widget");
        builder.Writeln("WIDGET");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Configure find-and-replace options: ignore case and match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // Case‑insensitive search.
            FindWholeWordsOnly = true   // Replace only whole word matches.
        };

        // Replace the product name "widget" with "Gadget".
        int replacedCount = loaded.Range.Replace("widget", "Gadget", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
