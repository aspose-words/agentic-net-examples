using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;
using Aspose.Drawing; // Included to satisfy package requirement

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing product names.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Our new Widget is great.");
        builder.Writeln("The widget is available.");
        builder.Writeln("Widgets are popular.");          // Should NOT be replaced (not a whole word).
        builder.Writeln("SuperWidget is a brand.");      // Should NOT be replaced (part of a larger word).

        // Configure FindReplaceOptions to ignore case and match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,          // Ignore case.
            FindWholeWordsOnly = true   // Replace only whole-word matches.
        };

        // Perform the replacement: replace "Widget" with "Gadget".
        int replacedCount = doc.Range.Replace("Widget", "Gadget", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the updated document.
        const string outputPath = "updated-product-names.docx";
        doc.Save(outputPath);

        // Optional: write a simple JSON report of the operation.
        var report = new
        {
            SearchTerm = "Widget",
            Replacement = "Gadget",
            ReplacementsMade = replacedCount,
            OutputFile = Path.GetFullPath(outputPath)
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("replacement-report.json", jsonReport);
    }
}
