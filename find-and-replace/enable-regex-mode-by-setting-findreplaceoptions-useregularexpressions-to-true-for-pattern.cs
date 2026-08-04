using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing text that matches a regex pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Order 123 and Order 456");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Prepare find‑replace options (no special flags are required for regex usage).
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace all occurrences of the pattern "Order <number>" with "Order ###".
        // Use the Regex overload of Range.Replace to enable regular‑expression matching.
        int replacedCount = loaded.Range.Replace(new Regex(@"Order \d+"), "Order ###", options);

        // Verify that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Indicate success (no interactive input required).
        Console.WriteLine($"Replacements made: {replacedCount}");
    }
}
