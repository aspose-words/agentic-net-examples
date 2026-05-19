using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing text that matches the regex pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Order 123 and Order 456");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Use a Regex pattern for the replace operation.
        Regex pattern = new Regex(@"Order \d+");
        FindReplaceOptions options = new FindReplaceOptions(); // No Need for UseRegularExpressions property.

        // Replace all occurrences of "Order <number>" with a placeholder.
        int replacedCount = loaded.Range.Replace(pattern, "Order ###", options);

        // Verify that the expected number of replacements occurred.
        if (replacedCount != 2)
            throw new InvalidOperationException($"Expected 2 replacements, but got {replacedCount}.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
