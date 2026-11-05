using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that matches a regex pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Order 123 and Order 456");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define the regex pattern to find "Order <digits>".
        Regex pattern = new Regex(@"Order \d+");

        // Perform the replacement using the regex overload.
        int replacedCount = loaded.Range.Replace(pattern, "Order ###", new FindReplaceOptions());

        // Validate that the expected number of replacements occurred.
        if (replacedCount != 2)
            throw new InvalidOperationException($"Expected 2 replacements, but got {replacedCount}.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
