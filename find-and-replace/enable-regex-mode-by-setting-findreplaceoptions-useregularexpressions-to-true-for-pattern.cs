using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that matches a regular expression pattern.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Order 123 has been shipped.");
        builder.Writeln("Order 456 is pending.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Configure FindReplaceOptions (no special settings needed for regex).
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace any occurrence of "Order <number>" with "Order ###" using a regex pattern.
        int replacedCount = loaded.Range.Replace(new Regex(@"Order \d+"), "Order ###", options);

        // Verify that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
