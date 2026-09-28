using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with text that matches a pattern.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Order 123 has been shipped.");
        builder.Writeln("Order 456 is pending.");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for find‑and‑replace.
        var loadedDoc = new Document(inputPath);

        // Pattern to find: "Order " followed by one or more digits.
        var regex = new Regex(@"Order \d+");
        // Replacement text.
        const string replacement = "Order ###";

        // Perform the replacement using a regular expression.
        int replacedCount = loadedDoc.Range.Replace(regex, replacement, new FindReplaceOptions());

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
