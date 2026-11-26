using System;
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
        builder.Writeln("Order 123 has been shipped.");
        builder.Writeln("Order 456 is pending.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define the regex pattern and the replacement text.
        Regex regex = new Regex(@"Order \d+");
        string replacement = "Order ###";

        // Perform the replacement using the regex overload.
        int replacedCount = loaded.Range.Replace(regex, replacement, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
