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

        // Save the source document (demonstrates the create/save lifecycle).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system (demonstrates the load lifecycle).
        Document loaded = new Document(inputPath);

        // Configure FindReplaceOptions (no need to enable regex explicitly;
        // using the Regex overload of Replace automatically performs a regex replace).
        FindReplaceOptions options = new FindReplaceOptions();

        // Define the regex pattern to find.
        Regex regex = new Regex(@"Order \d+");

        // Perform the regex‑based replacement.
        int replacedCount = loaded.Range.Replace(regex, "Order ###", options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
