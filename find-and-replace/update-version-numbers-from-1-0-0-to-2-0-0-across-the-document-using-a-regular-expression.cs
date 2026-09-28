using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing version numbers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Product A version 1.0.0 released.");
        builder.Writeln("Product B version 1.0.0 is now deprecated.");
        builder.Writeln("No version here.");
        builder.Writeln("Another reference: 1.0.0.");

        // Save the sample input (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document to simulate a typical workflow.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the exact version string "1.0.0".
        Regex versionRegex = new Regex(@"\b1\.0\.0\b", RegexOptions.Compiled);

        // Perform the replacement using Aspose.Words Range.Replace with regex.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = loaded.Range.Replace(versionRegex, "2.0.0", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement, but none were made.");

        // Save the updated document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
