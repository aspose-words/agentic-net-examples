using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // Create a sample document containing version numbers.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Version 1.0.0 of the product.");
        builder.Writeln("Another reference: 1.0.0.");
        builder.Writeln("Do not change 1.0.1."); // Should remain unchanged.
        doc.Save(inputPath);

        // -------------------------------------------------
        // Load the document and perform a regex replace.
        // -------------------------------------------------
        Document loaded = new Document(inputPath);

        // Regex that matches the exact version string "1.0.0".
        Regex versionRegex = new Regex(@"\b1\.0\.0\b");

        // Use default FindReplaceOptions.
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace all occurrences of "1.0.0" with "2.0.0".
        int replacedCount = loaded.Range.Replace(versionRegex, "2.0.0", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement.");

        // -------------------------------------------------
        // Save the modified document.
        // -------------------------------------------------
        loaded.Save(outputPath);
    }
}
