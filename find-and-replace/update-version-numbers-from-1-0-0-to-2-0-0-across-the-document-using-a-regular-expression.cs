using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing version numbers "1.0.0".
        Document sample = new Document();
        DocumentBuilder builder = new DocumentBuilder(sample);
        builder.Writeln("Product release notes:");
        builder.Writeln("Current version: 1.0.0");
        builder.Writeln("Previous version was 1.0.0, now updated.");
        // Save the sample document to a local file.
        const string inputPath = "input.docx";
        sample.Save(inputPath);

        // Load the document we just created.
        Document doc = new Document(inputPath);

        // Define a regular expression that matches the exact version string "1.0.0".
        Regex versionPattern = new Regex(@"\b1\.0\.0\b", RegexOptions.Compiled);

        // Perform the replacement using Aspose.Words Range.Replace with a regex.
        int replaceCount = doc.Range.Replace(versionPattern, "2.0.0", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replaceCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement, but none were found.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
