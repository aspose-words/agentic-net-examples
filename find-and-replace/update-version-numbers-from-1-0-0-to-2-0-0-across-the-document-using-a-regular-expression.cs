using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text containing the version number to be replaced.
        builder.Writeln("Product release notes:");
        builder.Writeln("Current version: 1.0.0");
        builder.Writeln("Previous version was 0.9.5");
        builder.Writeln("Upgrade path: 1.0.0 -> 2.0.0");
        builder.Writeln("Another reference to version 1.0.0 in the footer.");

        // Define a regular expression that matches the exact version string "1.0.0".
        Regex versionPattern = new Regex(@"\b1\.0\.0\b", RegexOptions.Compiled);

        // Perform the replacement across the whole document.
        int replacementCount = doc.Range.Replace(versionPattern, "2.0.0");

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No occurrences of version 1.0.0 were found to replace.");

        // Save the modified document to a local file.
        const string outputPath = "UpdatedDocument.docx";
        doc.Save(outputPath);
    }
}
