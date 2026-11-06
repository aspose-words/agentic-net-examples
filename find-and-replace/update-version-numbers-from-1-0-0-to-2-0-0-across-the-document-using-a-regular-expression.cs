using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with version numbers.
        var inputPath = "input.docx";
        var outputPath = "output.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is version 1.0.0 of the product.");
        builder.Writeln("Previous releases: 0.9.5, 1.0.0, 1.0.1.");
        builder.Writeln("Another reference: v1.0.0.");
        doc.Save(inputPath);

        // Load the document from the file.
        var loadedDoc = new Document(inputPath);

        // Define a regular expression that matches the exact version string "1.0.0".
        var versionRegex = new Regex(@"\b1\.0\.0\b");

        // Perform the replacement using Range.Replace with the regex.
        var replaceOptions = new FindReplaceOptions();
        int replacedCount = loadedDoc.Range.Replace(versionRegex, "2.0.0", replaceOptions);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement, but none were made.");

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}
