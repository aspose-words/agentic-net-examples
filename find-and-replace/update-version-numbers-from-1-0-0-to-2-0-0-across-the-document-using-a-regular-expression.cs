using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        // Create a sample document with version numbers.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);
        builder.Writeln("Application version: 1.0.0");
        builder.Writeln("Release notes for version 1.0.0 are available.");
        builder.Writeln("Previous version was 0.9.5.");
        string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // Load the document to perform replacement.
        Document doc = new Document(inputPath);

        // Define a regex that matches the exact version "1.0.0".
        Regex versionRegex = new Regex(@"\b1\.0\.0\b");

        // Replace all occurrences with the new version "2.0.0".
        int replacedCount = doc.Range.Replace(versionRegex, "2.0.0", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one version number replacement.");

        // Save the updated document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
