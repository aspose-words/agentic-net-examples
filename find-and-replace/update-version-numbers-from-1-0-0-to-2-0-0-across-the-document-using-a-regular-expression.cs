using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the sample files.
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workFolder);

        // Create a sample document containing version numbers "1.0.0".
        string inputPath = Path.Combine(workFolder, "input.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Application version: 1.0.0");
        builder.Writeln("Previous release: 1.0.0");
        sampleDoc.Save(inputPath);

        // Load the document we just created.
        Document doc = new Document(inputPath);

        // Define a regular expression that matches the exact version string "1.0.0".
        Regex versionRegex = new Regex(@"\b1\.0\.0\b", RegexOptions.Compiled);

        // Perform the replacement using Range.Replace with a regex pattern.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(versionRegex, "2.0.0", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No version numbers were replaced.");

        // Save the updated document.
        string outputPath = Path.Combine(workFolder, "output.docx");
        doc.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replaced {replacedCount} occurrence(s). Output saved to: {outputPath}");
    }
}
