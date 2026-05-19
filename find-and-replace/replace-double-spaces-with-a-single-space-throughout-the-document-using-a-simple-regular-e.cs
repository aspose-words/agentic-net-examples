using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with double spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is  a  sample  text  with  double  spaces.");
        builder.Writeln("Another  line  with  double  spaces  here.");
        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file system.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches two or more consecutive spaces.
        Regex doubleSpaceRegex = new Regex(@" {2,}");

        // Perform the replacement: replace any occurrence of double (or more) spaces with a single space.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementCount = loaded.Range.Replace(doubleSpaceRegex, " ", options);

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one double‑space replacement, but none were found.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replaced {replacementCount} double‑space occurrences. Output saved to '{outputPath}'.");
    }
}
