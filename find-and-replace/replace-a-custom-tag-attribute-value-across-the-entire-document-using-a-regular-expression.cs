using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by the package list, not used directly
using Newtonsoft.Json; // Required by the package list, not used directly

public class ReplaceCustomTagAttribute
{
    public static void Main()
    {
        // Create a sample document with custom tag attributes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<mytag attr=\"value1\">First content</mytag>");
        builder.Writeln("<mytag attr=\"value2\">Second content</mytag>");
        builder.Writeln("<mytag attr=\"value3\">Third content</mytag>");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches the attribute value.
        Regex regex = new Regex(@"attr=""[^""]*""", RegexOptions.Compiled);

        // Replace all attribute values with a new value.
        const string replacement = @"attr=""newvalue""";
        int replacedCount = loaded.Range.Replace(regex, replacement, new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one attribute replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optionally, report the result.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {Path.GetFullPath(outputPath)}");
    }
}
