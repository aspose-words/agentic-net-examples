using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple consecutive spaces.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is   a    test document.  It contains   irregular   spacing.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches two or more spaces.
        Regex regex = new Regex(@" {2,}");

        // Perform the replacement: replace matches with a single space.
        int replacedCount = loaded.Range.Replace(regex, " ", new FindReplaceOptions());

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
