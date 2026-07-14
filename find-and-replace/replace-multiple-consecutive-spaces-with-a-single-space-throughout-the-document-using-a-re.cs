using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with irregular spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is   a    test.  Multiple   spaces   are   present.");
        builder.Writeln("Another   line    with  extra   spaces.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches two or more consecutive spaces.
        Regex multipleSpaces = new Regex(@" {2,}");

        // Perform the replacement: each match is replaced with a single space.
        int replacementCount = loaded.Range.Replace(multipleSpaces, " ", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}
