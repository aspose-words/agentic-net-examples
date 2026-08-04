using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document with irregular spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is   a    test.  Multiple   spaces   here.");
        builder.Writeln("Another   line   with  spaces.");
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Define a regular expression that matches two or more consecutive spaces.
        Regex multipleSpaces = new Regex(@" {2,}");

        // Perform the replacement: replace each match with a single space.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacementCount = loaded.Range.Replace(multipleSpaces, " ", options);

        // Ensure that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one space reduction replacement.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Output the result count for verification (no interactive input required).
        Console.WriteLine($"Replaced {replacementCount} occurrence(s) of multiple spaces.");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }
}
