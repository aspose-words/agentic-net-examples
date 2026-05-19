using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ReplaceMultipleSpaces
{
    public static void Main()
    {
        // Create a sample document with irregular spacing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This  is   a    test.  Multiple   spaces   here.");
        builder.Writeln("Another   line    with  extra   spaces.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document we just created.
        Document loaded = new Document(inputPath);

        // Define a regex that matches two or more consecutive space characters.
        Regex multipleSpaces = new Regex(@" {2,}");

        // Perform the replacement: each match is replaced with a single space.
        int replacementCount = loaded.Range.Replace(multipleSpaces, " ", new FindReplaceOptions());

        // Ensure that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one space reduction.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: indicate success.
        Console.WriteLine($"Replaced {replacementCount} occurrences. Output saved to '{outputPath}'.");
    }
}
