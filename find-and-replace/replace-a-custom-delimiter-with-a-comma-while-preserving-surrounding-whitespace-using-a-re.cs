using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample text containing a custom delimiter ';' with varying whitespace.
        builder.Writeln("Apple ; Banana ;  Cherry ;Date");
        builder.Writeln("One;Two ;Three ; Four");

        // Regular expression that captures any whitespace before and after the delimiter.
        // Group 1: whitespace before ';'
        // Group 2: whitespace after ';'
        Regex delimiterPattern = new Regex(@"(\s*);(\s*)");

        // Replace the delimiter with a comma while preserving the captured whitespace.
        int replacementCount = doc.Range.Replace(delimiterPattern, "$1,$2");

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No delimiters were replaced.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "Result.docx");
        doc.Save(outputPath);

        // Optional: write a simple confirmation to the console.
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
