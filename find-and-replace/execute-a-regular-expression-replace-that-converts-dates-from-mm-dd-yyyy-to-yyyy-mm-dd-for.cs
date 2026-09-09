using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with dates in MM-DD-YYYY format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample dates:");
        builder.Writeln("12-31-2020");
        builder.Writeln("01-01-2021");
        builder.Writeln("07-04-2022");

        // Save the source document.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regular expression that captures month, day, and year.
        Regex datePattern = new Regex(@"(\d{2})-(\d{2})-(\d{4})");

        // Replace matches with the format YYYY-MM-DD.
        // $3 = year, $1 = month, $2 = day.
        int replacedCount = loaded.Range.Replace(datePattern, "$3-$1-$2", new FindReplaceOptions());

        // Verify that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement, but none were made.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loaded.Save(outputPath);

        // Optional: output the number of replacements to the console.
        Console.WriteLine($"Replaced {replacedCount} date(s). Output saved to: {outputPath}");
    }
}
