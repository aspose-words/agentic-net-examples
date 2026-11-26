using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;   // Needed for FindReplaceOptions

public class Program
{
    public static void Main()
    {
        // Create a sample document containing dates in MM-DD-YYYY format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample dates: 12-31-2020, 01-15-2021, and 07-04-2022.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Regular expression that captures month, day, and year parts.
        Regex dateRegex = new Regex(@"(\d{2})-(\d{2})-(\d{4})");

        // Replace each match with the format YYYY-MM-DD.
        int replacedCount = loaded.Range.Replace(dateRegex, "$3-$1-$2", new FindReplaceOptions());

        // Ensure that at least one replacement was performed.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("No date patterns were replaced.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Confirmation output.
        Console.WriteLine($"Replaced {replacedCount} date(s). Output saved to '{outputPath}'.");
    }
}
