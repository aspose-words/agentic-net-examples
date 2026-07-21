using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        // Create a sample document with dates in MM-DD-YYYY format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Meeting dates:");
        builder.Writeln("12-31-2020");
        builder.Writeln("01-01-2021");
        builder.Writeln("Invalid date 13-40-2020 should stay unchanged.");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Define a regex that captures month, day, and year.
        Regex datePattern = new Regex(@"\b(\d{2})-(\d{2})-(\d{4})\b");

        // Configure replace options to enable substitution groups ($1, $2, $3).
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true,
            LegacyMode = false
        };

        // Perform the replacement: MM-DD-YYYY -> YYYY-MM-DD.
        int replacedCount = loaded.Range.Replace(datePattern, "$3-$1-$2", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Optional: display the resulting text in the console.
        Console.WriteLine("Replacements performed: " + replacedCount);
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}
