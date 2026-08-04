using System;
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
        builder.Writeln("Sample dates: 12-31-2020, 01-15-2021, 07-04-2022.");
        doc.Save("input.docx");

        // Load the document we just created.
        Document loaded = new Document("input.docx");

        // Define a regular expression that captures month, day and year.
        Regex datePattern = new Regex(@"(\d{2})-(\d{2})-(\d{4})");

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
        loaded.Save("output.docx");

        // Output the resulting text to the console (optional verification).
        Console.WriteLine("Replaced text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}
