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
        builder.Writeln("Meeting dates: 12-31-2020, 01-15-2021, and 07-04-2022.");
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Define a regular expression that captures month, day, and year.
        Regex datePattern = new Regex(@"(\d{2})-(\d{2})-(\d{4})");

        // Configure replace options to enable substitution groups.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true,
            LegacyMode = false
        };

        // Replace each match with YYYY-MM-DD using captured groups.
        int replacedCount = loaded.Range.Replace(datePattern, "$3-$1-$2", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement.");

        // Save the modified document.
        loaded.Save("output.docx");

        // Optional: display the resulting text in the console.
        Console.WriteLine("Replacements performed: " + replacedCount);
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(loaded.GetText().Trim());
    }
}
