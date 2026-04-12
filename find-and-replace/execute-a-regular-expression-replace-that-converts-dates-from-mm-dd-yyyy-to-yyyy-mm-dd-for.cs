using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample paragraphs containing dates in MM-DD-YYYY format.
        builder.Writeln("The project started on 03-15-2022 and ended on 12-31-2022.");
        builder.Writeln("Another milestone was on 07-04-2021.");

        // Define a regular expression that captures month, day, and year.
        Regex datePattern = new Regex(@"(\d{2})-(\d{2})-(\d{4})");

        // Configure find‑replace options to enable substitution groups.
        FindReplaceOptions options = new FindReplaceOptions
        {
            UseSubstitutions = true // Allows $1, $2, $3 in the replacement string.
        };

        // Perform the replacement: MM-DD-YYYY → YYYY-MM-DD.
        int replacedCount = doc.Range.Replace(datePattern, "$3-$1-$2", options);

        // Validate that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("No dates were replaced.");

        // Save the modified document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Optional: write a short confirmation to the console.
        Console.WriteLine($"Replaced {replacedCount} date(s). Output saved to: {outputPath}");
    }
}
