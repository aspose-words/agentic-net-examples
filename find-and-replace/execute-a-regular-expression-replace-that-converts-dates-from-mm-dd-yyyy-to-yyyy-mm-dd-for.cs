using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing dates in MM-DD-YYYY format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The project started on 03-15-2022 and ended on 12-31-2022.");
        builder.Writeln("Another date: 01-01-2023.");

        // Save the original document (optional, for verification).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Define a regular expression that captures month, day, and year.
        Regex dateRegex = new Regex(@"\b(\d{2})-(\d{2})-(\d{4})\b");

        // Replacement pattern reorders the captured groups to YYYY-MM-DD.
        string replacementPattern = "$3-$1-$2";

        // Perform the regex replace across the whole document.
        int replacedCount = doc.Range.Replace(dateRegex, replacementPattern, new FindReplaceOptions());

        // Ensure that at least one date was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one date replacement, but none occurred.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
