using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing a custom delimiter '|'.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Alpha|Beta");
        builder.Writeln("Gamma | Delta");
        builder.Writeln("Epsilon   |   Zeta");
        builder.Writeln("NoDelimiterHere");

        // Define a regular expression that captures any whitespace before and after the delimiter.
        // Group 1 = whitespace before, Group 2 = whitespace after.
        Regex delimiterRegex = new Regex(@"(\s*)\|(\s*)");

        // Replace the delimiter with a comma while preserving the captured whitespace.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = doc.Range.Replace(delimiterRegex, "$1,$2", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one delimiter replacement.");

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: Write a simple JSON report of the operation.
        var report = new { ReplacementsMade = replacedCount, OutputFile = Path.GetFullPath(outputPath) };
        File.WriteAllText("report.json", JsonConvert.SerializeObject(report, Formatting.Indented));
    }
}
