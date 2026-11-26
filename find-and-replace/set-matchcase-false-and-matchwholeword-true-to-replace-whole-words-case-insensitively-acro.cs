using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with words in different cases.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World.");
        builder.Writeln("hello world.");
        builder.Writeln("HELLO WORLD.");
        builder.Writeln("This is a test of hello replacement.");

        // Configure find‑replace options: case‑insensitive, whole‑word match.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("hello", "Hi", options);

        // Validate that replacements occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Create a simple JSON report of the operation.
        var report = new
        {
            SearchTerm = "hello",
            Replacement = "Hi",
            MatchCase = options.MatchCase,
            FindWholeWordsOnly = options.FindWholeWordsOnly,
            ReplacementsMade = replacedCount,
            OutputDocument = outputPath
        };

        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }
}
