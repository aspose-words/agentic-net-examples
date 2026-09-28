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
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "FindReplaceDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the input, output, and report files.
        string inputPath = Path.Combine(workDir, "input.docx");
        string outputPath = Path.Combine(workDir, "output.docx");
        string reportPath = Path.Combine(workDir, "report.json");

        // Build a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a Sample text.");
        builder.Writeln("Another sample line.");
        builder.Writeln("SAMPLE appears in uppercase.");
        builder.Writeln("An example without the target word.");
        doc.Save(inputPath);

        // Load the document for replacement.
        Document loaded = new Document(inputPath);

        // Configure find‑replace options: case‑insensitive, whole‑word only.
        // Use a regular expression with word boundaries to achieve whole‑word matching.
        Regex regex = new Regex(@"\bsample\b", RegexOptions.IgnoreCase);
        FindReplaceOptions options = new FindReplaceOptions();

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace(regex, "demo", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
        {
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");
        }

        // Save the modified document.
        loaded.Save(outputPath);

        // Prepare a simple report.
        var report = new
        {
            SearchTerm = "sample",
            Replacement = "demo",
            ReplacementsMade = replacedCount,
            InputFile = inputPath,
            OutputFile = outputPath
        };

        // Serialize the report to JSON and write it to disk.
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(reportPath, json);
    }
}
