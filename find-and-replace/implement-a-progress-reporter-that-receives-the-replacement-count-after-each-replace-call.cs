using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ProgressReporter
{
    public void Report(string oldValue, string newValue, int replacedCount)
    {
        Console.WriteLine($"Replaced \"{oldValue}\" with \"{newValue}\": {replacedCount} occurrence(s).");
    }
}

public class ReplacementResult
{
    public string OldValue { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;
    public int ReplacedCount { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Hello old value, this is an old example.");
        builder.Writeln("Replace old and old again.");
        builder.Writeln("Another line with old value.");

        // Save the initial document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Define replacements to perform.
        var replacements = new List<(string Old, string New)>
        {
            ("old", "new"),
            ("quick", "swift"),
            ("lazy", "sleepy")
        };

        var reporter = new ProgressReporter();
        var reportData = new List<ReplacementResult>();

        // Perform each replacement and report progress.
        foreach (var (oldValue, newValue) in replacements)
        {
            int count = doc.Range.Replace(oldValue, newValue, new FindReplaceOptions());
            reporter.Report(oldValue, newValue, count);

            // Record the result for later JSON report.
            reportData.Add(new ReplacementResult
            {
                OldValue = oldValue,
                NewValue = newValue,
                ReplacedCount = count
            });
        }

        // Validate that at least one replacement occurred overall.
        int totalReplacements = 0;
        foreach (var result in reportData)
            totalReplacements += result.ReplacedCount;

        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Write a JSON report of the replacements.
        const string jsonReportPath = "report.json";
        string json = JsonConvert.SerializeObject(reportData, Formatting.Indented);
        File.WriteAllText(jsonReportPath, json);

        // Verify that the report file was created.
        if (!File.Exists(jsonReportPath))
            throw new InvalidOperationException("Failed to create the JSON report file.");
    }
}
