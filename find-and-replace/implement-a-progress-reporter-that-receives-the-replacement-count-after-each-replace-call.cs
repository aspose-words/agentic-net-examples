using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplacementInfo
{
    public int Step { get; set; }
    public string OldValue { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;
    public int Count { get; set; }
}

public class ProgressReporter
{
    private readonly List<ReplacementInfo> _replacements = new List<ReplacementInfo>();

    public void Report(string oldValue, string newValue, int count, int step, int totalSteps)
    {
        Console.WriteLine($"Step {step}/{totalSteps}: Replaced \"{oldValue}\" with \"{newValue}\" – {count} occurrence(s).");
        _replacements.Add(new ReplacementInfo
        {
            Step = step,
            OldValue = oldValue,
            NewValue = newValue,
            Count = count
        });
    }

    public void SaveReport(string filePath)
    {
        string json = JsonConvert.SerializeObject(_replacements, Formatting.Indented);
        File.WriteAllText(filePath, json);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("The quick brown fox is quick.");
        builder.Writeln("Lazy dog sleeps.");
        string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define replacements.
        var replacements = new List<(string OldValue, string NewValue)>
        {
            ("quick", "swift"),
            ("lazy", "sleepy"),
            ("fox", "wolf")
        };

        // Initialize progress reporter.
        ProgressReporter reporter = new ProgressReporter();
        int totalSteps = replacements.Count;

        // Perform replacements with progress reporting.
        for (int i = 0; i < totalSteps; i++)
        {
            var (oldValue, newValue) = replacements[i];
            int replacedCount = loadedDoc.Range.Replace(oldValue, newValue, new FindReplaceOptions());
            reporter.Report(oldValue, newValue, replacedCount, i + 1, totalSteps);
        }

        // Validate that at least one replacement occurred.
        if (reporter == null)
        {
            throw new InvalidOperationException("Progress reporter was not initialized.");
        }

        // Save the modified document.
        string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Save a JSON report of the replacement counts.
        string reportPath = "report.json";
        reporter.SaveReport(reportPath);
    }
}
