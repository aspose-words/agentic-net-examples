using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Use Aspose.Drawing namespace for drawing-related types
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        var inputPath = "input.docx";
        var outputPath = "output.docx";
        var reportPath = "replacementReport.json";

        // Create a sample document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta gamma");
        builder.Writeln("alpha appears twice: alpha.");
        builder.Writeln("beta will become delta.");
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Prepare the progress reporter.
        var reporter = new ReplacementProgressReporter();

        // Define the replacements to perform.
        var replacements = new[]
        {
            new ReplacementPair("alpha", "omega"),
            new ReplacementPair("beta", "delta"),
            new ReplacementPair("gamma", "theta")
        };

        int totalReplacements = 0;

        // Perform each replacement and report the count.
        foreach (var pair in replacements)
        {
            var options = new FindReplaceOptions(); // Default options.
            int count = loadedDoc.Range.Replace(pair.Find, pair.Replace, options);
            totalReplacements += count;
            reporter.Report(pair.Find, pair.Replace, count);
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        loadedDoc.Save(outputPath);

        // Write the replacement report to a JSON file.
        reporter.SaveReport(reportPath);
    }
}

// Simple data holder for a find/replace pair.
public class ReplacementPair
{
    public string Find { get; }
    public string Replace { get; }

    public ReplacementPair(string find, string replace)
    {
        Find = find ?? throw new ArgumentNullException(nameof(find));
        Replace = replace ?? throw new ArgumentNullException(nameof(replace));
    }
}

// Holds information about a single replacement operation.
public class ReplacementInfo
{
    public string Find { get; set; } = string.Empty;
    public string Replace { get; set; } = string.Empty;
    public int Count { get; set; }
}

// Collects replacement results and writes a JSON report.
public class ReplacementProgressReporter
{
    private readonly List<ReplacementInfo> _records = new List<ReplacementInfo>();

    public void Report(string find, string replace, int count)
    {
        _records.Add(new ReplacementInfo { Find = find, Replace = replace, Count = count });
        Console.WriteLine($"Replaced \"{find}\" with \"{replace}\" {count} time(s).");
    }

    public void SaveReport(string filePath)
    {
        var json = JsonConvert.SerializeObject(_records, Formatting.Indented);
        File.WriteAllText(filePath, json);
    }
}
