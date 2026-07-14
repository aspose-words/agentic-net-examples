using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplacementProgressEntry
{
    public string Pattern { get; set; } = "";
    public int Count { get; set; }
}

public class ProgressReporter
{
    private readonly string _filePath;
    private readonly List<ReplacementProgressEntry> _entries = new();

    public ProgressReporter(string filePath)
    {
        _filePath = filePath;
    }

    public void Report(string pattern, int count)
    {
        _entries.Add(new ReplacementProgressEntry { Pattern = pattern, Count = count });
        File.WriteAllText(_filePath, JsonConvert.SerializeObject(_entries, Formatting.Indented));
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";
        const string progressPath = "progress.json";

        // -----------------------------------------------------------------
        // 1. Create a sample document with placeholders to be replaced.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello PLACEHOLDER1, this is a test.");
        builder.Writeln("Another PLACEHOLDER2 appears.");
        builder.Writeln("PLACEHOLDER1 again.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Set up the progress reporter.
        // -----------------------------------------------------------------
        ProgressReporter reporter = new ProgressReporter(progressPath);

        // -----------------------------------------------------------------
        // 4. Define the replacement operations.
        // -----------------------------------------------------------------
        var replacements = new (string pattern, string replacement)[]
        {
            ("PLACEHOLDER1", "Alice"),
            ("PLACEHOLDER2", "Bob")
        };

        int totalReplacements = 0;

        // -----------------------------------------------------------------
        // 5. Perform each replacement and report the count.
        // -----------------------------------------------------------------
        foreach (var (pattern, replacement) in replacements)
        {
            FindReplaceOptions options = new FindReplaceOptions(); // default options
            int count = loaded.Range.Replace(pattern, replacement, options);
            totalReplacements += count;

            // Report the count for this specific pattern.
            reporter.Report(pattern, count);
        }

        // -----------------------------------------------------------------
        // 6. Validate that at least one replacement occurred.
        // -----------------------------------------------------------------
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // -----------------------------------------------------------------
        // 7. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);
    }
}
