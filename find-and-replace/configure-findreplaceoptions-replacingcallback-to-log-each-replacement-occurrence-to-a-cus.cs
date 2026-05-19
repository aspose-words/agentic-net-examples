using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma alpha");

        // Set up the custom logger.
        ReplacementLogger logger = new ReplacementLogger();

        // Configure FindReplaceOptions with the logger callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("alpha", "omega", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        doc.Save("output.docx");

        // Write the log to plain text and JSON files.
        File.WriteAllText("log.txt", logger.GetPlainLog());
        File.WriteAllText("log.json", logger.GetJsonLog());

        // Output summary to console.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine("Log written to log.txt and log.json.");
    }
}

// Custom logger that records each replacement occurrence.
public class ReplacementLogger : IReplacingCallback
{
    private readonly List<ReplacementRecord> _records = new List<ReplacementRecord>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Record the original match and the replacement text.
        _records.Add(new ReplacementRecord
        {
            Original = args.Match.Value,
            Replacement = args.Replacement,
            MatchOffset = args.MatchOffset,
            NodeType = args.MatchNode.NodeType.ToString()
        });

        // Continue with the default replacement.
        return ReplaceAction.Replace;
    }

    // Returns a plain‑text representation of the log.
    public string GetPlainLog()
    {
        StringBuilder sb = new StringBuilder();
        foreach (var r in _records)
        {
            sb.AppendLine($"\"{r.Original}\" => \"{r.Replacement}\" at offset {r.MatchOffset} in {r.NodeType} node.");
        }
        return sb.ToString();
    }

    // Returns a JSON representation of the log.
    public string GetJsonLog()
    {
        return JsonConvert.SerializeObject(_records, Formatting.Indented);
    }

    // Simple DTO for serialization.
    private class ReplacementRecord
    {
        public string Original { get; set; } = string.Empty;
        public string Replacement { get; set; } = string.Empty;
        public int MatchOffset { get; set; }
        public string NodeType { get; set; } = string.Empty;
    }
}
