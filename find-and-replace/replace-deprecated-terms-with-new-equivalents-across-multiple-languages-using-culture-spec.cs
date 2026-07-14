using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with deprecated terms in several languages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The old term is obsolete.");                     // English
        builder.Writeln("Le terme ancien est obsolète.");                // French
        builder.Writeln("Der alte Begriff ist veraltet.");               // German
        builder.Writeln("旧术语已过时。");                                 // Chinese (no replacement needed)

        // Define replacement rules per language using regular expressions.
        var rules = new List<ReplacementRule>
        {
            new ReplacementRule
            {
                Language = "English",
                Pattern = new Regex(@"\bold\b", RegexOptions.IgnoreCase),
                Replacement = "new"
            },
            new ReplacementRule
            {
                Language = "French",
                Pattern = new Regex(@"ancien", RegexOptions.IgnoreCase),
                Replacement = "nouveau"
            },
            new ReplacementRule
            {
                Language = "German",
                Pattern = new Regex(@"alte", RegexOptions.IgnoreCase),
                Replacement = "neue"
            }
        };

        // Collect all replacement log entries.
        var allLogEntries = new List<ReplacementLogEntry>();
        int totalReplacements = 0;

        foreach (var rule in rules)
        {
            var logger = new ReplacementLogger(rule.Language);
            var options = new FindReplaceOptions
            {
                ReplacingCallback = logger
            };

            int replaced = doc.Range.Replace(rule.Pattern, rule.Replacement, options);
            totalReplacements += replaced;

            allLogEntries.AddRange(logger.Entries);
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputDocPath = "output.docx";
        doc.Save(outputDocPath);

        // Serialize the replacement log to JSON.
        const string jsonReportPath = "replacements.json";
        string json = JsonConvert.SerializeObject(allLogEntries, Formatting.Indented);
        File.WriteAllText(jsonReportPath, json);

        // Simple verification that output files exist.
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);
        if (!File.Exists(jsonReportPath))
            throw new FileNotFoundException("The JSON report was not created.", jsonReportPath);
    }
}

// Represents a single replacement rule.
public class ReplacementRule
{
    public string Language { get; set; } = string.Empty;
    public Regex Pattern { get; set; } = new Regex(string.Empty);
    public string Replacement { get; set; } = string.Empty;
}

// Holds information about a single replacement occurrence.
public class ReplacementLogEntry
{
    public string Language { get; set; } = string.Empty;
    public string Original { get; set; } = string.Empty;
    public string Replacement { get; set; } = string.Empty;
    public int MatchOffset { get; set; }
    public string NodeType { get; set; } = string.Empty;
}

// Callback that records each replacement made during a find‑replace operation.
public class ReplacementLogger : IReplacingCallback
{
    public string Language { get; }
    public List<ReplacementLogEntry> Entries { get; } = new List<ReplacementLogEntry>();

    public ReplacementLogger(string language)
    {
        Language = language ?? string.Empty;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        var entry = new ReplacementLogEntry
        {
            Language = Language,
            Original = args.Match.Value,
            Replacement = args.Replacement,
            MatchOffset = args.MatchOffset,
            NodeType = args.MatchNode.NodeType.ToString()
        };
        Entries.Add(entry);
        return ReplaceAction.Replace;
    }
}
