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
        // Create a sample document with deprecated terms in different languages.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("The colour of the sky is blue.");                     // English (British)
        builder.Writeln("Our organisation follows strict guidelines.");        // English (British)
        builder.Writeln("Ce terme est déprécié et doit être remplacé.");       // French
        builder.Writeln("Un autre texte sans terme obsolète.");               // French
        doc.Save("sample.docx");

        // Load the document for processing.
        var loadedDoc = new Document("sample.docx");

        // Prepare replacement definitions: pattern, replacement, culture identifier.
        var replacements = new List<(Regex Pattern, string Replacement, string Culture)>
        {
            (new Regex(@"\bcolour\b", RegexOptions.IgnoreCase), "color", "en-GB"),
            (new Regex(@"\borganisation\b", RegexOptions.IgnoreCase), "organization", "en-GB"),
            (new Regex(@"\bdéprécié\b", RegexOptions.IgnoreCase), "obsolète", "fr-FR")
        };

        // Collect all log entries from each replacement operation.
        var allLogEntries = new List<ReplacementLogEntry>();

        foreach (var (pattern, replacement, culture) in replacements)
        {
            var logger = new ReplacementLogger(culture);
            var options = new FindReplaceOptions(logger);

            int count = loadedDoc.Range.Replace(pattern, replacement, options);
            if (count == 0)
                throw new InvalidOperationException($"Expected at least one replacement for culture '{culture}'.");

            allLogEntries.AddRange(logger.LogEntries);
        }

        // Save the modified document.
        loadedDoc.Save("output.docx");

        // Serialize the replacement log to JSON.
        string jsonReport = JsonConvert.SerializeObject(allLogEntries, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }
}

// Holds information about a single replacement operation.
public class ReplacementLogEntry
{
    public string Culture { get; set; } = string.Empty;
    public string OriginalText { get; set; } = string.Empty;
    public string ReplacementText { get; set; } = string.Empty;
    public int MatchOffset { get; set; }
    public string NodeType { get; set; } = string.Empty;
}

// Callback that records each replacement.
public class ReplacementLogger : IReplacingCallback
{
    public List<ReplacementLogEntry> LogEntries { get; } = new List<ReplacementLogEntry>();
    private readonly string _culture;

    public ReplacementLogger(string culture)
    {
        _culture = culture;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Record details before the replacement is applied.
        LogEntries.Add(new ReplacementLogEntry
        {
            Culture = _culture,
            OriginalText = args.Match.Value,
            ReplacementText = args.Replacement,
            MatchOffset = args.MatchOffset,
            NodeType = args.MatchNode.NodeType.ToString()
        });

        // Perform the replacement.
        return ReplaceAction.Replace;
    }
}
