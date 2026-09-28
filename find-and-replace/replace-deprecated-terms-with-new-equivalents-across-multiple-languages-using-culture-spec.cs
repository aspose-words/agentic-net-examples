using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplacementLog
{
    public string Culture { get; set; } = string.Empty;
    public string OldValue { get; set; } = string.Empty;
    public string NewValue { get; set; } = string.Empty;
}

public class ReplacementCallback : IReplacingCallback
{
    private readonly string _culture;
    private readonly string _newValue;
    private readonly List<ReplacementLog> _logs;

    public ReplacementCallback(string culture, string newValue, List<ReplacementLog> logs)
    {
        _culture = culture;
        _newValue = newValue;
        _logs = logs;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        _logs.Add(new ReplacementLog
        {
            Culture = _culture,
            OldValue = args.Match.Value,
            NewValue = _newValue
        });
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with deprecated terms in different languages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The old term is 'foo' in English.");
        builder.Writeln("Le terme ancien est 'ancien' en français.");
        builder.Writeln("Das alte Wort ist 'alt' auf Deutsch.");
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Define culture‑specific replacement rules.
        var rules = new List<(string Culture, Regex Pattern, string Replacement)>
        {
            ("en", new Regex(@"\bfoo\b", RegexOptions.IgnoreCase), "bar"),
            ("fr", new Regex(@"\bancien\b", RegexOptions.IgnoreCase), "nouveau"),
            ("de", new Regex(@"\balt\b", RegexOptions.IgnoreCase), "neu")
        };

        // Collect logs of all replacements.
        var allLogs = new List<ReplacementLog>();

        // Apply each rule using a callback to record matches.
        foreach (var rule in rules)
        {
            var callback = new ReplacementCallback(rule.Culture, rule.Replacement, allLogs);
            var options = new FindReplaceOptions { ReplacingCallback = callback };
            int replacedCount = loadedDoc.Range.Replace(rule.Pattern, rule.Replacement, options);
            // Optional: you could validate per‑rule replacement count here.
        }

        // Ensure at least one replacement occurred.
        if (allLogs.Count == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Serialize the replacement log to JSON.
        string jsonReport = JsonConvert.SerializeObject(allLogs, Formatting.Indented);
        const string reportPath = "replacements.json";
        File.WriteAllText(reportPath, jsonReport);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("The replacement report was not created.");
    }
}
