using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class ReplacementInfo
{
    public string Language { get; set; } = string.Empty;
    public string Original { get; set; } = string.Empty;
    public string Replacement { get; set; } = string.Empty;
}

public class ReplacementLogger : IReplacingCallback
{
    private readonly string _language;
    private readonly string _replacement;
    private readonly List<ReplacementInfo> _log;

    public ReplacementLogger(string language, string replacement, List<ReplacementInfo> log)
    {
        _language = language;
        _replacement = replacement;
        _log = log;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Log the original match and the intended replacement.
        _log.Add(new ReplacementInfo
        {
            Language = _language,
            Original = args.Match.Value,
            Replacement = _replacement
        });

        // Perform the replacement with the provided replacement string.
        args.Replacement = _replacement;
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a list to collect replacement details for reporting.
        var replacementLog = new List<ReplacementInfo>();

        // Create a sample document containing deprecated terms in different languages.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("The old term 'foobar' is deprecated.");               // English
        builder.Writeln("Le terme ancien 'ancien' est obsolète.");            // French
        builder.Writeln("Der alte Begriff 'alt' ist veraltet.");              // German
        builder.Writeln("El término viejo 'viejo' está obsoleto.");           // Spanish (extra example)

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Define culture‑specific patterns and their replacements.
        var replacements = new[]
        {
            new { Language = "English", Pattern = new Regex(@"\bfoobar\b", RegexOptions.IgnoreCase), Replacement = "newfoo" },
            new { Language = "French",  Pattern = new Regex(@"\bancien\b", RegexOptions.IgnoreCase), Replacement = "nouveau" },
            new { Language = "German",  Pattern = new Regex(@"\balt\b",    RegexOptions.IgnoreCase), Replacement = "neu" },
            new { Language = "Spanish", Pattern = new Regex(@"\bviejo\b",  RegexOptions.IgnoreCase), Replacement = "nuevo" }
        };

        int totalReplacements = 0;

        // Perform replacements using regex and a callback to log each match.
        foreach (var item in replacements)
        {
            var options = new FindReplaceOptions();
            options.ReplacingCallback = new ReplacementLogger(item.Language, item.Replacement, replacementLog);
            int replaced = loadedDoc.Range.Replace(item.Pattern, item.Replacement, options);
            totalReplacements += replaced;
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Serialize the replacement log to JSON.
        const string reportPath = "report.json";
        string json = JsonConvert.SerializeObject(replacementLog, Formatting.Indented);
        File.WriteAllText(reportPath, json);

        // Verify that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");
    }
}
