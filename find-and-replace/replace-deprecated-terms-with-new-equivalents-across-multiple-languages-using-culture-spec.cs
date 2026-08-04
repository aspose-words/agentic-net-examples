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
        // Create a sample document containing deprecated terms in several languages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("DeprecatedTerm1 is old.");
        builder.Writeln("DeprecatedTerm2 is outdated.");
        builder.Writeln("Ancien terme en français.");
        builder.Writeln("Veralteter Begriff auf Deutsch.");

        // Prepare a logger that will record every replacement performed.
        var logger = new ReplacementLogger();

        // Common FindReplaceOptions used for all replacements.
        var options = new FindReplaceOptions
        {
            MatchCase = false,               // Case‑insensitive search.
            FindWholeWordsOnly = true,       // Replace whole words only.
            ReplacingCallback = logger
        };

        // Define the culture‑specific patterns and their replacements.
        var replacements = new List<(Regex Pattern, string Replacement)>
        {
            // English terms.
            (new Regex(@"\bDeprecatedTerm1\b", RegexOptions.Compiled | RegexOptions.CultureInvariant), "NewTerm1"),
            (new Regex(@"\bDeprecatedTerm2\b", RegexOptions.Compiled | RegexOptions.CultureInvariant), "NewTerm2"),
            // French term (case‑insensitive, Unicode aware).
            (new Regex(@"\bAncien\b", RegexOptions.Compiled | RegexOptions.CultureInvariant), "Nouveau"),
            // German term.
            (new Regex(@"\bVeralteter\b", RegexOptions.Compiled | RegexOptions.CultureInvariant), "Aktuell")
        };

        // Apply each replacement to the document.
        int totalReplacements = 0;
        foreach (var (pattern, replacement) in replacements)
        {
            int count = doc.Range.Replace(pattern, replacement, options);
            totalReplacements += count;
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputDocPath = Path.Combine(outputDir, "ReplacedDocument.docx");
        doc.Save(outputDocPath);

        // Serialize the replacement log to JSON.
        string jsonReport = JsonConvert.SerializeObject(logger.Records, Formatting.Indented);
        string jsonPath = Path.Combine(outputDir, "ReplacementReport.json");
        File.WriteAllText(jsonPath, jsonReport);
    }

    // Simple record to hold details of each replacement.
    private class ReplacementRecord
    {
        public string Original { get; set; } = string.Empty;
        public string Replacement { get; set; } = string.Empty;
    }

    // Callback that logs every match that is replaced.
    private class ReplacementLogger : IReplacingCallback
    {
        public List<ReplacementRecord> Records { get; } = new List<ReplacementRecord>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            Records.Add(new ReplacementRecord
            {
                Original = args.Match.Value,
                Replacement = args.Replacement
            });
            // No modification of the replacement string; just perform the replace.
            return ReplaceAction.Replace;
        }
    }
}
