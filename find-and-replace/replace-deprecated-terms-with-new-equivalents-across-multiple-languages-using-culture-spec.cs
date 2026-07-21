using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required package, not used directly
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document containing deprecated terms in several languages.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("The colour of the sky is blue.");
        builder.Writeln("Die Farbe des Himmels ist blau.");
        builder.Writeln("La couleur du ciel est bleue.");
        doc.Save("input.docx");

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        var loadedDoc = new Document("input.docx");

        // -----------------------------------------------------------------
        // 3. Define regex patterns for each language (culture‑specific, case‑insensitive).
        // -----------------------------------------------------------------
        var replacementPatterns = new List<(Regex Pattern, string Replacement)>
        {
            (new Regex(@"\bcolour\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant), "color"),
            (new Regex(@"\bFarbe\b",   RegexOptions.IgnoreCase | RegexOptions.CultureInvariant), "color"),
            (new Regex(@"\bcouleur\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant), "color")
        };

        // -----------------------------------------------------------------
        // 4. Set up a callback to log each replacement.
        // -----------------------------------------------------------------
        var logger = new ReplacementLogger();
        var options = new FindReplaceOptions(logger)
        {
            MatchCase = false // Ensure case‑insensitive matching.
        };

        // -----------------------------------------------------------------
        // 5. Perform replacements and count them.
        // -----------------------------------------------------------------
        int totalReplacements = 0;
        foreach (var (pattern, replacement) in replacementPatterns)
        {
            totalReplacements += loadedDoc.Range.Replace(pattern, replacement, options);
        }

        // -----------------------------------------------------------------
        // 6. Validate that at least one replacement occurred.
        // -----------------------------------------------------------------
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // -----------------------------------------------------------------
        // 7. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save("output.docx");

        // -----------------------------------------------------------------
        // 8. Build a simple report object and serialize it to JSON.
        // -----------------------------------------------------------------
        var report = new
        {
            TotalReplacements = totalReplacements,
            Details = logger.Replacements
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);
    }

    // -----------------------------------------------------------------
    // Callback implementation that records each replacement.
    // -----------------------------------------------------------------
    private class ReplacementLogger : IReplacingCallback
    {
        public List<ReplacementInfo> Replacements { get; } = new List<ReplacementInfo>();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Record original match and the replacement text.
            Replacements.Add(new ReplacementInfo
            {
                OriginalText   = args.Match.Value,
                ReplacementText = args.Replacement,
                MatchOffset    = args.MatchOffset,
                NodeType       = args.MatchNode.NodeType.ToString()
            });

            // Proceed with the default replacement.
            return ReplaceAction.Replace;
        }
    }

    // -----------------------------------------------------------------
    // Simple DTO for JSON serialization.
    // -----------------------------------------------------------------
    private class ReplacementInfo
    {
        public string OriginalText    { get; set; } = string.Empty;
        public string ReplacementText { get; set; } = string.Empty;
        public int    MatchOffset     { get; set; }
        public string NodeType        { get; set; } = string.Empty;
    }
}
