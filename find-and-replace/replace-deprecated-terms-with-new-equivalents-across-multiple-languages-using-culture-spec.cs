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
        // Create a blank document and add sample paragraphs containing deprecated terms.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("The colour of the sky is blue.");                     // British English
        builder.Writeln("Please update the organisation's policy.");          // British English
        builder.Writeln("Le programme est disponible.");                     // French
        builder.Writeln("Le département de recherche.");                     // French with accent
        builder.Writeln("Die Straße ist lang.");                              // German with ß
        builder.Writeln("Das ist ein gutes Beispiel.");                      // German (no change)
        builder.Writeln("La canción es popular.");                            // Spanish with accent

        // Save the original document (optional, for inspection).
        doc.Save("original.docx");

        // Prepare a logger to capture details of each replacement.
        var logger = new ReplacementLogger();

        // Define culture‑specific patterns and their replacements.
        var replacements = new List<(Regex Pattern, string Replacement)>
        {
            // British English to American English.
            (new Regex(@"\bcolour\b", RegexOptions.IgnoreCase), "color"),
            (new Regex(@"\borganisation\b", RegexOptions.IgnoreCase), "organization"),

            // French: remove accents for compatibility.
            (new Regex(@"\bprogramme\b", RegexOptions.IgnoreCase), "program"),
            (new Regex(@"\bdépartement\b", RegexOptions.IgnoreCase), "departement"),

            // German: replace ß with ss.
            (new Regex(@"\bStraße\b", RegexOptions.IgnoreCase), "Strasse")
        };

        int totalReplacements = 0;

        // Apply each replacement using the same FindReplaceOptions (with the logger).
        foreach (var (pattern, replacement) in replacements)
        {
            var options = new FindReplaceOptions
            {
                MatchCase = false,
                ReplacingCallback = logger
            };

            int count = doc.Range.Replace(pattern, replacement, options);
            totalReplacements += count;
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        doc.Save("updated.docx");

        // Serialize the replacement log to JSON.
        string jsonReport = JsonConvert.SerializeObject(logger.Replacements, Formatting.Indented);
        File.WriteAllText("replacement_report.json", jsonReport);
    }

    // Holds information about a single replacement operation.
    private class ReplacementInfo
    {
        public string Original { get; set; } = string.Empty;
        public string Replacement { get; set; } = string.Empty;
        public int Offset { get; set; }
    }

    // Callback that records each replacement made by the Range.Replace method.
    private class ReplacementLogger : IReplacingCallback
    {
        public List<ReplacementInfo> Replacements { get; } = new List<ReplacementInfo>();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Record the original matched text, the replacement text, and the offset within the node.
            Replacements.Add(new ReplacementInfo
            {
                Original = args.Match.Value,
                Replacement = args.Replacement,
                Offset = args.MatchOffset
            });

            // Proceed with the default replacement.
            return ReplaceAction.Replace;
        }
    }
}
