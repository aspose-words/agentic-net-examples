using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample paragraphs containing deprecated terms.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The term foobar is deprecated.");
        builder.Writeln("Le terme ancien est obsolète.");
        builder.Writeln("Der Begriff veraltet ist veraltet.");

        // Prepare culture‑specific patterns and their replacements.
        var replacements = new (Regex pattern, string replacement)[]
        {
            // English
            (new Regex(@"\bfoobar\b", RegexOptions.IgnoreCase), "fooBar"),
            // French
            (new Regex(@"\bancien\b", RegexOptions.IgnoreCase), "nouveau"),
            // German
            (new Regex(@"\bveraltet\b", RegexOptions.IgnoreCase), "aktuell")
        };

        // Logger that records each replacement performed.
        var logger = new ReplacementLogger();

        int totalReplacements = 0;
        foreach (var (pattern, replacement) in replacements)
        {
            var options = new FindReplaceOptions
            {
                MatchCase = false,
                ReplacingCallback = logger
            };

            totalReplacements += doc.Range.Replace(pattern, replacement, options);
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputDoc = "ReplacedDocument.docx";
        doc.Save(outputDoc);

        // Write a simple log file with details of the replacements.
        const string logFile = "ReplacementLog.txt";
        File.WriteAllText(logFile, logger.GetLog());

        // Inform the user via console (no input required).
        Console.WriteLine($"Replacements performed: {totalReplacements}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputDoc)}");
        Console.WriteLine($"Log saved to: {Path.GetFullPath(logFile)}");
    }

    // Implements IReplacingCallback to capture details of each match.
    private class ReplacementLogger : IReplacingCallback
    {
        private readonly StringBuilder _log = new StringBuilder();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            _log.AppendLine($"Matched \"{args.Match.Value}\" replaced with \"{args.Replacement}\" at offset {args.MatchOffset} in node type {args.MatchNode.NodeType}.");
            // Do not modify the replacement; just record it.
            return ReplaceAction.Replace;
        }

        public string GetLog() => _log.ToString();
    }
}
