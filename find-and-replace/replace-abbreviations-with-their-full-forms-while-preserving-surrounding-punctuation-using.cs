using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add sample text containing abbreviations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("In this document we will see examples such as e.g., i.e., and etc., within sentences.");
        builder.Writeln("Sometimes abbreviations appear in parentheses (e.g.) or with commas, i.e., like this.");

        // Regex that matches the abbreviations and looks ahead for optional punctuation.
        // The trailing \b was removed because a period followed by another punctuation character
        // is not considered a word boundary, causing the original pattern to miss matches.
        Regex abbreviationRegex = new Regex(@"\b(e\.g\.|i\.e\.|etc\.)(?=[\s,;:\)\.]?)",
                                            RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback that maps each abbreviation to its full form.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // Ensure case‑insensitive matching (default is false, but set explicitly for clarity).
            MatchCase = false,
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replacement. The replacement string parameter is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(abbreviationRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No abbreviations were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Simple confirmation.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {outputPath}");
    }

    // Callback that replaces each matched abbreviation with its full form.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> _map = new(StringComparer.OrdinalIgnoreCase)
        {
            { "e.g.", "for example" },
            { "i.e.", "that is" },
            { "etc.", "and so on" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match value may include trailing punctuation (e.g., "e.g.") without the following comma.
            // Use the map to find the appropriate replacement.
            if (_map.TryGetValue(args.Match.Value, out string replacement))
            {
                args.Replacement = replacement;
                return ReplaceAction.Replace;
            }

            // If the match is not in the map, skip it.
            return ReplaceAction.Skip;
        }
    }
}
