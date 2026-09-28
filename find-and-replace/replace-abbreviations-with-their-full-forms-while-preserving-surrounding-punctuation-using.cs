using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing abbreviations with various punctuation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example, e.g., of usage. Another case i.e.; something etc.");
        builder.Writeln("Multiple abbreviations: e.g., i.e., etc., and also e.g; i.e: etc:");
        builder.Writeln("No abbreviation here.");

        // Regex that matches the abbreviations and captures any following punctuation.
        Regex regex = new Regex(@"\b(e\.g|i\.e|etc)\b(?<punctuation>[.,;:]?)", RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replacement. The replacement string argument is ignored when a callback is used.
        int replacedCount = doc.Range.Replace(regex, string.Empty, options);

        // Verify that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one abbreviation replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces each abbreviation with its full form while preserving captured punctuation.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> FullForms = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "e.g", "for example" },
            { "i.e", "that is" },
            { "etc", "and so on" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the matched abbreviation (group 1) and any trailing punctuation.
            Match match = args.Match;
            string abbreviation = match.Groups[1].Value;
            string punctuation = match.Groups["punctuation"].Value;

            // Determine the full form; fall back to the original abbreviation if not found.
            string fullForm = FullForms.TryGetValue(abbreviation, out string value) ? value : abbreviation;

            // Build the replacement text, preserving the original punctuation.
            string replacementText = fullForm + punctuation;

            // Assign the replacement text. Using a string is sufficient for this scenario.
            args.Replacement = replacementText;

            return ReplaceAction.Replace;
        }
    }
}
