using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with abbreviations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dr. Smith went to the St. Mary's hospital.");
        builder.Writeln("Mr. Jones arrived, and Dr Brown stayed.");
        builder.Writeln("The St is old.");
        // Save the input document (optional, just for reference).
        doc.Save("input.docx");

        // Define abbreviation mappings.
        var abbreviationMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Dr", "Doctor" },
            { "Mr", "Mister" },
            { "St", "Saint" }
        };

        // Regular expression to match abbreviations with optional trailing period.
        // Group 1: abbreviation, Group 2: optional period.
        Regex regex = new Regex(@"\b(Dr|Mr|St)\b(\.)?", RegexOptions.IgnoreCase);

        // Set up find-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new AbbreviationReplacingCallback(abbreviationMap);

        // Perform the replacement.
        int replacedCount = doc.Range.Replace(regex, "", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one abbreviation replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces each abbreviation with its full form while preserving punctuation.
    private class AbbreviationReplacingCallback : IReplacingCallback
    {
        private readonly Dictionary<string, string> _map;

        public AbbreviationReplacingCallback(Dictionary<string, string> map)
        {
            _map = map ?? throw new ArgumentNullException(nameof(map));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ensure the match is not null.
            if (args.Match == null)
                return ReplaceAction.Skip;

            // Extract abbreviation and punctuation.
            string abbreviation = args.Match.Groups[1].Value;
            string punctuation = args.Match.Groups[2].Value; // May be empty.

            // Look up the full form; if not found, keep original.
            if (!_map.TryGetValue(abbreviation, out string fullForm))
                fullForm = abbreviation;

            // Set the replacement text, preserving any trailing punctuation.
            args.Replacement = fullForm + punctuation;
            return ReplaceAction.Replace;
        }
    }
}
