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
        builder.Writeln("Dr. Smith visited St. Mary's, and Mr. Johnson stayed at the St. George's Hospital.");
        builder.Writeln("The meeting was scheduled for 10 a.m., but Dr, Jones arrived late.");
        // Save the input document (optional, for inspection).
        doc.Save("input.docx");

        // Define a regex that matches the abbreviations and captures any following punctuation.
        Regex abbreviationRegex = new Regex(@"\b(Dr|St)\b(?<punct>[.,;:]?)", RegexOptions.IgnoreCase);

        // Set up find-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replacement. The replacement string parameter is ignored when a callback is used.
        int replacedCount = doc.Range.Replace(abbreviationRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No abbreviations were replaced.");

        // Save the modified document.
        doc.Save("output.docx");

        // Optionally, write a simple console message.
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }

    // Callback that replaces each abbreviation with its full form while preserving punctuation.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Dr", "Doctor" },
            { "St", "Saint" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the abbreviation (group 1) and the captured punctuation (named group "punct").
            string abbreviation = args.Match.Groups[1].Value;
            string punctuation = args.Match.Groups["punct"].Value;

            // Determine the full form; if not found, keep the original text.
            if (!_map.TryGetValue(abbreviation, out string fullForm))
                fullForm = abbreviation;

            // Set the replacement text, preserving the original punctuation.
            args.Replacement = fullForm + punctuation;

            return ReplaceAction.Replace;
        }
    }
}
