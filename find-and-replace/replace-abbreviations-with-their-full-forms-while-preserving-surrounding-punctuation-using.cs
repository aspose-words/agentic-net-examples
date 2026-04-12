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
        // Create a blank document and add sample text containing abbreviations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The quick brown fox jumps over the lazy dog, e.g., when it is tired. " +
                        "Sometimes, i.e., when it is sleepy, it rests. " +
                        "Various items are listed, etc.");

        // Regular expression that matches the abbreviations and captures any following punctuation.
        Regex abbreviationRegex = new Regex(@"\b(e\.g|i\.e|etc)\b(?<punctuation>[.,;:]?)",
                                            RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replace operation. The callback supplies the actual replacement text.
        int replacements = doc.Range.Replace(abbreviationRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacements == 0)
            throw new InvalidOperationException("No abbreviations were replaced.");

        // Save the modified document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AbbreviationReplaced.docx");
        doc.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Replacements made: {replacements}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that replaces each matched abbreviation with its full form,
    // preserving any trailing punctuation captured by the regex.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> _fullForms = new(StringComparer.OrdinalIgnoreCase)
        {
            { "e.g", "for example" },
            { "i.e", "that is" },
            { "etc", "and so on" }
        };

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Determine the abbreviation matched (case‑insensitive lookup).
            string abbreviation = args.Match.Value;
            if (!_fullForms.TryGetValue(abbreviation, out string? fullForm))
                fullForm = abbreviation; // Fallback: keep original if not found.

            // Retrieve any punctuation captured by the named group.
            string punctuation = args.Match.Groups["punctuation"].Value;

            // Set the replacement text.
            args.Replacement = fullForm + punctuation;
            return ReplaceAction.Replace;
        }
    }
}
