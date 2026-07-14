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
        builder.Writeln("The meeting will start at 10 a.m., e.g., after the briefing. Please bring your ID, i.e., a passport.");
        builder.Writeln("Remember to bring snacks, etc., for the break.");

        // Regex that matches the abbreviations and captures any following punctuation.
        // The trailing word‑boundary is removed so that a comma or other punctuation
        // immediately after the abbreviation is still part of the match.
        Regex abbreviationRegex = new Regex(@"\b(e\.g\.|i\.e\.|etc\.)(?<punctuation>[\.,;:]?)",
                                            RegexOptions.IgnoreCase);

        // Set up find‑and‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new AbbreviationReplacer();

        // Perform the replace operation. The replacement string is ignored because the callback supplies the value.
        int replacedCount = doc.Range.Replace(abbreviationRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No abbreviations were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output the resulting text to the console for verification.
        Console.WriteLine("Replacements performed: " + replacedCount);
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(doc.GetText().Trim());
    }

    // Callback that replaces each abbreviation with its full form while preserving captured punctuation.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> FullForms = new(StringComparer.OrdinalIgnoreCase)
        {
            { "e.g.", "for example" },
            { "i.e.", "that is" },
            { "etc.", "and so on" }
        };

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Determine the full form based on the matched abbreviation.
            string abbreviation = args.Match.Value;
            if (!FullForms.TryGetValue(abbreviation, out string fullForm))
                fullForm = abbreviation; // Fallback: keep original if not found.

            // Retrieve any trailing punctuation captured by the regex.
            string punctuation = args.Match.Groups["punctuation"].Value;

            // Set the replacement text.
            args.Replacement = fullForm + punctuation;

            return ReplaceAction.Replace;
        }
    }
}
