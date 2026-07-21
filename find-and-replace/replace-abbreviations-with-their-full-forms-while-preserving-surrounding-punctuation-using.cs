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
        builder.Writeln("This is an example, e.g., of abbreviations i.e., used in text etc.");
        builder.Writeln("Another line with e.g; and i.e: and etc!");

        // Define a regex that matches the abbreviations and captures any following punctuation.
        Regex abbreviationRegex = new Regex(@"\b(e\.g|i\.e|etc)\b(?<punct>[\.,;:]?)",
                                            RegexOptions.IgnoreCase);

        // Set up find‑replace options with a custom callback to map each abbreviation to its full form.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replacement. The replacement string is ignored because the callback supplies it.
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

    // Callback that replaces each matched abbreviation with its full form,
    // preserving any trailing punctuation captured by the regex.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> FullFormMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { "e.g", "for example" },
            { "i.e", "that is" },
            { "etc", "and so on" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Group 1 contains the abbreviation without punctuation.
            string abbreviation = args.Match.Groups[1].Value;
            // Named group "punct" contains any trailing punctuation.
            string punctuation = args.Match.Groups["punct"].Value;

            if (FullFormMap.TryGetValue(abbreviation, out string fullForm))
            {
                args.Replacement = fullForm + punctuation;
            }
            else
            {
                // Fallback: keep the original text unchanged.
                args.Replacement = args.Match.Value;
            }

            return ReplaceAction.Replace;
        }
    }
}
