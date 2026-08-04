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
        builder.Writeln("This is an example, e.g., of abbreviations i.e., usage etc.");
        builder.Writeln("Another line with e.g. and i.e. and etc.");

        // Define a regex that captures the abbreviation and any following punctuation.
        Regex abbreviationRegex = new Regex(@"\b(?<abbr>e\.g|i\.e|etc)\b(?<punct>[.,;:]?)", RegexOptions.IgnoreCase);

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new AbbreviationReplacer()
        };

        // Perform the replacement. The replacement string is ignored because the callback sets it.
        int replacedCount = doc.Range.Replace(abbreviationRegex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No abbreviations were replaced.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Callback that replaces each abbreviation with its full form while preserving punctuation.
    private class AbbreviationReplacer : IReplacingCallback
    {
        private static readonly Dictionary<string, string> FullForms = new(StringComparer.OrdinalIgnoreCase)
        {
            { "e.g", "for example" },
            { "i.e", "that is" },
            { "etc", "and so on" }
        };

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Retrieve captured groups.
            string abbr = args.Match.Groups["abbr"].Value;
            string punct = args.Match.Groups["punct"].Value;

            // Determine the full form; if not found, keep original.
            if (!FullForms.TryGetValue(abbr, out string fullForm))
                fullForm = abbr;

            // Set the replacement text, preserving any trailing punctuation.
            args.Replacement = fullForm + punct;
            return ReplaceAction.Replace;
        }
    }
}
