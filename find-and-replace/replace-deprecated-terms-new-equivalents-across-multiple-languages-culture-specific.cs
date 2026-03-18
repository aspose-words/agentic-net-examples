using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Define deprecated‑to‑new term mappings per culture (language).
        var replacements = new Dictionary<string, Dictionary<string, string>>
        {
            // English (United States)
            { "en-US", new Dictionary<string, string>
                {
                    { "colour", "color" },
                    { "organisation", "organization" },
                    { "centre", "center" }
                }
            },
            // German (Germany)
            { "de-DE", new Dictionary<string, string>
                {
                    { "Farbe", "Farbton" },
                    { "Zentrum", "Mitte" }
                }
            },
            // French (France)
            { "fr-FR", new Dictionary<string, string>
                {
                    { "adresse e‑mail", "email" },
                    { "centre", "centre" } // example where term stays same
                }
            }
        };

        // Iterate over each culture and perform replacements.
        foreach (var cultureEntry in replacements)
        {
            // CultureInfo is used only for potential future culture‑specific handling.
            CultureInfo culture = new CultureInfo(cultureEntry.Key);
            var termMap = cultureEntry.Value;

            // Build a regex that matches any deprecated term, respecting word boundaries.
            string pattern = @"\b(" + string.Join("|", EscapeKeys(termMap.Keys)) + @")\b";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            // Configure FindReplaceOptions with a callback that supplies the correct replacement.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new TermReplacingCallback(termMap);

            // Execute the replace on the whole document range.
            doc.Range.Replace(regex, string.Empty, options);
        }

        // Save the updated document.
        doc.Save("Output.docx");
    }

    // Escapes regex meta‑characters in each key.
    static IEnumerable<string> EscapeKeys(IEnumerable<string> keys)
    {
        foreach (var key in keys)
            yield return Regex.Escape(key);
    }

    // Callback that sets the replacement text based on the matched term.
    class TermReplacingCallback : IReplacingCallback
    {
        private readonly Dictionary<string, string> _termMap;

        public TermReplacingCallback(Dictionary<string, string> termMap)
        {
            _termMap = termMap;
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            string matched = args.Match.Value;

            // Find the replacement ignoring case.
            foreach (var kvp in _termMap)
            {
                if (string.Equals(kvp.Key, matched, StringComparison.OrdinalIgnoreCase))
                {
                    args.Replacement = kvp.Value;
                    break;
                }
            }

            return ReplaceAction.Replace;
        }
    }
}
