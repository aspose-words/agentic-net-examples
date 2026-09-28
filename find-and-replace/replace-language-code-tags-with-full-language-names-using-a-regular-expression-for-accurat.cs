using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome to our site.");
        builder.Writeln("[en] Hello!");
        builder.Writeln("[fr] Bonjour!");
        builder.Writeln("[es] Hola!");
        builder.Writeln("[de] Guten Tag!");

        // Save the input document (optional, demonstrates file creation).
        doc.Save("input.docx");

        // Mapping from language codes to full language names.
        var languageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en", "English" },
            { "fr", "French" },
            { "es", "Spanish" },
            { "de", "German" }
        };

        // Define a callback that replaces each tag with the full language name.
        var callback = new LanguageTagReplacer(languageMap);

        // Set up find‑replace options with the callback.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = callback,
            // Ensure the search is case‑insensitive.
            MatchCase = false
        };

        // Regular expression to match tags like [en], [fr], etc.
        Regex regex = new Regex(@"\[(?<code>[a-z]{2})\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Perform the replacement. The replacement string is ignored because the callback provides the value.
        int replacedCount = doc.Range.Replace(regex, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No language tags were replaced.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback implementation for custom replacement logic.
    private class LanguageTagReplacer : IReplacingCallback
    {
        private readonly IDictionary<string, string> _languageMap;

        public LanguageTagReplacer(IDictionary<string, string> languageMap)
        {
            _languageMap = languageMap ?? throw new ArgumentNullException(nameof(languageMap));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Extract the language code from the match.
            string code = args.Match.Groups["code"].Value;

            // Look up the full language name; if not found, keep the original tag.
            if (_languageMap.TryGetValue(code, out string fullName))
            {
                args.Replacement = fullName;
            }
            else
            {
                args.Replacement = args.Match.Value; // fallback to original tag
            }

            return ReplaceAction.Replace;
        }
    }
}
