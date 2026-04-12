using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample text containing language code tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Welcome in [en] language.");
        builder.Writeln("Bonjour en [fr]!");
        builder.Writeln("Hola en [es]!");
        builder.Writeln("Guten Tag in [de]!");

        // Mapping from language codes to full language names.
        var languageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "en", "English" },
            { "fr", "French" },
            { "es", "Spanish" },
            { "de", "German" }
        };

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new LanguageTagReplacer(languageMap));

        // Regular expression to find tags like [en], [fr], etc.
        Regex tagPattern = new Regex(@"\[(\w{2})\]");

        // Perform the replacement. The replacement string is ignored because the callback supplies the actual text.
        int replacementCount = doc.Range.Replace(tagPattern, string.Empty, options);

        // Validate that at least one replacement occurred.
        if (replacementCount == 0)
            throw new InvalidOperationException("No language code tags were replaced.");

        // Save the modified document.
        const string outputPath = "LanguageTagReplaced.docx";
        doc.Save(outputPath);

        // Inform the user (console output is allowed as it does not require interaction).
        Console.WriteLine($"Replacements performed: {replacementCount}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that replaces a matched language code with its full name.
    private class LanguageTagReplacer : IReplacingCallback
    {
        private readonly Dictionary<string, string> _map;

        public LanguageTagReplacer(Dictionary<string, string> map)
        {
            _map = map ?? throw new ArgumentNullException(nameof(map));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The language code is captured in group 1 of the regex.
            string code = args.Match.Groups[1].Value;
            if (_map.TryGetValue(code, out string fullName))
                args.Replacement = fullName;
            else
                args.Replacement = code; // Fallback: keep the code if not found.

            return ReplaceAction.Replace;
        }
    }
}
